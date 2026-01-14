import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { startWorkbookSync } from "../workbookSync";

async function flushMicrotasks(times = 4): Promise<void> {
  for (let i = 0; i < times; i++) {
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
}

async function flushNextTick(): Promise<void> {
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
  await flushMicrotasks();
}

function createMaterializedDocument(): DocumentController {
  const document = new DocumentController();
  // DocumentController lazily materializes sheets. workbookSync captures an initial sheet snapshot
  // when it starts, so ensure at least the default sheet exists before wiring up sync.
  document.getCell("Sheet1", { row: 0, col: 0 });
  return document;
}

describe("workbookSync", () => {
  const originalTauri = (globalThis as any).__TAURI__;

  beforeEach(() => {
    const invoke = vi.fn().mockResolvedValue(null);
    (globalThis as any).__TAURI__ = { core: { invoke } };
  });

  afterEach(() => {
    (globalThis as any).__TAURI__ = originalTauri;
    vi.restoreAllMocks();
  });

  it("batches consecutive change events into a single set_range call", async () => {
    const document = createMaterializedDocument();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    document.setCellValue("Sheet1", { row: 0, col: 0 }, 1);
    document.setCellValue("Sheet1", { row: 0, col: 1 }, 2);

    expect(invoke).not.toHaveBeenCalled();

    await flushMicrotasks();

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("set_range", {
      sheet_id: "Sheet1",
      start_row: 0,
      start_col: 0,
      end_row: 0,
      end_col: 1,
      values: [
        [
          { value: 1, formula: null },
          { value: 2, formula: null }
        ]
      ]
    });

    sync.stop();
  });

  it("maps formulas and literals into the value/formula shape expected by the backend", async () => {
    const document = createMaterializedDocument();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    // DocumentController.setCellFormula stores formulas without "="; we must send canonical "=..."
    document.setCellFormula("Sheet1", { row: 0, col: 0 }, "SUM(1,2)");
    document.setCellValue("Sheet1", { row: 0, col: 1 }, 123);

    await flushMicrotasks();

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("set_range", {
      sheet_id: "Sheet1",
      start_row: 0,
      start_col: 0,
      end_row: 0,
      end_col: 1,
      values: [
        [
          { value: null, formula: "=SUM(1,2)" },
          { value: 123, formula: null }
        ]
      ]
    });

    sync.stop();
  });

  it("syncs deleteSheet via delete_sheet (without mirroring sparse cell clears)", async () => {
    const document = new DocumentController();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    // DocumentController creates sheets lazily; ensure Sheet1 exists so we can delete Sheet2
    // without tripping the "Cannot delete the last sheet" guard.
    document.getCell("Sheet1", { row: 0, col: 0 });

    // Create a second sheet with a value so deletion would normally produce a cell-clear delta.
    document.setCellValue("Sheet2", { row: 0, col: 0 }, "hello");
    await flushMicrotasks();

    invoke.mockClear();

    // Deleting Sheet2 emits per-cell deltas for the removed sheet. Those must be skipped by
    // the workbook sync bridge because we persist deletions via the dedicated `delete_sheet`
    // command (mirroring sparse clears is expensive and can race with deletion).
    document.deleteSheet("Sheet2");
    await flushMicrotasks();

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("delete_sheet", { sheet_id: "Sheet2" });

    sync.stop();
  });

  it("ignores document changes tagged as macro updates (already applied in the backend)", async () => {
    const document = createMaterializedDocument();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    const before = document.getCell("Sheet1", { row: 0, col: 0 });
    document.applyExternalDeltas(
      [
        {
          sheetId: "Sheet1",
          row: 0,
          col: 0,
          before,
          after: { value: 1, formula: null, styleId: before.styleId },
        },
      ],
      { source: "macro" },
    );

    expect(document.getCell("Sheet1", { row: 0, col: 0 }).value).toBe(1);

    await flushMicrotasks();

    expect(invoke).not.toHaveBeenCalled();

    sync.stop();
  });

  it("ignores document changes tagged as native python updates (already applied in the backend)", async () => {
    const document = createMaterializedDocument();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    const before = document.getCell("Sheet1", { row: 0, col: 0 });
    document.applyExternalDeltas(
      [
        {
          sheetId: "Sheet1",
          row: 0,
          col: 0,
          before,
          after: { value: 1, formula: null, styleId: before.styleId },
        },
      ],
      { source: "python" },
    );

    expect(document.getCell("Sheet1", { row: 0, col: 0 }).value).toBe(1);

    await flushMicrotasks();

    expect(invoke).not.toHaveBeenCalled();

    sync.stop();
  });

  it("applies backend updates returned from set_range (e.g. pivot output auto-refresh)", async () => {
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      if (cmd === "set_range") {
        const raw = args?.values?.[0]?.[0]?.value ?? null;
        const pivotValue = raw == null ? 0 : raw;
        return [
          {
            sheet_id: "Pivot",
            row: 0,
            col: 0,
            value: pivotValue,
            formula: null,
            display_value: String(pivotValue),
          },
        ];
      }
      return null;
    });
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const document = createMaterializedDocument();
    const sync = startWorkbookSync({ document: document as any });

    document.setCellValue("Sheet1", { row: 0, col: 0 }, 1);
    await flushMicrotasks();

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(document.getCell("Pivot", { row: 0, col: 0 }).value).toBe(1);
    expect(document.isDirty).toBe(true);

    expect(document.undo()).toBe(true);
    expect(document.isDirty).toBe(false);

    await flushMicrotasks(8);
    expect(document.getCell("Pivot", { row: 0, col: 0 }).value).toBe(0);
    expect(document.isDirty).toBe(false);

    sync.stop();
  });

  it("ignores backend updates that reference deleted sheets (no resurrection)", async () => {
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      if (cmd === "set_range") {
        return [
          {
            sheet_id: "Sheet2",
            row: 0,
            col: 0,
            value: "stale",
            formula: null,
            display_value: "stale",
          },
        ];
      }
      return null;
    });
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const document = new DocumentController();
    // Ensure Sheet1 exists so we can delete Sheet2 without tripping the last-sheet guard.
    document.getCell("Sheet1", { row: 0, col: 0 });
    document.setCellValue("Sheet2", { row: 0, col: 0 }, "two");
    expect(document.getSheetIds()).toEqual(["Sheet1", "Sheet2"]);

    const sync = startWorkbookSync({ document: document as any });

    // Delete Sheet2.
    document.deleteSheet("Sheet2");
    expect(document.getSheetIds()).toEqual(["Sheet1"]);

    // Trigger a backend-synced edit; backend (incorrectly) returns an update for the deleted sheet id.
    document.setCellValue("Sheet1", { row: 0, col: 0 }, "one");
    await flushMicrotasks();

    // The stale backend update should not recreate Sheet2.
    expect(document.getSheetIds()).toEqual(["Sheet1"]);

    sync.stop();
  });

  it("ignores document changes tagged as pivot updates (already applied in the backend)", async () => {
    const document = createMaterializedDocument();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    document.setCellValue("Sheet1", { row: 0, col: 0 }, 123, { source: "pivot" });
    expect(document.getCell("Sheet1", { row: 0, col: 0 }).value).toBe(123);

    await flushMicrotasks();

    expect(invoke).not.toHaveBeenCalled();

    sync.stop();
  });

  it("markSaved flushes pending edits before saving and clears frontend dirty state", async () => {
    const document = createMaterializedDocument();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    document.setCellValue("Sheet1", { row: 0, col: 0 }, "hello");
    expect(document.isDirty).toBe(true);

    // markSaved should force a flush even if the microtask batch hasn't executed yet.
    await sync.markSaved();

    const cmds = invoke.mock.calls.map((c) => c[0]);
    expect(cmds[cmds.length - 1]).toBe("save_workbook");
    expect(cmds[0]).toMatch(/set_(cell|range)/);
    expect(document.isDirty).toBe(false);

    sync.stop();
  });

  it("style-only delta triggers apply_sheet_formatting_deltas and does not call set_cell/set_range", async () => {
    const document = new DocumentController();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    document.setRangeFormat("Sheet1", "A1", { font: { bold: true } });

    await flushMicrotasks();

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("apply_sheet_formatting_deltas", {
      payload: {
        sheetId: "Sheet1",
        cellFormats: [
          {
            row: 0,
            col: 0,
            format: { font: { bold: true } },
          },
        ],
      },
    });

    const cmds = invoke.mock.calls.map((c) => c[0]);
    expect(cmds).not.toContain("set_cell");
    expect(cmds).not.toContain("set_range");

    sync.stop();
  });

  it("syncs sheet view deltas (colWidths/rowHeights) via apply_sheet_view_deltas and restores via applyState", async () => {
    const document = createMaterializedDocument();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    document.setColWidth("Sheet1", 1, 120);
    document.setRowHeight("Sheet1", 3, 44);

    await flushMicrotasks();

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith("apply_sheet_view_deltas", {
      payload: {
        sheetId: "Sheet1",
        colWidths: [{ col: 1, width: 120 }],
        rowHeights: [{ row: 3, height: 44 }],
      },
    });

    const payload = invoke.mock.calls[0]?.[1]?.payload as any;
    const persistedColWidths: Record<string, number> = {};
    for (const d of payload.colWidths ?? []) {
      if (typeof d?.col === "number" && typeof d?.width === "number") {
        persistedColWidths[String(d.col)] = d.width;
      }
    }
    const persistedRowHeights: Record<string, number> = {};
    for (const d of payload.rowHeights ?? []) {
      if (typeof d?.row === "number" && typeof d?.height === "number") {
        persistedRowHeights[String(d.row)] = d.height;
      }
    }

    const snapshot = new TextEncoder().encode(
      JSON.stringify({
        schemaVersion: 1,
        sheets: [
          {
            id: "Sheet1",
            name: "Sheet1",
            visibility: "visible",
            frozenRows: 0,
            frozenCols: 0,
            cells: [],
            colWidths: persistedColWidths,
            rowHeights: persistedRowHeights,
          },
        ],
      }),
    );

    const restored = new DocumentController();
    restored.applyState(snapshot);
    expect(restored.getSheetView("Sheet1")).toMatchObject({
      colWidths: { "1": 120 },
      rowHeights: { "3": 44 },
    });

    sync.stop();
  });

  it("combined value+style delta triggers both set_range and apply_sheet_formatting_deltas", async () => {
    const document = new DocumentController();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    document.setCellValue("Sheet1", { row: 0, col: 0 }, 123);
    document.setRangeFormat("Sheet1", "A1", { fill: { color: "#ff0000" } });

    await flushMicrotasks();

    const cmds = invoke.mock.calls.map((c) => c[0]);
    expect(cmds).toEqual(["set_range", "apply_sheet_formatting_deltas"]);

    expect(invoke).toHaveBeenCalledWith("set_range", {
      sheet_id: "Sheet1",
      start_row: 0,
      start_col: 0,
      end_row: 0,
      end_col: 0,
      values: [[{ value: 123, formula: null }]],
    });

    expect(invoke).toHaveBeenCalledWith("apply_sheet_formatting_deltas", {
      payload: {
        sheetId: "Sheet1",
        cellFormats: [
          {
            row: 0,
            col: 0,
            format: { fill: { color: "#ff0000" } },
          },
        ],
      },
    });

    sync.stop();
  });

  it("row/col/sheet/range-run deltas are forwarded", async () => {
    const document = new DocumentController();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    document.setSheetFormat("Sheet1", { font: { italic: true } });
    document.setRowFormat("Sheet1", 1, { font: { bold: true } });
    document.setColFormat("Sheet1", 2, { font: { underline: true } });
    // Big enough to trigger range-run formatting deltas.
    document.setRangeFormat(
      "Sheet1",
      {
        start: { row: 0, col: 0 },
        end: { row: 10_000, col: 4 },
      },
      { fill: { color: "#00ff00" } },
    );

    await flushMicrotasks();

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke).toHaveBeenCalledWith(
      "apply_sheet_formatting_deltas",
      expect.objectContaining({
        payload: expect.objectContaining({
          sheetId: "Sheet1",
          defaultFormat: { font: { italic: true } },
          rowFormats: [
            {
              row: 1,
              format: { font: { bold: true } },
            },
          ],
          colFormats: [
            {
              col: 2,
              format: { font: { underline: true } },
            },
          ],
        }),
      }),
    );

    const payload = invoke.mock.calls[0]?.[1]?.payload as any;
    expect(payload.formatRunsByCol.length).toBeGreaterThan(0);
    expect(payload.formatRunsByCol[0]).toEqual(
      expect.objectContaining({
        col: 0,
        runs: [
          {
            startRow: 0,
            endRowExclusive: 10_001,
            format: { fill: { color: "#00ff00" } },
          },
        ],
      }),
    );

    const cmds = invoke.mock.calls.map((c) => c[0]);
    expect(cmds).not.toContain("set_cell");
    expect(cmds).not.toContain("set_range");

    sync.stop();
  });

  it("clears the backend dirty flag when undo returns the document to a saved state", async () => {
    const document = createMaterializedDocument();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    document.setCellValue("Sheet1", { row: 0, col: 0 }, "hello");
    await flushMicrotasks();
    expect(document.isDirty).toBe(true);

    expect(document.undo()).toBe(true);
    expect(document.isDirty).toBe(false);

    await flushMicrotasks(8);
    await flushNextTick();

    const cmds = invoke.mock.calls.map((c) => c[0]);
    expect(cmds[0]).toMatch(/set_(cell|range)/);
    expect(cmds[1]).toMatch(/set_(cell|range)/);
    expect(cmds[2]).toBe("mark_saved");

    sync.stop();
  });

  it("yields before mark_saved so delayed backend ops (e.g. sheet reorder) can enqueue first", async () => {
    const document = createMaterializedDocument();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    document.setCellValue("Sheet1", { row: 0, col: 0 }, "hello");
    await flushMicrotasks();
    expect(document.isDirty).toBe(true);

    expect(document.undo()).toBe(true);
    expect(document.isDirty).toBe(false);

    // Simulate a doc-driven sheet reorder persistence layer that schedules a backend call
    // after a microtask tick (like `main.ts`'s reorder coalescing logic).
    queueMicrotask(() => {
      void (async () => {
        await new Promise<void>((resolve) => queueMicrotask(resolve));
        await invoke("move_sheet", { sheet_id: "Sheet1", to_index: 0 });
      })();
    });

    await flushMicrotasks(8);
    await flushNextTick();

    const cmds = invoke.mock.calls.map((c) => c[0]);
    const moveIdx = cmds.lastIndexOf("move_sheet");
    const markIdx = cmds.lastIndexOf("mark_saved");
    expect(moveIdx).toBeGreaterThanOrEqual(0);
    expect(markIdx).toBeGreaterThan(moveIdx);

    sync.stop();
  });

  it("clears the backend dirty flag when undo returns the document to a saved state (sheet metadata only)", async () => {
    const document = createMaterializedDocument();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    document.renameSheet("Sheet1", "RenamedSheet1");
    expect(document.isDirty).toBe(true);

    await flushMicrotasks();
    expect(invoke).toHaveBeenCalledWith("rename_sheet", { sheet_id: "Sheet1", name: "RenamedSheet1" });

    expect(document.undo()).toBe(true);
    expect(document.isDirty).toBe(false);

    await flushMicrotasks(8);
    await flushNextTick();

    const cmds = invoke.mock.calls.map((c) => c[0]);
    expect(cmds[0]).toBe("rename_sheet");
    expect(cmds[1]).toBe("rename_sheet");
    expect(cmds[2]).toBe("mark_saved");

    sync.stop();
  });

  it("filters applyState deleted-sheet cell deltas (avoids per-cell clears for removed sheets)", async () => {
    const document = createMaterializedDocument();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    // Populate a second sheet with a value so `applyState` must delete it.
    document.setCellValue("Sheet2", { row: 0, col: 0 }, "hello");
    await flushMicrotasks();

    invoke.mockClear();

    // Build a snapshot that only contains Sheet1. `applyState` does not emit sheetMetaDeltas
    // for structural changes; workbookSync must instead filter cell deltas against the
    // post-applyState sheet snapshot.
    const snapshotDoc = new DocumentController();
    snapshotDoc.getCell("Sheet1", { row: 0, col: 0 });
    const snapshot = snapshotDoc.encodeState();

    document.applyState(snapshot);

    await flushMicrotasks(8);
    await flushNextTick();

    expect(document.getSheetIds()).toEqual(["Sheet1"]);

    // The important invariant: workbookSync must not mirror per-cell sparse clears for deleted sheets
    // (that would be extremely expensive and can race with deletion).
    const cmds = invoke.mock.calls.map((c) => c[0]);
    expect(cmds.some((cmd) => cmd === "set_cell" || cmd === "set_range")).toBe(false);

    // Depending on the backend/sync configuration, applyState may or may not be mirrored as a
    // sheet-level delete. If it is, it should use the dedicated delete_sheet command.
    if (cmds.length > 0) {
      expect(cmds).toEqual(["delete_sheet"]);
      expect(invoke).toHaveBeenCalledWith("delete_sheet", { sheet_id: "Sheet2" });
    }

    sync.stop();
  });
});
