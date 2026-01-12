import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { startWorkbookSync } from "../workbookSync";

async function flushMicrotasks(times = 4): Promise<void> {
  for (let i = 0; i < times; i++) {
    await new Promise<void>((resolve) => queueMicrotask(resolve));
  }
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
    const document = new DocumentController();
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
    const document = new DocumentController();
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

  it("ignores document changes tagged as macro updates (already applied in the backend)", async () => {
    const document = new DocumentController();
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
    const document = new DocumentController();
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

    const document = new DocumentController();
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

  it("ignores document changes tagged as pivot updates (already applied in the backend)", async () => {
    const document = new DocumentController();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    document.setCellValue("Sheet1", { row: 0, col: 0 }, 123, { source: "pivot" });
    expect(document.getCell("Sheet1", { row: 0, col: 0 }).value).toBe(123);

    await flushMicrotasks();

    expect(invoke).not.toHaveBeenCalled();

    sync.stop();
  });

  it("markSaved flushes pending edits before saving and clears frontend dirty state", async () => {
    const document = new DocumentController();
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

  it("clears the backend dirty flag when undo returns the document to a saved state", async () => {
    const document = new DocumentController();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    document.setCellValue("Sheet1", { row: 0, col: 0 }, "hello");
    await flushMicrotasks();
    expect(document.isDirty).toBe(true);

    expect(document.undo()).toBe(true);
    expect(document.isDirty).toBe(false);

    await flushMicrotasks(8);

    const cmds = invoke.mock.calls.map((c) => c[0]);
    expect(cmds[0]).toMatch(/set_(cell|range)/);
    expect(cmds[1]).toMatch(/set_(cell|range)/);
    expect(cmds[2]).toBe("mark_saved");

    sync.stop();
  });

  it("clears the backend dirty flag when undo returns the document to a saved state (sheet metadata only)", async () => {
    const document = new DocumentController();
    const sync = startWorkbookSync({ document: document as any });
    const invoke = (globalThis as any).__TAURI__?.core?.invoke as ReturnType<typeof vi.fn>;

    document.renameSheet("Sheet1", "RenamedSheet1");
    expect(document.isDirty).toBe(true);

    // Renames are persisted to the backend via main.ts in desktop mode, so workbookSync should
    // only mirror the undo/redo direction.
    await flushMicrotasks();
    expect(invoke).not.toHaveBeenCalled();

    expect(document.undo()).toBe(true);
    expect(document.isDirty).toBe(false);

    await flushMicrotasks(8);

    const cmds = invoke.mock.calls.map((c) => c[0]);
    expect(cmds[0]).toBe("rename_sheet");
    expect(cmds[1]).toBe("mark_saved");

    sync.stop();
  });
});
