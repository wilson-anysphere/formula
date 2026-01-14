/**
 * @vitest-environment jsdom
 */

import { afterEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { fireWorkbookBeforeCloseBestEffort, installVbaEventMacros } from "../event_macros";

type Range = { startRow: number; startCol: number; endRow: number; endCol: number };

class FakeApp {
  private sheetId = "Sheet1";
  private active = { row: 0, col: 0 };
  private selection: Range[] = [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }];
  private readonly selectionListeners = new Set<(state: { ranges: Range[] }) => void>();

  readonly refresh = vi.fn();
  readonly whenIdle = vi.fn(async () => undefined);

  constructor(private readonly document: DocumentController) {}

  getCurrentSheetId(): string {
    return this.sheetId;
  }

  getActiveCell(): { row: number; col: number } {
    return { ...this.active };
  }

  getSelectionRanges(): Range[] {
    return this.selection;
  }

  subscribeSelection(listener: (selection: { ranges: Range[] }) => void): () => void {
    this.selectionListeners.add(listener);
    listener({ ranges: this.selection });
    return () => this.selectionListeners.delete(listener);
  }

  emitSelection(next: Range[], opts: { sheetId?: string; active?: { row: number; col: number } } = {}): void {
    if (opts.sheetId) this.sheetId = opts.sheetId;
    if (opts.active) this.active = { ...opts.active };
    else {
      const first = next[0];
      if (first) this.active = { row: first.startRow, col: first.startCol };
    }
    this.selection = next;
    const payload = { ranges: next };
    for (const listener of this.selectionListeners) listener(payload);
  }

  getDocument(): DocumentController {
    return this.document;
  }
}

async function flushAsync(): Promise<void> {
  // Allow queued microtasks (queueMicrotask) and promise chains to settle.
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
}

async function flushMicrotasks(turns = 10): Promise<void> {
  for (let i = 0; i < turns; i++) {
    await Promise.resolve();
  }
}

describe("VBA event macros wiring", () => {
  afterEach(() => {
    vi.useRealTimers();
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    delete (globalThis as any).__TAURI__;
  });

  it("coalesces DocumentController change deltas into a single Worksheet_Change bounding rect", async () => {
    const calls: Array<{ cmd: string; args?: any }> = [];

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_worksheet_change") return { ok: true, output: [], updates: [] };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushAsync(); // let Workbook_Open settle
    calls.length = 0;

    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "A");
    doc.setCellValue("Sheet1", { row: 2, col: 3 }, "B");

    await flushAsync();

    const changeCalls = calls.filter((c) => c.cmd === "fire_worksheet_change");
    expect(changeCalls).toHaveLength(1);
    expect(changeCalls[0]?.args).toMatchObject({
      sheet_id: "Sheet1",
      start_row: 0,
      start_col: 0,
      end_row: 2,
      end_col: 3,
    });

    const setIdx = calls.findIndex((c) => c.cmd === "set_macro_ui_context");
    const changeIdx = calls.findIndex((c) => c.cmd === "fire_worksheet_change");
    expect(setIdx).toBeGreaterThanOrEqual(0);
    expect(changeIdx).toBeGreaterThan(setIdx);

    wiring.dispose();
  });

  it("captures Worksheet_Change edits that occur before macro security status resolves", async () => {
    const calls: Array<{ cmd: string; args?: any }> = [];

    let resolveStatus: ((value: any) => void) | null = null;
    const statusPromise = new Promise<any>((resolve) => {
      resolveStatus = resolve;
    });

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return await statusPromise;
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_worksheet_change") return { ok: true, output: [], updates: [] };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    // Ensure the status call has been issued.
    await Promise.resolve();

    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "A");
    doc.setCellValue("Sheet1", { row: 2, col: 3 }, "B");

    expect(calls.some((c) => c.cmd === "fire_worksheet_change")).toBe(false);

    (resolveStatus as ((value: any) => void) | null)?.({
      has_macros: true,
      origin_path: null,
      workbook_fingerprint: null,
      signature: null,
      trust: "trusted_always",
    });

    await flushAsync();
    await flushAsync();

    const changeCalls = calls.filter((c) => c.cmd === "fire_worksheet_change");
    expect(changeCalls).toHaveLength(1);
    expect(changeCalls[0]?.args).toMatchObject({
      sheet_id: "Sheet1",
      start_row: 0,
      start_col: 0,
      end_row: 2,
      end_col: 3,
    });

    wiring.dispose();
  });

  it("ignores backend-derived DocumentController updates (e.g. pivot auto-refresh output)", async () => {
    const calls: Array<{ cmd: string; args?: any }> = [];

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_worksheet_change") return { ok: true, output: [], updates: [] };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushAsync(); // let Workbook_Open settle
    calls.length = 0;

    const before = doc.getCell("Sheet1", { row: 0, col: 0 });
    doc.applyExternalDeltas(
      [
        {
          sheetId: "Sheet1",
          row: 0,
          col: 0,
          before,
          after: { value: 1, formula: null, styleId: before.styleId },
        },
      ],
      { source: "backend", markDirty: false },
    );

    await flushAsync();

    expect(calls.some((c) => c.cmd === "fire_worksheet_change")).toBe(false);

    wiring.dispose();
  });

  it("does not resurrect deleted sheets when applying Workbook_BeforeClose updates", async () => {
    const doc = new DocumentController();
    // Ensure Sheet1 exists so we can delete Sheet2 without tripping the last-sheet guard.
    doc.getCell("Sheet1", { row: 0, col: 0 });
    doc.setCellValue("Sheet2", { row: 0, col: 0 }, "two");
    expect(doc.getSheetIds()).toEqual(["Sheet1", "Sheet2"]);
    doc.deleteSheet("Sheet2");
    expect(doc.getSheetIds()).toEqual(["Sheet1"]);

    const app = new FakeApp(doc);

    const invoke = vi.fn(async (cmd: string) => {
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_before_close") {
        return {
          ok: true,
          output: [],
          updates: [{ sheet_id: "Sheet2", row: 0, col: 0, value: "stale", formula: null, display_value: "stale" }],
        };
      }
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    await fireWorkbookBeforeCloseBestEffort({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    expect(doc.getSheetIds()).toEqual(["Sheet1"]);
  });

  it("uses the UI context captured at edit time when firing Worksheet_Change", async () => {
    const calls: Array<{ cmd: string; args?: any }> = [];

    let resolveStatus: ((value: any) => void) | null = null;
    const statusPromise = new Promise<any>((resolve) => {
      resolveStatus = resolve;
    });

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return await statusPromise;
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_worksheet_change") return { ok: true, output: [], updates: [] };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);
    // Establish a non-default selection before wiring installs the selection listener (so the
    // initial selection emission is suppressed but Worksheet_Change captures this state).
    app.emitSelection([{ startRow: 1, startCol: 1, endRow: 1, endCol: 1 }], { sheetId: "Sheet1", active: { row: 1, col: 1 } });

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    // Ensure the status call has been issued.
    await Promise.resolve();

    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "A");

    // Simulate the user switching sheets before the delayed macro execution begins.
    (app as any).sheetId = "Sheet2";
    (app as any).active = { row: 9, col: 9 };
    (app as any).selection = [{ startRow: 9, startCol: 9, endRow: 9, endCol: 9 }];

    resolveStatus?.({
      has_macros: true,
      origin_path: null,
      workbook_fingerprint: null,
      signature: null,
      trust: "trusted_always",
    });

    await flushAsync();
    await flushAsync();

    const changeIdx = calls.findIndex((c) => c.cmd === "fire_worksheet_change");
    expect(changeIdx).toBeGreaterThanOrEqual(0);

    const uiIdx = calls
      .slice(0, changeIdx)
      .map((call, idx) => (call.cmd === "set_macro_ui_context" ? idx : -1))
      .filter((idx) => idx >= 0)
      .pop();
    expect(uiIdx).not.toBeUndefined();

    expect(calls[uiIdx!]?.args).toMatchObject({
      sheet_id: "Sheet1",
      active_row: 1,
      active_col: 1,
      selection: { start_row: 1, start_col: 1, end_row: 1, end_col: 1 },
    });

    wiring.dispose();
  });

  it("preserves captured UI context when a Worksheet_Change flush is deferred by another event macro", async () => {
    vi.useFakeTimers();

    const calls: Array<{ cmd: string; args?: any }> = [];

    let resolveSheet1Macro: (() => void) | null = null;
    const sheet1MacroPromise = new Promise<void>((resolve) => {
      resolveSheet1Macro = resolve;
    });

    let resolveSelectionMacro: (() => void) | null = null;
    const selectionMacroPromise = new Promise<void>((resolve) => {
      resolveSelectionMacro = resolve;
    });

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_selection_change") {
        await selectionMacroPromise;
        return { ok: true, output: [], updates: [] };
      }
      if (cmd === "fire_worksheet_change") {
        if (args?.sheet_id === "Sheet1") {
          await sheet1MacroPromise;
        }
        return { ok: true, output: [], updates: [] };
      }
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushMicrotasks();
    calls.length = 0;

    // Edit Sheet1 with Sheet1 active.
    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "A");

    // Switch to Sheet2 and edit it, capturing a different UI context per sheet.
    app.emitSelection([{ startRow: 5, startCol: 5, endRow: 5, endCol: 5 }], { sheetId: "Sheet2", active: { row: 5, col: 5 } });
    doc.setCellValue("Sheet2", { row: 1, col: 1 }, "B");

    // Change the visible UI state without emitting a selection event, so if Sheet2's context is
    // lost, the rerun will sync the macro runtime to the wrong sheet.
    (app as any).sheetId = "Sheet3";
    (app as any).active = { row: 9, col: 9 };
    (app as any).selection = [{ startRow: 9, startCol: 9, endRow: 9, endCol: 9 }];

    await flushMicrotasks();

    // Allow the debounced selection handler to fire while Sheet1's macro is still running, so the
    // SelectionChange macro starts as soon as Sheet1 finishes and overlaps the Sheet2 rerun.
    await vi.advanceTimersByTimeAsync(100);
    await flushMicrotasks();

    resolveSheet1Macro?.();
    await flushMicrotasks();

    // Let the selection macro remain in flight long enough for Sheet2's first attempt to defer.
    await flushMicrotasks();

    resolveSelectionMacro?.();
    await flushMicrotasks();
    await flushMicrotasks();

    const sheet2ChangeIdx = calls.findIndex((c) => c.cmd === "fire_worksheet_change" && c.args?.sheet_id === "Sheet2");
    expect(sheet2ChangeIdx).toBeGreaterThanOrEqual(0);

    const uiCall = calls[sheet2ChangeIdx - 1];
    expect(uiCall?.cmd).toBe("set_macro_ui_context");
    expect(uiCall?.args).toMatchObject({
      sheet_id: "Sheet2",
      active_row: 5,
      active_col: 5,
      selection: { start_row: 5, start_col: 5, end_row: 5, end_col: 5 },
    });

    wiring.dispose();
  });

  it("captures Worksheet_SelectionChange changes that occur before macro security status resolves", async () => {
    vi.useFakeTimers();

    const calls: Array<{ cmd: string; args?: any }> = [];

    let resolveStatus: ((value: any) => void) | null = null;
    const statusPromise = new Promise<any>((resolve) => {
      resolveStatus = resolve;
    });

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return await statusPromise;
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_selection_change") return { ok: true, output: [], updates: [] };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushMicrotasks();

    // Initial selection is suppressed.
    app.emitSelection([{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);
    app.emitSelection([{ startRow: 2, startCol: 3, endRow: 4, endCol: 5 }], { active: { row: 4, col: 5 } });

    await vi.advanceTimersByTimeAsync(100);

    expect(calls.some((c) => c.cmd === "fire_selection_change")).toBe(false);

    (resolveStatus as ((value: any) => void) | null)?.({
      has_macros: true,
      origin_path: null,
      workbook_fingerprint: null,
      signature: null,
      trust: "trusted_always",
    });

    await flushMicrotasks();

    const selectionCalls = calls.filter((c) => c.cmd === "fire_selection_change");
    expect(selectionCalls).toHaveLength(1);
    expect(selectionCalls[0]?.args).toMatchObject({
      sheet_id: "Sheet1",
      start_row: 2,
      start_col: 3,
      end_row: 4,
      end_col: 5,
    });

    wiring.dispose();
  });

  it("ignores format-only deltas when firing Worksheet_Change", async () => {
    const calls: Array<{ cmd: string; args?: any }> = [];

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_worksheet_change") return { ok: true, output: [], updates: [] };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushAsync(); // let Workbook_Open settle
    calls.length = 0;

    // Style-only edits should not trigger Worksheet_Change.
    doc.setRangeFormat("Sheet1", { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }, { numberFormat: "0.00" });

    await flushAsync();

    const changeCalls = calls.filter((c) => c.cmd === "fire_worksheet_change");
    expect(changeCalls).toHaveLength(0);

    wiring.dispose();
  });

  it("ignores applyState deltas when firing Worksheet_Change", async () => {
    const calls: Array<{ cmd: string; args?: any }> = [];

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_worksheet_change") return { ok: true, output: [], updates: [] };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushAsync(); // let Workbook_Open settle
    calls.length = 0;

    const snapshot = new TextEncoder().encode(
      JSON.stringify({
        schemaVersion: 1,
        sheets: [{ id: "Sheet1", cells: [{ row: 0, col: 0, value: "X", formula: null, format: null }] }],
      }),
    );
    doc.applyState(snapshot);

    await flushAsync();

    const changeCalls = calls.filter((c) => c.cmd === "fire_worksheet_change");
    expect(changeCalls).toHaveLength(0);

    wiring.dispose();
  });

  it("ignores undo/redo deltas when firing Worksheet_Change", async () => {
    const calls: Array<{ cmd: string; args?: any }> = [];

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_worksheet_change") return { ok: true, output: [], updates: [] };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushAsync(); // let Workbook_Open settle
    calls.length = 0;

    doc.setCellValue("Sheet1", { row: 0, col: 0 }, "X");
    await flushAsync();

    expect(calls.filter((c) => c.cmd === "fire_worksheet_change")).toHaveLength(1);

    calls.length = 0;
    expect(doc.undo()).toBe(true);
    await flushAsync();
    expect(calls.filter((c) => c.cmd === "fire_worksheet_change")).toHaveLength(0);

    calls.length = 0;
    expect(doc.redo()).toBe(true);
    await flushAsync();
    expect(calls.filter((c) => c.cmd === "fire_worksheet_change")).toHaveLength(0);

    wiring.dispose();
  });

  it("ignores collaboration deltas when firing Worksheet_Change", async () => {
    const calls: Array<{ cmd: string; args?: any }> = [];

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_worksheet_change") return { ok: true, output: [], updates: [] };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushAsync(); // let Workbook_Open settle
    calls.length = 0;

    const before = doc.getCell("Sheet1", { row: 0, col: 0 });
    doc.applyExternalDeltas(
      [
        {
          sheetId: "Sheet1",
          row: 0,
          col: 0,
          before,
          after: { value: "remote", formula: null, styleId: before.styleId },
        },
      ],
      { source: "collab" },
    );

    await flushAsync();

    expect(calls.filter((c) => c.cmd === "fire_worksheet_change")).toHaveLength(0);

    wiring.dispose();
  });

  it("does not re-trigger Worksheet_Change when applying macro updates", async () => {
    const calls: Array<{ cmd: string; args?: any }> = [];

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_worksheet_change") {
        return {
          ok: true,
          output: [],
          updates: [
            {
              sheet_id: "Sheet1",
              row: 5,
              col: 5,
              value: 123,
              formula: null,
              display_value: "123",
            },
          ],
        };
      }
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushAsync(); // let Workbook_Open settle
    calls.length = 0;

    doc.setCellValue("Sheet1", { row: 1, col: 1 }, "X");
    await flushAsync();

    const changeCalls = calls.filter((c) => c.cmd === "fire_worksheet_change");
    expect(changeCalls).toHaveLength(1);

    const updated = doc.getCell("Sheet1", { row: 5, col: 5 }) as { value: unknown; formula: string | null };
    expect(updated.formula).toBeNull();
    expect(updated.value).toBe(123);

    wiring.dispose();
  });

  it("debounces Worksheet_SelectionChange so drags coalesce into a single call", async () => {
    vi.useFakeTimers();

    const calls: Array<{ cmd: string; args?: any }> = [];
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_selection_change") return { ok: true, output: [], updates: [] };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushMicrotasks();
    calls.length = 0;

    // Simulate a drag-selection that updates rapidly.
    app.emitSelection([{ startRow: 0, startCol: 0, endRow: 0, endCol: 0 }]);
    vi.advanceTimersByTime(25);
    app.emitSelection([{ startRow: 1, startCol: 1, endRow: 1, endCol: 1 }]);
    vi.advanceTimersByTime(25);
    app.emitSelection([{ startRow: 2, startCol: 2, endRow: 3, endCol: 3 }]);

    await vi.advanceTimersByTimeAsync(100);
    await flushMicrotasks();

    const selectionCalls = calls.filter((c) => c.cmd === "fire_selection_change");
    expect(selectionCalls).toHaveLength(1);
    expect(selectionCalls[0]?.args).toMatchObject({
      sheet_id: "Sheet1",
      start_row: 2,
      start_col: 2,
      end_row: 3,
      end_col: 3,
    });

    const setIdx = calls.findIndex((c) => c.cmd === "set_macro_ui_context");
    const selIdx = calls.findIndex((c) => c.cmd === "fire_selection_change");
    expect(selIdx).toBeGreaterThan(setIdx);

    wiring.dispose();
  });

  it("fires Worksheet_SelectionChange when only the active cell changes within a selection", async () => {
    vi.useFakeTimers();

    const calls: Array<{ cmd: string; args?: any }> = [];
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_selection_change") return { ok: true, output: [], updates: [] };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushMicrotasks();
    calls.length = 0;

    // Select a 2x2 range.
    app.emitSelection([{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }]);
    // Move the active cell within that selection without changing the selection rect.
    app.emitSelection([{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }], { active: { row: 1, col: 0 } });

    await vi.advanceTimersByTimeAsync(100);
    await flushMicrotasks();

    const selectionCalls = calls.filter((c) => c.cmd === "fire_selection_change");
    expect(selectionCalls).toHaveLength(1);
    expect(selectionCalls[0]?.args).toMatchObject({
      sheet_id: "Sheet1",
      start_row: 0,
      start_col: 0,
      end_row: 1,
      end_col: 1,
    });

    wiring.dispose();
  });

  it("uses the selection range containing the active cell when firing Worksheet_SelectionChange", async () => {
    vi.useFakeTimers();

    const calls: Array<{ cmd: string; args?: any }> = [];
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_selection_change") return { ok: true, output: [], updates: [] };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushMicrotasks();
    calls.length = 0;

    // Multi-range selection where the active cell is in the second range.
    app.emitSelection(
      [
        { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
        { startRow: 2, startCol: 2, endRow: 3, endCol: 4 },
      ],
      { active: { row: 3, col: 4 } },
    );

    await vi.advanceTimersByTimeAsync(100);
    await flushMicrotasks();

    const uiIdx = calls.findIndex((c) => c.cmd === "set_macro_ui_context");
    expect(calls[uiIdx]?.args).toMatchObject({
      sheet_id: "Sheet1",
      active_row: 3,
      active_col: 4,
      selection: { start_row: 2, start_col: 2, end_row: 3, end_col: 4 },
    });

    const selectionCalls = calls.filter((c) => c.cmd === "fire_selection_change");
    expect(selectionCalls).toHaveLength(1);
    expect(selectionCalls[0]?.args).toMatchObject({
      sheet_id: "Sheet1",
      start_row: 2,
      start_col: 2,
      end_row: 3,
      end_col: 4,
    });

    wiring.dispose();
  });

  it("uses the sheet id captured at selection time when firing Worksheet_SelectionChange", async () => {
    vi.useFakeTimers();

    const calls: Array<{ cmd: string; args?: any }> = [];
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_selection_change") return { ok: true, output: [], updates: [] };
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushMicrotasks();
    calls.length = 0;

    // Fire a selection change on Sheet1.
    app.emitSelection([{ startRow: 2, startCol: 3, endRow: 4, endCol: 5 }], { sheetId: "Sheet1" });

    // Simulate a sheet switch occurring before the debounced handler fires (without another
    // selection event being delivered).
    (app as any).sheetId = "Sheet2";

    await vi.advanceTimersByTimeAsync(100);
    await flushMicrotasks();

    const uiIdx = calls.findIndex((c) => c.cmd === "set_macro_ui_context");
    expect(calls[uiIdx]?.args).toMatchObject({
      sheet_id: "Sheet1",
      active_row: 2,
      active_col: 3,
      selection: { start_row: 2, start_col: 3, end_row: 4, end_col: 5 },
    });

    const selectionCalls = calls.filter((c) => c.cmd === "fire_selection_change");
    expect(selectionCalls).toHaveLength(1);
    expect(selectionCalls[0]?.args).toMatchObject({
      sheet_id: "Sheet1",
      start_row: 2,
      start_col: 3,
      end_row: 4,
      end_col: 5,
    });

    wiring.dispose();
  });

  it("does not recursively fire Worksheet_SelectionChange when the event macro changes selection", async () => {
    vi.useFakeTimers();

    const calls: Array<{ cmd: string; args?: any }> = [];
    const doc = new DocumentController();
    const app = new FakeApp(doc);

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_open") return { ok: true, output: [], updates: [] };
      if (cmd === "fire_selection_change") {
        // Simulate the macro selecting a different range while the SelectionChange handler
        // is running (this can cause recursion in Excel unless events are disabled).
        app.emitSelection([{ startRow: 9, startCol: 9, endRow: 9, endCol: 9 }], { active: { row: 9, col: 9 } });
        return { ok: true, output: [], updates: [] };
      }
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const wiring = installVbaEventMacros({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    await flushMicrotasks();
    calls.length = 0;

    app.emitSelection([{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }]);
    await vi.advanceTimersByTimeAsync(100);
    await flushMicrotasks();

    // Allow any follow-up timers to run; there should be no second SelectionChange macro.
    await vi.advanceTimersByTimeAsync(200);
    await flushMicrotasks();

    const selectionCalls = calls.filter((c) => c.cmd === "fire_selection_change");
    expect(selectionCalls).toHaveLength(1);

    wiring.dispose();
  });

  it("fires Workbook_BeforeClose via the best-effort helper and applies updates", async () => {
    const calls: Array<{ cmd: string; args?: any }> = [];

    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") return null;
      if (cmd === "fire_workbook_before_close") {
        return {
          ok: true,
          output: [],
          updates: [
            {
              sheet_id: "Sheet1",
              row: 0,
              col: 0,
              value: 42,
              formula: null,
              display_value: "42",
            },
          ],
        };
      }
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    const doc = new DocumentController();
    const app = new FakeApp(doc);

    await fireWorkbookBeforeCloseBestEffort({
      app,
      workbookId: "workbook-1",
      invoke,
      drainBackendSync: vi.fn(async () => undefined),
    });

    expect(doc.getCell("Sheet1", { row: 0, col: 0 }).value).toBe(42);

    const setIdx = calls.findIndex((c) => c.cmd === "set_macro_ui_context");
    const closeIdx = calls.findIndex((c) => c.cmd === "fire_workbook_before_close");
    expect(closeIdx).toBeGreaterThan(setIdx);
  });
});
