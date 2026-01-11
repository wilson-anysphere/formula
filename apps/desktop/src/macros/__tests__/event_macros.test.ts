/**
 * @vitest-environment jsdom
 */

import { afterEach, describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { installVbaEventMacros } from "../event_macros";

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
});
