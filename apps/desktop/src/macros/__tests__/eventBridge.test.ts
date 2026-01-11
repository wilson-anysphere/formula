/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { MacroEventBridge } from "../eventBridge";

describe("MacroEventBridge", () => {
  it("fires Worksheet_Change once per debounce window with the correct bounding box", async () => {
    vi.useFakeTimers();

    const calls: Array<{ cmd: string; args: any }> = [];
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "set_macro_ui_context") return null;
      return { ok: true, output: [], updates: [] };
    });
 
    const doc = new DocumentController();
    let selection = { sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0, activeRow: 7, activeCol: 8 };
    const bridge = new MacroEventBridge({
      workbookId: "local-workbook",
      document: doc,
      invoke,
      drainBackendSync: async () => {},
      getSelection: () => selection,
      debounceWorksheetMs: 200,
    });
    bridge.start();

    doc.setCellValue("Sheet1", { row: 1, col: 2 }, 1);
    doc.setCellValue("Sheet1", { row: 3, col: 1 }, 2);

    // Simulate the user changing selection after editing but before the debounced
    // Worksheet_Change macro executes. We should still run with the selection context
    // captured at edit time.
    selection = { sheetId: "Sheet1", startRow: 9, startCol: 9, endRow: 9, endCol: 9, activeRow: 9, activeCol: 9 };

    await vi.advanceTimersByTimeAsync(250);
    await bridge.whenIdle();

    const worksheetCalls = calls.filter((c) => c.cmd === "fire_worksheet_change");
    expect(worksheetCalls).toHaveLength(1);
    expect(worksheetCalls[0]?.args).toMatchObject({
      sheet_id: "Sheet1",
      start_row: 1,
      start_col: 1,
      end_row: 3,
      end_col: 2,
    });
 
    const ctxCalls = calls.filter((c) => c.cmd === "set_macro_ui_context");
    expect(ctxCalls).toHaveLength(1);
    expect(ctxCalls[0]?.args).toMatchObject({
      sheet_id: "Sheet1",
      active_row: 7,
      active_col: 8,
      selection: { start_row: 0, start_col: 0, end_row: 0, end_col: 0 },
    });
 
    vi.useRealTimers();
  });

  it("fires Worksheet_SelectionChange for debounced selection changes", async () => {
    vi.useFakeTimers();

    const calls: Array<{ cmd: string; args: any }> = [];
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "set_macro_ui_context") return null;
      return { ok: true, output: [], updates: [] };
    });

    const doc = new DocumentController();
    const bridge = new MacroEventBridge({
      workbookId: "local-workbook",
      document: doc,
      invoke,
      drainBackendSync: async () => {},
      getSelection: () => ({
        sheetId: "Sheet1",
        startRow: 0,
        startCol: 0,
        endRow: 0,
        endCol: 0,
        activeRow: 0,
        activeCol: 0,
      }),
      debounceSelectionMs: 150,
    });
    bridge.start();

    bridge.notifySelectionChanged({ sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0, activeRow: 0, activeCol: 0 });
    bridge.notifySelectionChanged({ sheetId: "Sheet1", startRow: 2, startCol: 3, endRow: 4, endCol: 5, activeRow: 4, activeCol: 5 });

    await vi.advanceTimersByTimeAsync(200);
    await bridge.whenIdle();

    const selectionCalls = calls.filter((c) => c.cmd === "fire_selection_change");
    expect(selectionCalls).toHaveLength(1);
    expect(selectionCalls[0]?.args).toMatchObject({
      sheet_id: "Sheet1",
      start_row: 2,
      start_col: 3,
      end_row: 4,
      end_col: 5,
    });

    const ctxCall = calls.find((c) => c.cmd === "set_macro_ui_context");
    expect(ctxCall?.args).toMatchObject({
      sheet_id: "Sheet1",
      active_row: 4,
      active_col: 5,
      selection: { start_row: 2, start_col: 3, end_row: 4, end_col: 5 },
    });

    vi.useRealTimers();
  });

  it("fires Worksheet_SelectionChange when only the active cell changes within a selection", async () => {
    vi.useFakeTimers();

    const calls: Array<{ cmd: string; args: any }> = [];
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "set_macro_ui_context") return null;
      return { ok: true, output: [], updates: [] };
    });

    const doc = new DocumentController();
    let selection = { sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 1, endCol: 1, activeRow: 0, activeCol: 0 };
    const bridge = new MacroEventBridge({
      workbookId: "local-workbook",
      document: doc,
      invoke,
      drainBackendSync: async () => {},
      getSelection: () => selection,
      debounceSelectionMs: 100,
    });
    bridge.start();

    // Move the active cell while keeping the selected range unchanged.
    selection = { ...selection, activeRow: 1, activeCol: 0 };
    bridge.notifySelectionChanged(selection);

    await vi.advanceTimersByTimeAsync(150);
    await bridge.whenIdle();

    expect(calls.some((c) => c.cmd === "fire_selection_change")).toBe(true);
    const ctxCall = calls.find((c) => c.cmd === "set_macro_ui_context");
    expect(ctxCall?.args).toMatchObject({
      sheet_id: "Sheet1",
      active_row: 1,
      active_col: 0,
      selection: { start_row: 0, start_col: 0, end_row: 1, end_col: 1 },
    });

    vi.useRealTimers();
  });

  it("suppresses Worksheet_Change while applying macro updates", async () => {
    vi.useFakeTimers();

    const calls: Array<{ cmd: string; args: any }> = [];
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      calls.push({ cmd, args });
      if (cmd === "set_macro_ui_context") return null;
      return { ok: true, output: [], updates: [] };
    });

    const doc = new DocumentController();
    const bridge = new MacroEventBridge({
      workbookId: "local-workbook",
      document: doc,
      invoke,
      drainBackendSync: async () => {},
      getSelection: () => ({ sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 }),
      debounceWorksheetMs: 100,
    });
    bridge.start();

    bridge.applyMacroUpdates(
      [{ sheetId: "Sheet1", row: 0, col: 0, value: 42, formula: null, displayValue: "42" }],
      { label: "Apply macro updates" },
    );

    await vi.advanceTimersByTimeAsync(200);
    await bridge.whenIdle();

    expect(calls.some((c) => c.cmd === "fire_worksheet_change")).toBe(false);

    vi.useRealTimers();
  });

  it("silently ignores macro event invocations when no workbook is loaded", async () => {
    const invoke = vi.fn(async () => {
      throw "no workbook loaded";
    });

    const doc = new DocumentController();
    const bridge = new MacroEventBridge({
      workbookId: "local-workbook",
      document: doc,
      invoke,
      drainBackendSync: async () => {},
      getSelection: () => ({ sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 }),
    });

    await expect(bridge.fireWorkbookOpen()).resolves.toBeUndefined();
    expect(document.getElementById("macro-event-banner-container")).toBeNull();
  });
});
