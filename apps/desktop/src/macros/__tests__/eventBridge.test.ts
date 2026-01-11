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
    const bridge = new MacroEventBridge({
      workbookId: "local-workbook",
      document: doc,
      invoke,
      drainBackendSync: async () => {},
      getSelection: () => ({ sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 }),
      debounceWorksheetMs: 200,
    });
    bridge.start();

    doc.setCellValue("Sheet1", { row: 1, col: 2 }, 1);
    doc.setCellValue("Sheet1", { row: 3, col: 1 }, 2);

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
      getSelection: () => ({ sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 }),
      debounceSelectionMs: 150,
    });
    bridge.start();

    bridge.notifySelectionChanged({ sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 });
    bridge.notifySelectionChanged({ sheetId: "Sheet1", startRow: 2, startCol: 3, endRow: 4, endCol: 5 });

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
});
