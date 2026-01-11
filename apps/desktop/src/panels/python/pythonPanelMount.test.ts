// @vitest-environment jsdom

import { describe, expect, it, vi } from "vitest";

vi.mock("@formula/python-runtime", () => ({
  PyodideRuntime: class {
    initialize = vi.fn(async () => {});
    execute = vi.fn(async () => ({ stdout: "", stderr: "" }));
    destroy = vi.fn();
  },
}));

vi.mock("@formula/python-runtime/document-controller", () => ({
  DocumentControllerBridge: class {
    activeSheetId: string;
    sheetIds: Set<string>;
    selection: any;

    constructor(_doc: any, options: any = {}) {
      this.activeSheetId = options.activeSheetId ?? "Sheet1";
      this.sheetIds = new Set([this.activeSheetId]);
      this.selection = {
        sheet_id: this.activeSheetId,
        start_row: 0,
        start_col: 0,
        end_row: 0,
        end_col: 0,
      };
    }

    get_active_sheet_id() {
      return this.activeSheetId;
    }

    get_selection() {
      return { ...this.selection };
    }

    set_selection({ selection }: any) {
      this.selection = { ...selection };
      this.activeSheetId = selection.sheet_id;
      this.sheetIds.add(selection.sheet_id);
      return null;
    }
  },
}));

import { mountPythonPanel } from "./pythonPanelMount.js";

describe("pythonPanelMount", () => {
  it("invokes run_python_script when native runtime is selected", async () => {
    const invoke = vi.fn(async () => ({ ok: true, stdout: "", stderr: "", updates: [] }));
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const container = document.createElement("div");
    const doc: any = { setCellInput: vi.fn(), beginBatch: vi.fn(), endBatch: vi.fn(), cancelBatch: vi.fn() };

    mountPythonPanel({
      doc,
      container,
      workbookId: "wb1",
      drainBackendSync: async () => {},
      getActiveSheetId: () => "Sheet1",
      getSelection: () => ({ sheet_id: "Sheet1", start_row: 0, start_col: 0, end_row: 0, end_col: 0 }),
    });

    const runtimeSelect = container.querySelector<HTMLSelectElement>('[data-testid="python-panel-runtime"]');
    expect(runtimeSelect).not.toBeNull();
    expect(runtimeSelect?.value).toBe("native");

    const runButton = container.querySelector<HTMLButtonElement>('[data-testid="python-panel-run"]');
    expect(runButton).not.toBeNull();

    runButton?.dispatchEvent(new MouseEvent("click"));

    // Wait for the click handler's microtasks + invoke promise.
    for (let i = 0; i < 5 && invoke.mock.calls.length === 0; i++) {
      // eslint-disable-next-line no-await-in-loop
      await new Promise((resolve) => setTimeout(resolve, 0));
    }

    expect(invoke).toHaveBeenCalledTimes(1);
    expect(invoke.mock.calls[0]?.[0]).toBe("run_python_script");
    expect(invoke.mock.calls[0]?.[1]).toMatchObject({
      workbook_id: "wb1",
      code: expect.any(String),
    });
  });
});

