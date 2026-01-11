// @vitest-environment jsdom

import { describe, expect, it, vi } from "vitest";

vi.mock("@formula/python-runtime", () => {
  const initialize = vi.fn(async () => {});
  const execute = vi.fn(async () => ({ stdout: "", stderr: "" }));
  const destroy = vi.fn();

  return {
    PyodideRuntime: vi.fn().mockImplementation(() => ({ initialize, execute, destroy })),
    __pyodideMocks: { initialize, execute, destroy },
  };
});

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

  it("applies native Python updates via applyExternalDeltas (and tags them as python)", async () => {
    const invoke = vi.fn(async () => ({
      ok: true,
      stdout: "",
      stderr: "",
      updates: [{ sheet_id: "Sheet1", row: 0, col: 0, value: 123, formula: null, display_value: "123" }],
    }));
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const container = document.createElement("div");
    const applyExternalDeltas = vi.fn();
    const doc: any = {
      getCell: vi.fn(() => ({ value: null, formula: null, styleId: 0 })),
      applyExternalDeltas,
    };

    mountPythonPanel({
      doc,
      container,
      workbookId: "wb1",
      drainBackendSync: async () => {},
      getActiveSheetId: () => "Sheet1",
      getSelection: () => ({ sheet_id: "Sheet1", start_row: 0, start_col: 0, end_row: 0, end_col: 0 }),
    });

    const runButton = container.querySelector<HTMLButtonElement>('[data-testid="python-panel-run"]');
    expect(runButton).not.toBeNull();
    runButton?.dispatchEvent(new MouseEvent("click"));

    for (let i = 0; i < 10 && applyExternalDeltas.mock.calls.length === 0; i++) {
      // eslint-disable-next-line no-await-in-loop
      await new Promise((resolve) => setTimeout(resolve, 0));
    }

    expect(applyExternalDeltas).toHaveBeenCalledTimes(1);
    expect(applyExternalDeltas.mock.calls[0]?.[1]).toMatchObject({ source: "python" });
  });

  it("falls back to main-thread Pyodide when SharedArrayBuffer is unavailable", async () => {
    const originalSab = (globalThis as any).SharedArrayBuffer;
    const originalIsolation = (globalThis as any).crossOriginIsolated;
    delete (globalThis as any).__TAURI__;
    delete (globalThis as any).SharedArrayBuffer;
    (globalThis as any).crossOriginIsolated = false;

    try {
      const { __pyodideMocks } = await import("@formula/python-runtime");

      const container = document.createElement("div");
      const doc: any = {};

      mountPythonPanel({
        doc,
        container,
        getActiveSheetId: () => "Sheet1",
        getSelection: () => ({ sheet_id: "Sheet1", start_row: 0, start_col: 0, end_row: 0, end_col: 0 }),
      });

      const runtimeSelect = container.querySelector<HTMLSelectElement>('[data-testid="python-panel-runtime"]');
      expect(runtimeSelect).not.toBeNull();
      expect(runtimeSelect?.value).toBe("pyodide");

      const isolationLabel = container.querySelector<HTMLElement>('[data-testid="python-panel-isolation"]');
      expect(isolationLabel?.textContent).toBe(
        "SharedArrayBuffer unavailable — running Pyodide on main thread (may freeze UI)",
      );

      const runButton = container.querySelector<HTMLButtonElement>('[data-testid="python-panel-run"]');
      expect(runButton).not.toBeNull();
      runButton?.dispatchEvent(new MouseEvent("click"));

      // The run handler awaits runtime initialization before calling `execute`, so
      // the `initialize` spy can be called before the `execute` spy. Wait until the
      // script execution call is observed to avoid racey assertions.
      for (let i = 0; i < 20 && __pyodideMocks.execute.mock.calls.length === 0; i++) {
        // eslint-disable-next-line no-await-in-loop
        await new Promise((resolve) => setTimeout(resolve, 0));
      }

      expect(__pyodideMocks.initialize).toHaveBeenCalledTimes(1);
      expect(__pyodideMocks.execute).toHaveBeenCalledTimes(1);
    } finally {
      (globalThis as any).SharedArrayBuffer = originalSab;
      (globalThis as any).crossOriginIsolated = originalIsolation;
    }
  });
});
