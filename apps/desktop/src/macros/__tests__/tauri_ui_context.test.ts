/**
 * @vitest-environment jsdom
 */

import { afterEach, describe, expect, it, vi } from "vitest";

import { renderMacroRunner } from "../dom_ui";
import { TauriMacroBackend, wrapTauriMacroBackendWithUiContext } from "../tauri_backend";

describe("Tauri macro backend UI context", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    delete (globalThis as any).__TAURI__;
  });

  it("invokes set_macro_ui_context before run_macro", async () => {
    const calls: string[] = [];
    const invoke = vi.fn(async (cmd: string) => {
      calls.push(cmd);
      if (cmd === "list_macros") {
        return [{ id: "Macro1", name: "Macro1", language: "vba" }];
      }
      if (cmd === "get_macro_security_status") {
        return {
          has_macros: true,
          origin_path: null,
          workbook_fingerprint: null,
          signature: null,
          trust: "trusted_always",
        };
      }
      if (cmd === "set_macro_ui_context") {
        return null;
      }
      if (cmd === "run_macro") {
        return { ok: true, output: [], updates: [] };
      }
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = { core: { invoke } };
    vi.stubGlobal("confirm", vi.fn(() => true));

    const base = new TauriMacroBackend();
    const backend = wrapTauriMacroBackendWithUiContext(base, () => ({
      sheetId: "Sheet2",
      activeRow: 2,
      activeCol: 2,
      selection: { startRow: 2, startCol: 2, endRow: 3, endCol: 3 },
    }));

    const container = document.createElement("div");
    document.body.appendChild(container);
    await renderMacroRunner(container, backend, "workbook-1");

    const runButton = Array.from(container.querySelectorAll("button")).find((btn) => btn.textContent === "Run");
    expect(runButton).toBeTruthy();

    await (runButton as any).onclick?.();

    const runIdx = calls.indexOf("run_macro");
    expect(runIdx).toBeGreaterThan(0);
    expect(calls[runIdx - 1]).toBe("set_macro_ui_context");

    container.remove();
  });

  it("treats \"no workbook loaded\" as no VBA macros (so script macros can still run)", async () => {
    const invoke = vi.fn(async (cmd: string) => {
      if (cmd === "list_macros") {
        throw "no workbook loaded";
      }
      if (cmd === "get_macro_security_status") {
        throw "no workbook loaded";
      }
      throw new Error(`Unexpected invoke: ${cmd}`);
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const backend = new TauriMacroBackend();
    await expect(backend.listMacros("workbook-1")).resolves.toEqual([]);
    await expect(backend.getMacroSecurityStatus("workbook-1")).resolves.toMatchObject({ hasMacros: false, trust: "blocked" });
  });

  it("sanitizes macro UI context indices before invoking set_macro_ui_context", async () => {
    const invoke = vi.fn(async (cmd: string, args?: any) => {
      if (cmd !== "set_macro_ui_context") throw new Error(`Unexpected invoke: ${cmd}`);
      return args;
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (globalThis as any).__TAURI__ = { core: { invoke } };

    const backend = new TauriMacroBackend();
    await backend.setMacroUiContext({
      workbookId: "workbook-1",
      sheetId: "Sheet1",
      activeRow: -1,
      activeCol: 1.9,
      selection: { startRow: -5, startCol: 2.4, endRow: 3.1, endCol: Number.NaN },
    });

    expect(invoke).toHaveBeenCalledWith("set_macro_ui_context", {
      workbook_id: "workbook-1",
      sheet_id: "Sheet1",
      active_row: 0,
      active_col: 1,
      selection: { start_row: 0, start_col: 0, end_row: 3, end_col: 2 },
    });
  });
});
