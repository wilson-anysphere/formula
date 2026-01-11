// @vitest-environment jsdom

import { describe, expect, it, vi } from "vitest";

import { renderMacroRunner } from "./dom_ui";
import type { MacroBackend } from "./types";

describe("renderMacroRunner", () => {
  it("calls onApplyUpdates when the backend returns cell updates", async () => {
    vi.spyOn(window, "confirm").mockReturnValue(true);

    const runMacro = vi.fn<MacroBackend["runMacro"]>(async () => ({
      ok: true,
      output: [],
      updates: [
        {
          sheetId: "Sheet1",
          row: 0,
          col: 0,
          value: 123,
          formula: null,
          displayValue: "123",
        },
      ],
    }));

    const backend: MacroBackend = {
      listMacros: async () => [{ id: "m1", name: "Macro1", language: "vba" }],
      getMacroSecurityStatus: async () => ({ hasMacros: true, trust: "trusted_once", signature: { status: "unsigned" } }),
      setMacroTrust: async (_workbookId, decision) => ({ hasMacros: true, trust: decision, signature: { status: "unsigned" } }),
      runMacro,
    };

    const onApplyUpdates = vi.fn();

    const container = document.createElement("div");
    await renderMacroRunner(container, backend, "wb1", { onApplyUpdates });

    const runButton = container.querySelectorAll("button")[1];
    expect(runButton).not.toBeNull();

    await (runButton as any).onclick(new MouseEvent("click"));

    expect(onApplyUpdates).toHaveBeenCalledTimes(1);
    expect(onApplyUpdates).toHaveBeenCalledWith([
      {
        sheetId: "Sheet1",
        row: 0,
        col: 0,
        value: 123,
        formula: null,
        displayValue: "123",
      },
    ]);
    expect(runMacro).toHaveBeenCalledTimes(1);
  });
});
