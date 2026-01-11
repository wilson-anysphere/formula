/**
 * @vitest-environment jsdom
 */

import { afterEach, describe, expect, it, vi } from "vitest";

import type { MacroBackend, MacroCellUpdate } from "../types";
import { renderMacroRunner } from "../dom_ui";

describe("renderMacroRunner", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("invokes onApplyUpdates when the backend returns updates", async () => {
    vi.stubGlobal("confirm", vi.fn(() => true));

    const updates: MacroCellUpdate[] = [
      { sheetId: "Sheet1", row: 0, col: 0, value: 42, formula: null, displayValue: "42" },
    ];

    const backend: MacroBackend = {
      listMacros: vi.fn(async () => [{ id: "macro-1", name: "Macro1", language: "vba" }]),
      runMacro: vi.fn(async () => ({ ok: true, output: ["Hello"], updates })),
    };

    const onApplyUpdates = vi.fn(async () => {});

    const container = document.createElement("div");
    document.body.appendChild(container);
    await renderMacroRunner(container, backend, "workbook-1", { onApplyUpdates });

    const runButton = container.querySelector("button");
    expect(runButton).toBeTruthy();

    await (runButton as any).onclick?.();

    expect(onApplyUpdates).toHaveBeenCalledTimes(1);
    expect(onApplyUpdates).toHaveBeenCalledWith(updates);

    const output = container.querySelector("pre");
    expect(output?.textContent).toContain("Applied 1 updates.");

    container.remove();
  });
});
