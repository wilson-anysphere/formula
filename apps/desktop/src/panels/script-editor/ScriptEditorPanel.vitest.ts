/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

const runSpy = vi.fn(async () => ({ logs: [{ level: "info", message: "ok" }], error: null }));

vi.mock("@formula/scripting/web", () => ({
  FORMULA_API_DTS: "declare const __formula: number;",
  ScriptRuntime: vi.fn().mockImplementation(() => ({ run: runSpy })),
}));

import { mountScriptEditorPanel } from "./ScriptEditorPanel.js";

describe("ScriptEditorPanel", () => {
  it("disables Run while editing or read-only", async () => {
    const container = document.createElement("div");
    let editing = true;
    let readOnly = false;

    const mounted = mountScriptEditorPanel({
      workbook: {},
      container,
      isEditing: () => editing,
      isReadOnly: () => readOnly,
    });

    const runButton = container.querySelector<HTMLButtonElement>('[data-testid="script-editor-run"]');
    expect(runButton).not.toBeNull();
    expect(runButton?.disabled).toBe(true);

    // Ending edit mode should re-enable.
    editing = false;
    window.dispatchEvent(new CustomEvent("formula:spreadsheet-editing-changed", { detail: { isEditing: false } }));
    expect(runButton?.disabled).toBe(false);

    runButton?.click();
    for (let i = 0; i < 10 && runSpy.mock.calls.length === 0; i++) {
      // eslint-disable-next-line no-await-in-loop
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    expect(runSpy).toHaveBeenCalledTimes(1);

    // Read-only should disable again.
    readOnly = true;
    window.dispatchEvent(new CustomEvent("formula:read-only-changed", { detail: { readOnly: true } }));
    expect(runButton?.disabled).toBe(true);

    mounted.dispose();
  });
});

