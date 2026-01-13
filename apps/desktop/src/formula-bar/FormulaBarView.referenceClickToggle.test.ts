/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView reference token click selection toggle", () => {
  it("selects a reference token on first click and toggles back to a caret on second click", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });
    view.focus({ cursor: "end" });

    view.textarea.value = "=A1+B1";

    // Place caret inside "A1" (between A and 1), then click:
    // Excel UX should expand selection to the full reference token.
    view.textarea.setSelectionRange(2, 2);
    view.textarea.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(view.textarea.selectionStart).toBe(1);
    expect(view.textarea.selectionEnd).toBe(3);

    // Second click on the same token should toggle back to a caret for manual edits.
    // (In browsers, the click itself collapses the selection before firing; emulate that.)
    view.textarea.setSelectionRange(2, 2);
    view.textarea.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(view.textarea.selectionStart).toBe(2);
    expect(view.textarea.selectionEnd).toBe(2);

    // Repeat for the second reference to ensure the correct token is selected.
    view.textarea.setSelectionRange(5, 5);
    view.textarea.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(view.textarea.selectionStart).toBe(4);
    expect(view.textarea.selectionEnd).toBe(6);

    view.textarea.setSelectionRange(5, 5);
    view.textarea.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(view.textarea.selectionStart).toBe(5);
    expect(view.textarea.selectionEnd).toBe(5);

    host.remove();
  });
});

