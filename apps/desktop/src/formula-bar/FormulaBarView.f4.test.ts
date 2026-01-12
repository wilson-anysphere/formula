/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView F4 absolute reference toggle", () => {
  it("toggles the active A1 reference and keeps the token selected for repeated toggles", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=A1";
    // Caret between A and 1.
    view.textarea.setSelectionRange(2, 2);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));

    expect(view.textarea.value).toBe("=$A$1");
    expect(view.model.draft).toBe("=$A$1");
    // Excel UX: keep the full reference token selected so repeated F4 presses
    // continue cycling the same token.
    expect(view.textarea.selectionStart).toBe(1);
    expect(view.textarea.selectionEnd).toBe(5);

    host.remove();
  });

  it("does not toggle non-formula text", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "hello";
    view.textarea.setSelectionRange(1, 1);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));

    expect(view.textarea.value).toBe("hello");
    expect(view.model.draft).toBe("hello");

    host.remove();
  });
});
