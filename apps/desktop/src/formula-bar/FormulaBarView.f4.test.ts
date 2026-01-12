/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView F4 absolute reference toggle", () => {
  it("toggles the active A1 reference and preserves a sane caret position", () => {
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
    // Caret should still be within the toggled reference token (between "$" and "1").
    expect(view.textarea.selectionStart).toBe(4);
    expect(view.textarea.selectionEnd).toBe(4);

    host.remove();
  });

  it("toggles when the caret is at the end of a reference token", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=A1";
    // Caret at end of token.
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));

    expect(view.textarea.value).toBe("=$A$1");
    expect(view.model.draft).toBe("=$A$1");
    // Caret remains at the end of the token after expansion.
    expect(view.textarea.selectionStart).toBe(5);
    expect(view.textarea.selectionEnd).toBe(5);

    host.remove();
  });

  it("cycles absolute modes correctly on repeated F4 presses", () => {
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
    expect(view.textarea.selectionStart).toBe(4);
    expect(view.textarea.selectionEnd).toBe(4);

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));
    expect(view.textarea.value).toBe("=A$1");
    expect(view.textarea.selectionStart).toBe(3);
    expect(view.textarea.selectionEnd).toBe(3);

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));
    expect(view.textarea.value).toBe("=$A1");
    expect(view.textarea.selectionStart).toBe(3);
    expect(view.textarea.selectionEnd).toBe(3);

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));
    expect(view.textarea.value).toBe("=A1");
    expect(view.textarea.selectionStart).toBe(2);
    expect(view.textarea.selectionEnd).toBe(2);

    host.remove();
  });

  it("does not toggle when the caret is not within a reference token", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=SUM(";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));

    expect(view.textarea.value).toBe("=SUM(");
    expect(view.model.draft).toBe("=SUM(");

    host.remove();
  });

  it("preserves a full-token selection when toggling absolute refs", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=A1";
    // Select the full reference token.
    view.textarea.setSelectionRange(1, 3);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));

    expect(view.textarea.value).toBe("=$A$1");
    // Full-token selection should expand to cover the toggled token.
    expect(view.textarea.selectionStart).toBe(1);
    expect(view.textarea.selectionEnd).toBe(5);

    host.remove();
  });

  it("does not toggle when the selection is not contained within a reference token", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=A1+B1";
    // Selection spans the first reference and the "+" operator.
    view.textarea.setSelectionRange(1, 4);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "F4", cancelable: true }));

    expect(view.textarea.value).toBe("=A1+B1");
    expect(view.model.draft).toBe("=A1+B1");
    expect(view.textarea.selectionStart).toBe(1);
    expect(view.textarea.selectionEnd).toBe(4);

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
