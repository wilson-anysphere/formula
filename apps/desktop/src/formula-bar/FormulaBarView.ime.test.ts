/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView IME composition safety", () => {
  it("does not commit on Enter during composition, but does after compositionend", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.textarea.focus();
    view.textarea.value = "=1+2";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new Event("compositionstart"));
    const enterDuringComposition = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    view.textarea.dispatchEvent(enterDuringComposition);

    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(enterDuringComposition.defaultPrevented).toBe(false);

    view.textarea.dispatchEvent(new Event("compositionend"));
    const enterAfterComposition = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    view.textarea.dispatchEvent(enterAfterComposition);

    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit.mock.calls[0]?.[0]).toBe("=1+2");
    expect(onCommit.mock.calls[0]?.[1]).toEqual({ reason: "enter", shift: false });
    expect(view.model.isEditing).toBe(false);
    expect(enterAfterComposition.defaultPrevented).toBe(true);

    host.remove();
  });

  it("does not accept an AI suggestion on Tab during composition", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.textarea.focus();
    view.textarea.value = "=1+";
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    view.setAiSuggestion("=1+2");

    view.textarea.dispatchEvent(new Event("compositionstart"));
    const tabDuringComposition = new KeyboardEvent("keydown", { key: "Tab", cancelable: true });
    view.textarea.dispatchEvent(tabDuringComposition);

    // While composing, Tab should not accept AI suggestions or commit, but should still
    // prevent browser focus traversal out of the formula bar.
    expect(tabDuringComposition.defaultPrevented).toBe(true);
    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(view.model.draft).toBe("=1+");

    view.textarea.dispatchEvent(new Event("compositionend"));
    const tabAfterComposition = new KeyboardEvent("keydown", { key: "Tab", cancelable: true });
    view.textarea.dispatchEvent(tabAfterComposition);

    expect(tabAfterComposition.defaultPrevented).toBe(true);
    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(view.model.draft).toBe("=1+2");

    host.remove();
  });

  it("does not cancel on Escape during composition", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onCancel });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.textarea.focus();
    view.textarea.value = "editing";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new Event("compositionstart"));
    const escDuringComposition = new KeyboardEvent("keydown", { key: "Escape", cancelable: true });
    view.textarea.dispatchEvent(escDuringComposition);

    expect(escDuringComposition.defaultPrevented).toBe(false);
    expect(onCancel).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);

    view.textarea.dispatchEvent(new Event("compositionend"));
    const escAfterComposition = new KeyboardEvent("keydown", { key: "Escape", cancelable: true });
    view.textarea.dispatchEvent(escAfterComposition);

    expect(escAfterComposition.defaultPrevented).toBe(true);
    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(view.model.isEditing).toBe(false);
    expect(onCommit).not.toHaveBeenCalled();

    host.remove();
  });

  it("does not intercept ArrowUp/ArrowDown (function autocomplete navigation) during composition", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.textarea.focus();
    view.textarea.value = "=VLO";
    view.textarea.setSelectionRange(4, 4);
    view.textarea.dispatchEvent(new Event("input"));

    const dropdown = host.querySelector<HTMLElement>('[data-testid="formula-function-autocomplete"]');
    expect(dropdown?.hasAttribute("hidden")).toBe(false);

    view.textarea.dispatchEvent(new Event("compositionstart"));

    const downDuringComposition = new KeyboardEvent("keydown", { key: "ArrowDown", cancelable: true });
    view.textarea.dispatchEvent(downDuringComposition);
    expect(downDuringComposition.defaultPrevented).toBe(false);

    const upDuringComposition = new KeyboardEvent("keydown", { key: "ArrowUp", cancelable: true });
    view.textarea.dispatchEvent(upDuringComposition);
    expect(upDuringComposition.defaultPrevented).toBe(false);

    view.textarea.dispatchEvent(new Event("compositionend"));

    const downAfterComposition = new KeyboardEvent("keydown", { key: "ArrowDown", cancelable: true });
    view.textarea.dispatchEvent(downAfterComposition);
    expect(downAfterComposition.defaultPrevented).toBe(true);

    host.remove();
  });
});
