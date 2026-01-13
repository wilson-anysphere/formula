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
});

