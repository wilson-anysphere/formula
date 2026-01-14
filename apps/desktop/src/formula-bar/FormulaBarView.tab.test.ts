/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView Tab/Shift+Tab commit semantics", () => {
  it("uses Tab to accept an AI suggestion (and does not commit)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=1+";
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    view.setAiSuggestion("=1+2");

    const e = new KeyboardEvent("keydown", { key: "Tab", cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(view.model.draft).toBe("=1+2");
    expect(view.isEditing()).toBe(true);
    expect(onCommit).not.toHaveBeenCalled();

    host.remove();
  });

  it("uses Tab to commit when no AI suggestion is present", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "hello";
    view.textarea.setSelectionRange(5, 5);
    view.textarea.dispatchEvent(new Event("input"));

    const e = new KeyboardEvent("keydown", { key: "Tab", cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit.mock.calls[0]?.[0]).toBe("hello");
    expect(onCommit.mock.calls[0]?.[1]).toEqual({ reason: "tab", shift: false });

    host.remove();
  });

  it("uses Shift+Tab to commit", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "hello";
    view.textarea.setSelectionRange(5, 5);
    view.textarea.dispatchEvent(new Event("input"));

    const e = new KeyboardEvent("keydown", { key: "Tab", shiftKey: true, cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit.mock.calls[0]?.[0]).toBe("hello");
    expect(onCommit.mock.calls[0]?.[1]).toEqual({ reason: "tab", shift: true });

    host.remove();
  });

  it("uses Shift+Tab to commit even when an AI suggestion is present (does not accept)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "", value: null });

    view.focus({ cursor: "end" });
    view.textarea.value = "=1+";
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    view.setAiSuggestion("=1+2");

    const e = new KeyboardEvent("keydown", { key: "Tab", shiftKey: true, cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit.mock.calls[0]?.[0]).toBe("=1+");
    expect(onCommit.mock.calls[0]?.[1]).toEqual({ reason: "tab", shift: true });

    host.remove();
  });
});
