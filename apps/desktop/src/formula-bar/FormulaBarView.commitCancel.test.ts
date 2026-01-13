/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

function queryActions(host: HTMLElement): {
  cancel: HTMLButtonElement;
  commit: HTMLButtonElement;
} {
  const cancel = host.querySelector<HTMLButtonElement>(".formula-bar-action-button--cancel");
  const commit = host.querySelector<HTMLButtonElement>(".formula-bar-action-button--commit");
  if (!cancel || !commit) {
    throw new Error("Expected commit/cancel buttons to exist");
  }
  return { cancel, commit };
}

describe("FormulaBarView commit/cancel UX", () => {
  it("hides commit/cancel buttons when not editing and shows them on focus", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    const { cancel, commit } = queryActions(host);

    expect(view.model.isEditing).toBe(false);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(false);
    expect(cancel.hidden).toBe(true);
    expect(cancel.disabled).toBe(true);
    expect(commit.hidden).toBe(true);
    expect(commit.disabled).toBe(true);

    view.textarea.focus();

    expect(view.model.isEditing).toBe(true);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(true);
    expect(cancel.hidden).toBe(false);
    expect(cancel.disabled).toBe(false);
    expect(commit.hidden).toBe(false);
    expect(commit.disabled).toBe(false);

    host.remove();
  });

  it("commits on Enter (without Alt), exits edit mode, and hides buttons again", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    const { cancel, commit } = queryActions(host);

    view.textarea.focus();
    view.textarea.value = "=1+2";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", cancelable: true }));

    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith("=1+2");
    expect(view.model.isEditing).toBe(false);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(false);
    expect(cancel.hidden).toBe(true);
    expect(cancel.disabled).toBe(true);
    expect(commit.hidden).toBe(true);
    expect(commit.disabled).toBe(true);

    host.remove();
  });

  it("does not commit on Alt+Enter (reserved for newline/indent)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    const { cancel, commit } = queryActions(host);

    view.textarea.focus();
    view.textarea.value = "line1";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", altKey: true, cancelable: true }));

    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(true);
    expect(cancel.hidden).toBe(false);
    expect(cancel.disabled).toBe(false);
    expect(commit.hidden).toBe(false);
    expect(commit.disabled).toBe(false);

    host.remove();
  });

  it("cancels on Escape, restores the active cell input, and exits edit mode", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onCancel });
    const { cancel, commit } = queryActions(host);

    view.setActiveCell({ address: "A1", input: "original", value: null });
    expect(view.textarea.value).toBe("original");

    view.textarea.focus();
    view.textarea.value = "changed";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", cancelable: true }));

    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(view.model.isEditing).toBe(false);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(false);
    expect(view.textarea.value).toBe("original");
    expect(cancel.hidden).toBe(true);
    expect(cancel.disabled).toBe(true);
    expect(commit.hidden).toBe(true);
    expect(commit.disabled).toBe(true);

    host.remove();
  });

  it("commits/cancels via ✓/✕ buttons with the same behavior as Enter/Escape", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onCancel });
    const { cancel, commit } = queryActions(host);

    view.setActiveCell({ address: "A1", input: "start", value: null });

    view.textarea.focus();
    view.textarea.value = "cancel-me";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    cancel.click();

    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(view.model.isEditing).toBe(false);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(false);
    expect(view.textarea.value).toBe("start");
    expect(cancel.hidden).toBe(true);
    expect(cancel.disabled).toBe(true);
    expect(commit.hidden).toBe(true);
    expect(commit.disabled).toBe(true);

    view.textarea.focus();
    view.textarea.value = "commit-me";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    commit.click();

    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith("commit-me");
    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(view.model.isEditing).toBe(false);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(false);
    expect(cancel.hidden).toBe(true);
    expect(cancel.disabled).toBe(true);
    expect(commit.hidden).toBe(true);
    expect(commit.disabled).toBe(true);

    host.remove();
  });
});
