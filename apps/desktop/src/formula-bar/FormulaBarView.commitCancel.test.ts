/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";
import { parseA1Range } from "../spreadsheet/a1.js";

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

function queryFxButton(host: HTMLElement): HTMLButtonElement {
  const fx = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]');
  if (!fx) throw new Error("Expected fx button to exist");
  return fx;
}

describe("FormulaBarView commit/cancel UX", () => {
  it("ignores commit/cancel actions while not editing", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onCancel });
    const { cancel, commit } = queryActions(host);

    expect(view.model.isEditing).toBe(false);
    expect(cancel.hidden).toBe(true);
    expect(commit.hidden).toBe(true);

    // Even if events are dispatched (e.g. programmatically), the view should guard
    // against committing/canceling while not editing.
    cancel.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    commit.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    const enter = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    view.textarea.dispatchEvent(enter);
    expect(enter.defaultPrevented).toBe(false);

    const escape = new KeyboardEvent("keydown", { key: "Escape", cancelable: true });
    view.textarea.dispatchEvent(escape);
    expect(escape.defaultPrevented).toBe(false);

    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(false);

    host.remove();
  });

  it("fires onBeginEdit once when focus begins editing", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onBeginEdit = vi.fn();
    const view = new FormulaBarView(host, { onCommit: () => {}, onBeginEdit });
    view.setActiveCell({ address: "B2", input: "", value: null });

    view.textarea.focus();

    expect(onBeginEdit).toHaveBeenCalledTimes(1);
    expect(onBeginEdit).toHaveBeenCalledWith("B2");

    // Re-focusing while already editing should not re-fire onBeginEdit.
    view.textarea.blur();
    view.textarea.focus();
    expect(onBeginEdit).toHaveBeenCalledTimes(1);

    host.remove();
  });

  it("commitEdit()/cancelEdit() are no-ops when not editing", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onCancel });

    view.commitEdit();
    view.cancelEdit();

    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(false);

    host.remove();
  });

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

  it("begins editing when clicking the highlight in view mode (click-to-edit)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    const { cancel, commit } = queryActions(host);

    view.setActiveCell({ address: "A1", input: "hello", value: null });
    expect(view.model.isEditing).toBe(false);
    expect(cancel.hidden).toBe(true);
    expect(commit.hidden).toBe(true);

    const highlight = host.querySelector<HTMLElement>('[data-testid="formula-highlight"]');
    expect(highlight).toBeTruthy();

    const e = new MouseEvent("mousedown", { bubbles: true, cancelable: true });
    highlight!.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(view.model.isEditing).toBe(true);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(true);
    expect(cancel.hidden).toBe(false);
    expect(commit.hidden).toBe(false);
    expect(document.activeElement).toBe(view.textarea);
    expect(view.textarea.value).toBe("hello");
    // Click-to-edit should focus with caret at end (see highlight mousedown handler).
    expect(view.textarea.selectionStart).toBe(5);
    expect(view.textarea.selectionEnd).toBe(5);

    host.remove();
  });

  it("keeps edit mode active when the textarea blurs (Excel-style range selection mode)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    const { cancel, commit } = queryActions(host);

    view.textarea.focus();
    view.textarea.value = "=SUM(";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.textarea.blur();

    expect(view.model.isEditing).toBe(true);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(true);
    expect(cancel.hidden).toBe(false);
    expect(cancel.disabled).toBe(false);
    expect(commit.hidden).toBe(false);
    expect(commit.disabled).toBe(false);

    host.remove();
  });

  it("commits via commitEdit() API with reason=command by default", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });

    view.textarea.focus();
    view.textarea.value = "api-commit";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.commitEdit();

    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith("api-commit", { reason: "command", shift: false });
    expect(view.model.isEditing).toBe(false);
    expect(view.model.activeCell.input).toBe("api-commit");

    host.remove();
  });

  it("cancels via cancelEdit() API and restores the active cell input", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onCancel });
    view.setActiveCell({ address: "A1", input: "original", value: null });

    view.textarea.focus();
    view.textarea.value = "changed";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    view.cancelEdit();

    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(view.model.isEditing).toBe(false);
    expect(view.model.draft).toBe("original");
    expect(view.model.activeCell.input).toBe("original");
    expect(view.textarea.value).toBe("original");

    host.remove();
  });

  it("commitEdit('enter') commits even when the textarea is not focused", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });

    view.textarea.focus();
    view.textarea.value = "grid-style";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));
    view.textarea.blur();

    view.commitEdit("enter", false);

    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith("grid-style", { reason: "enter", shift: false });
    expect(view.model.isEditing).toBe(false);

    host.remove();
  });

  it("cancelEdit() cancels even when the textarea is not focused", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit: () => {}, onCancel });
    view.setActiveCell({ address: "A1", input: "original", value: null });

    view.textarea.focus();
    view.textarea.value = "changed";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));
    view.textarea.blur();

    view.cancelEdit();

    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(view.model.isEditing).toBe(false);
    expect(view.model.draft).toBe("original");

    host.remove();
  });

  it("closes the function picker when canceling via ✕", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onCancel });
    const { cancel } = queryActions(host);
    const fx = queryFxButton(host);

    view.setActiveCell({ address: "A1", input: "start", value: null });
    fx.click();

    const picker = host.querySelector<HTMLElement>('[data-testid="formula-function-picker"]');
    expect(picker?.hidden).toBe(false);

    cancel.click();

    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(view.model.isEditing).toBe(false);
    expect(view.model.draft).toBe("start");
    expect(view.textarea.value).toBe("start");
    expect(picker?.hidden).toBe(true);

    host.remove();
  });

  it("does not cancel edit when pressing Escape in the function picker (it only closes the picker)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onCancel });
    const { cancel, commit } = queryActions(host);
    const fx = queryFxButton(host);

    fx.click();

    const picker = host.querySelector<HTMLElement>('[data-testid="formula-function-picker"]')!;
    const pickerInput = host.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]')!;
    expect(picker.hidden).toBe(false);
    expect(document.activeElement).toBe(pickerInput);
    expect(view.model.isEditing).toBe(true);
    expect(cancel.hidden).toBe(false);
    expect(commit.hidden).toBe(false);

    const e = new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true });
    pickerInput.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(picker.hidden).toBe(true);
    expect(document.activeElement).toBe(view.textarea);
    // Crucially: Escape should not cancel the entire edit; it should only close the picker.
    expect(view.model.isEditing).toBe(true);
    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).not.toHaveBeenCalled();

    host.remove();
  });

  it("does not commit edit when pressing Enter in the function picker (it inserts a function)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    const fx = queryFxButton(host);

    fx.click();

    const picker = host.querySelector<HTMLElement>('[data-testid="formula-function-picker"]')!;
    const pickerInput = host.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]')!;
    expect(picker.hidden).toBe(false);

    // Ensure a deterministic selection.
    pickerInput.value = "sum";
    pickerInput.dispatchEvent(new Event("input", { bubbles: true }));

    const e = new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true });
    pickerInput.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(view.textarea.value).toBe("=SUM(");
    expect(document.activeElement).toBe(view.textarea);
    expect(picker.hidden).toBe(true);

    host.remove();
  });

  it("closes the function picker when committing via ✓", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onCancel });
    const { commit } = queryActions(host);
    const fx = queryFxButton(host);

    view.setActiveCell({ address: "A1", input: "", value: null });
    fx.click();

    // Edit the draft while the picker is open.
    view.textarea.value = "=1+2";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    const picker = host.querySelector<HTMLElement>('[data-testid="formula-function-picker"]');
    expect(picker?.hidden).toBe(false);

    commit.click();

    expect(onCancel).not.toHaveBeenCalled();
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith("=1+2", { reason: "command", shift: false });
    expect(view.model.isEditing).toBe(false);
    expect(view.model.draft).toBe("=1+2");
    expect(view.model.activeCell.input).toBe("=1+2");
    expect(picker?.hidden).toBe(true);

    host.remove();
  });

  it("closes the function picker when canceling via Escape (while focused in textarea)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit: () => {}, onCancel });
    const fx = queryFxButton(host);

    view.setActiveCell({ address: "A1", input: "start", value: null });
    fx.click();

    const picker = host.querySelector<HTMLElement>('[data-testid="formula-function-picker"]');
    expect(picker?.hidden).toBe(false);

    // Bring focus back to the textarea so Escape hits the formula bar handler (not the picker handler).
    view.textarea.focus();

    const e = new KeyboardEvent("keydown", { key: "Escape", cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(view.model.isEditing).toBe(false);
    expect(picker?.hidden).toBe(true);

    host.remove();
  });

  it("closes the function picker when committing via Enter (while focused in textarea)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    const fx = queryFxButton(host);

    view.setActiveCell({ address: "A1", input: "", value: null });
    fx.click();

    const picker = host.querySelector<HTMLElement>('[data-testid="formula-function-picker"]');
    expect(picker?.hidden).toBe(false);

    view.textarea.value = "=1";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    // Bring focus back to the textarea so Enter hits the formula bar handler (not the picker handler).
    view.textarea.focus();

    const e = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith("=1", { reason: "enter", shift: false });
    expect(view.model.isEditing).toBe(false);
    expect(picker?.hidden).toBe(true);

    host.remove();
  });

  it("clears hover + reference highlights on commit", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onHoverRange = vi.fn();
    const onReferenceHighlights = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onHoverRange, onReferenceHighlights });
    view.setActiveCell({ address: "A1", input: "=A1", value: null });

    view.textarea.focus();

    const lastHoverBefore = onHoverRange.mock.calls.at(-1)?.[0] ?? null;
    const lastHighlightsBefore = onReferenceHighlights.mock.calls.at(-1)?.[0] ?? [];
    expect(lastHoverBefore).toEqual(parseA1Range("A1"));
    expect(lastHighlightsBefore.length).toBeGreaterThan(0);

    const e = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onHoverRange.mock.calls.at(-1)?.[0] ?? null).toBeNull();
    expect(onReferenceHighlights.mock.calls.at(-1)?.[0] ?? null).toEqual([]);

    host.remove();
  });

  it("clears hover + reference highlights on cancel", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const onHoverRange = vi.fn();
    const onReferenceHighlights = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onCancel, onHoverRange, onReferenceHighlights });
    view.setActiveCell({ address: "A1", input: "=A1", value: null });

    view.textarea.focus();

    const lastHoverBefore = onHoverRange.mock.calls.at(-1)?.[0] ?? null;
    const lastHighlightsBefore = onReferenceHighlights.mock.calls.at(-1)?.[0] ?? [];
    expect(lastHoverBefore).toEqual(parseA1Range("A1"));
    expect(lastHighlightsBefore.length).toBeGreaterThan(0);

    const e = new KeyboardEvent("keydown", { key: "Escape", cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(onHoverRange.mock.calls.at(-1)?.[0] ?? null).toBeNull();
    expect(onReferenceHighlights.mock.calls.at(-1)?.[0] ?? null).toEqual([]);

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

    const e = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith("=1+2", { reason: "enter", shift: false });
    expect(view.model.isEditing).toBe(false);
    expect(view.model.draft).toBe("=1+2");
    expect(view.model.activeCell.input).toBe("=1+2");
    expect(document.activeElement).not.toBe(view.textarea);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(false);
    expect(cancel.hidden).toBe(true);
    expect(cancel.disabled).toBe(true);
    expect(commit.hidden).toBe(true);
    expect(commit.disabled).toBe(true);

    host.remove();
  });

  it("commits the typed draft (not the AI suggestion) when pressing Enter with a suggestion present", () => {
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

    const e = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith("=1+", { reason: "enter", shift: false });
    expect(view.model.aiSuggestion()).toBeNull();
    expect(view.model.isEditing).toBe(false);

    host.remove();
  });

  it("clears AI suggestion state when canceling", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onCancel });
    view.setActiveCell({ address: "A1", input: "orig", value: null });

    view.textarea.focus();
    view.textarea.value = "=1+";
    view.textarea.setSelectionRange(3, 3);
    view.textarea.dispatchEvent(new Event("input"));

    view.setAiSuggestion("=1+2");
    expect(view.model.aiSuggestion()).toBe("=1+2");
    expect(view.model.aiGhostText()).toBe("2");

    const e = new KeyboardEvent("keydown", { key: "Escape", cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(view.model.isEditing).toBe(false);
    expect(view.model.draft).toBe("orig");
    expect(view.model.aiSuggestion()).toBeNull();
    expect(view.model.aiGhostText()).toBe("");

    host.remove();
  });

  it("does not commit twice if Enter is pressed multiple times", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });

    view.textarea.focus();
    view.textarea.value = "once";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    const first = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    view.textarea.dispatchEvent(first);
    expect(first.defaultPrevented).toBe(true);

    const second = new KeyboardEvent("keydown", { key: "Enter", cancelable: true });
    view.textarea.dispatchEvent(second);
    expect(second.defaultPrevented).toBe(false);

    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith("once", { reason: "enter", shift: false });
    expect(view.model.isEditing).toBe(false);

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

    const e = new KeyboardEvent("keydown", { key: "Enter", altKey: true, cancelable: true });
    view.textarea.dispatchEvent(e);

    // Alt+Enter inserts a newline (and editor indentation) rather than committing the edit.
    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(true);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(true);
    expect(view.textarea.value).toBe("line1\n");
    expect(view.model.draft).toBe("line1\n");
    expect(cancel.hidden).toBe(false);
    expect(cancel.disabled).toBe(false);
    expect(commit.hidden).toBe(false);
    expect(commit.disabled).toBe(false);

    host.remove();
  });

  it("does not cancel twice if Escape is pressed multiple times", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onCancel });

    view.setActiveCell({ address: "A1", input: "original", value: null });
    view.textarea.focus();
    view.textarea.value = "changed";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    const first = new KeyboardEvent("keydown", { key: "Escape", cancelable: true });
    view.textarea.dispatchEvent(first);
    expect(first.defaultPrevented).toBe(true);

    const second = new KeyboardEvent("keydown", { key: "Escape", cancelable: true });
    view.textarea.dispatchEvent(second);
    expect(second.defaultPrevented).toBe(false);

    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(view.model.isEditing).toBe(false);

    host.remove();
  });

  it("commits on Shift+Enter and forwards the shift modifier", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });

    view.textarea.focus();
    view.textarea.value = "shift-enter";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    const e = new KeyboardEvent("keydown", { key: "Enter", shiftKey: true, cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(onCommit).toHaveBeenCalledWith("shift-enter", { reason: "enter", shift: true });
    expect(view.model.isEditing).toBe(false);
    expect(view.model.activeCell.input).toBe("shift-enter");

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

    const e = new KeyboardEvent("keydown", { key: "Escape", cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(view.model.isEditing).toBe(false);
    expect(view.model.draft).toBe("original");
    expect(view.model.activeCell.input).toBe("original");
    expect(document.activeElement).not.toBe(view.textarea);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(false);
    expect(view.textarea.value).toBe("original");
    expect(cancel.hidden).toBe(true);
    expect(cancel.disabled).toBe(true);
    expect(commit.hidden).toBe(true);
    expect(commit.disabled).toBe(true);

    host.remove();
  });

  it("cancels cleanly even when onCancel is not provided", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const view = new FormulaBarView(host, { onCommit });
    view.setActiveCell({ address: "A1", input: "original", value: null });

    view.textarea.focus();
    view.textarea.value = "changed";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));

    const e = new KeyboardEvent("keydown", { key: "Escape", cancelable: true });
    view.textarea.dispatchEvent(e);

    expect(e.defaultPrevented).toBe(true);
    expect(onCommit).not.toHaveBeenCalled();
    expect(view.model.isEditing).toBe(false);
    expect(view.model.draft).toBe("original");
    expect(view.textarea.value).toBe("original");

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
    expect(view.model.draft).toBe("start");
    expect(view.model.activeCell.input).toBe("start");
    expect(document.activeElement).not.toBe(view.textarea);
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
    expect(onCommit).toHaveBeenCalledWith("commit-me", { reason: "command", shift: false });
    expect(onCancel).toHaveBeenCalledTimes(1);
    expect(view.model.isEditing).toBe(false);
    expect(view.model.draft).toBe("commit-me");
    expect(view.model.activeCell.input).toBe("commit-me");
    expect(document.activeElement).not.toBe(view.textarea);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(false);
    expect(cancel.hidden).toBe(true);
    expect(cancel.disabled).toBe(true);
    expect(commit.hidden).toBe(true);
    expect(commit.disabled).toBe(true);

    host.remove();
  });
});
