/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView fx function picker", () => {
  it("does not open when read-only (but may focus the textarea for copy)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onBeginEdit = vi.fn();
    const view = new FormulaBarView(host, { onCommit: () => {}, onBeginEdit });
    view.setActiveCell({ address: "A1", input: "=SUM(A1)", value: null });
    view.setReadOnly(true);

    const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]');
    expect(fxButton).toBeTruthy();

    const cancel = host.querySelector<HTMLButtonElement>(".formula-bar-action-button--cancel");
    const commit = host.querySelector<HTMLButtonElement>(".formula-bar-action-button--commit");
    expect(cancel).toBeTruthy();
    expect(commit).toBeTruthy();
    expect(cancel!.hidden).toBe(true);
    expect(commit!.hidden).toBe(true);

    fxButton!.click();

    const picker = host.querySelector<HTMLElement>('[data-testid="formula-function-picker"]');
    expect(picker).toBeTruthy();

    expect(view.model.isEditing).toBe(false);
    expect(onBeginEdit).not.toHaveBeenCalled();
    expect(cancel!.hidden).toBe(true);
    expect(cancel!.disabled).toBe(true);
    expect(commit!.hidden).toBe(true);
    expect(commit!.disabled).toBe(true);
    expect(picker!.hidden).toBe(true);
    expect(fxButton!.getAttribute("aria-expanded")).toBe("false");

    // Read-only mode still allows focusing/selecting the formula bar for copy.
    expect(document.activeElement).toBe(view.textarea);
    expect(view.root.classList.contains("formula-bar--editing")).toBe(true);

    host.remove();
  });

  it("opens, filters, and inserts the selected function call", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]');
    expect(fxButton).toBeTruthy();

    fxButton!.click();

    const picker = host.querySelector<HTMLElement>('[data-testid="formula-function-picker"]');
    expect(picker).toBeTruthy();
    expect(picker!.hidden).toBe(false);

    const pickerInput = picker!.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]');
    expect(pickerInput).toBeTruthy();
    expect(document.activeElement).toBe(pickerInput);

    pickerInput!.value = "sum";
    pickerInput!.dispatchEvent(new Event("input", { bubbles: true }));

    pickerInput!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", cancelable: true, bubbles: true }));

    expect(view.textarea.value).toBe("=SUM()");
    expect(view.textarea.selectionStart).toBe(view.textarea.value.length - 1);
    expect(view.textarea.selectionEnd).toBe(view.textarea.value.length - 1);
    expect(document.activeElement).toBe(view.textarea);
    expect(picker!.hidden).toBe(true);

    host.remove();
  });

  it("closes if the formula bar becomes read-only while the picker is open", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCommit = vi.fn();
    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit, onCancel });
    view.setActiveCell({ address: "A1", input: "orig", value: null });

    const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]');
    expect(fxButton).toBeTruthy();
    fxButton!.click();

    const picker = host.querySelector<HTMLElement>('[data-testid="formula-function-picker"]');
    expect(picker).toBeTruthy();
    expect(picker!.hidden).toBe(false);
    expect(fxButton!.getAttribute("aria-expanded")).toBe("true");

    // Edit the draft while the picker is open.
    view.textarea.value = "changed";
    view.textarea.setSelectionRange(view.textarea.value.length, view.textarea.value.length);
    view.textarea.dispatchEvent(new Event("input"));
    expect(view.model.draft).toBe("changed");

    view.setReadOnly(true);

    expect(view.model.isEditing).toBe(false);
    expect(view.model.draft).toBe("orig");
    expect(view.textarea.value).toBe("orig");
    expect(picker!.hidden).toBe(true);
    expect(fxButton!.getAttribute("aria-expanded")).toBe("false");
    expect(onCommit).not.toHaveBeenCalled();
    expect(onCancel).not.toHaveBeenCalled();

    host.remove();
  });

  it("does not insert the selected function on Enter during IME composition", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]');
    expect(fxButton).toBeTruthy();
    fxButton!.click();

    const picker = host.querySelector<HTMLElement>('[data-testid="formula-function-picker"]');
    expect(picker).toBeTruthy();
    expect(picker!.hidden).toBe(false);

    const pickerInput = picker!.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]');
    expect(pickerInput).toBeTruthy();

    pickerInput!.value = "sum";
    pickerInput!.dispatchEvent(new Event("input", { bubbles: true }));

    pickerInput!.dispatchEvent(new Event("compositionstart", { bubbles: true }));
    const enterDuringComposition = new KeyboardEvent("keydown", { key: "Enter", cancelable: true, bubbles: true });
    pickerInput!.dispatchEvent(enterDuringComposition);

    expect(enterDuringComposition.defaultPrevented).toBe(false);
    expect(picker!.hidden).toBe(false);
    expect(view.textarea.value).toBe("");

    pickerInput!.dispatchEvent(new Event("compositionend", { bubbles: true }));
    const enterAfterComposition = new KeyboardEvent("keydown", { key: "Enter", cancelable: true, bubbles: true });
    pickerInput!.dispatchEvent(enterAfterComposition);

    expect(enterAfterComposition.defaultPrevented).toBe(true);
    expect(view.textarea.value).toBe("=SUM()");
    expect(picker!.hidden).toBe(true);

    host.remove();
  });
});
