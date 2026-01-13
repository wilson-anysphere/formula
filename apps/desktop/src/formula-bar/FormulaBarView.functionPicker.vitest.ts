/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView fx function picker", () => {
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
