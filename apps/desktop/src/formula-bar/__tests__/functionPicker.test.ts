/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { FormulaBarView } from "../FormulaBarView.js";

describe("FormulaBarView fx function picker", () => {
  it("opens the function picker when clicking fx", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]');
    expect(fxButton).toBeTruthy();

    fxButton!.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    const picker = host.querySelector<HTMLElement>('[data-testid="formula-function-picker"]');
    const pickerInput = host.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]');

    expect(picker).toBeTruthy();
    expect(picker?.hidden).toBe(false);
    expect(pickerInput).toBeTruthy();
    expect(document.activeElement).toBe(pickerInput);

    host.remove();
  });

  it("inserts the selected function into the formula bar at the cursor", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]')!;
    fxButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    const pickerInput = host.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]')!;
    pickerInput.value = "sum";
    pickerInput.dispatchEvent(new Event("input", { bubbles: true }));

    pickerInput.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

    expect(view.textarea.value).toBe("=SUM()");
    expect(view.textarea.selectionStart).toBe(view.textarea.value.length - 1);
    expect(view.textarea.selectionEnd).toBe(view.textarea.value.length - 1);
    expect(document.activeElement).toBe(view.textarea);

    host.remove();
  });

  it("navigates function results with arrow keys", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]')!;
    fxButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    const pickerInput = host.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]')!;
    pickerInput.value = "sum";
    pickerInput.dispatchEvent(new Event("input", { bubbles: true }));

    const items = Array.from(
      host.querySelectorAll<HTMLLIElement>('[data-testid^="formula-function-picker-item-"]'),
    );
    expect(items.length).toBeGreaterThan(1);

    const secondName = items[1]!.dataset.testid!.replace("formula-function-picker-item-", "");

    // Move selection to the second result.
    pickerInput.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true, cancelable: true }));
    expect(items[0]!.getAttribute("aria-selected")).toBe("false");
    expect(items[1]!.getAttribute("aria-selected")).toBe("true");

    // Insert selected function.
    pickerInput.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

    expect(view.textarea.value).toBe(`=${secondName}()`);
    expect(view.textarea.selectionStart).toBe(view.textarea.value.length - 1);
    expect(view.textarea.selectionEnd).toBe(view.textarea.value.length - 1);

    host.remove();
  });

  it("inserts the function at the cursor when the draft is not empty", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "=1+", value: null });

    const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]')!;
    fxButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    const pickerInput = host.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]')!;
    pickerInput.value = "sum";
    pickerInput.dispatchEvent(new Event("input", { bubbles: true }));

    pickerInput.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

    expect(view.textarea.value).toBe("=1+SUM(");
    expect(view.textarea.selectionStart).toBe(view.textarea.value.length);
    expect(view.textarea.selectionEnd).toBe(view.textarea.value.length);

    host.remove();
  });

  it("closes on Escape and restores focus to the formula input", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]')!;
    fxButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    const picker = host.querySelector<HTMLElement>('[data-testid="formula-function-picker"]')!;
    const pickerInput = host.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]')!;
    expect(picker.hidden).toBe(false);
    expect(document.activeElement).toBe(pickerInput);

    pickerInput.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));

    expect(picker.hidden).toBe(true);
    expect(document.activeElement).toBe(view.textarea);
    expect(view.textarea.value).toBe("");

    host.remove();
  });
});
