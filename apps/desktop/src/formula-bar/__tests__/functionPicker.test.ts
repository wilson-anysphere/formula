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
    expect(fxButton!.getAttribute("aria-expanded")).toBe("true");

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

  it("renders a signature/description row for function results", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "", value: null });

    const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]')!;
    fxButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    const pickerInput = host.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]')!;
    pickerInput.value = "sum";
    pickerInput.dispatchEvent(new Event("input", { bubbles: true }));

    const sumItem = host.querySelector<HTMLElement>('[data-testid="formula-function-picker-item-SUM"]');
    expect(sumItem).toBeTruthy();

    const desc = sumItem!.querySelector<HTMLElement>(".command-palette__item-description");
    expect(desc).toBeTruthy();
    expect(desc!.textContent).toContain("SUM(");

    host.remove();
  });

  it("uses the locale list separator in signature previews (de-DE uses ';')", () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "de-DE";

    const host = document.createElement("div");
    document.body.appendChild(host);

    try {
      const view = new FormulaBarView(host, { onCommit: () => {} });
      view.setActiveCell({ address: "A1", input: "", value: null });

      const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]')!;
      fxButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));

      const pickerInput = host.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]')!;
      pickerInput.value = "sum";
      pickerInput.dispatchEvent(new Event("input", { bubbles: true }));

      // In de-DE, search results should prefer the localized name.
      const sumItem = host.querySelector<HTMLElement>('[data-testid="formula-function-picker-item-SUMME"]');
      expect(sumItem).toBeTruthy();

      const desc = sumItem!.querySelector<HTMLElement>(".command-palette__item-description");
      expect(desc).toBeTruthy();
      expect(desc!.textContent).toContain("SUMME(");
      expect(desc!.textContent).toContain(";");
    } finally {
      host.remove();
      document.documentElement.lang = prevLang;
    }
  });

  it("searches + inserts localized function names when the UI locale uses localized formulas (de-DE SUMME → SUMME())", () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "de-DE";

    const host = document.createElement("div");
    document.body.appendChild(host);

    try {
      const view = new FormulaBarView(host, { onCommit: () => {} });
      view.setActiveCell({ address: "A1", input: "", value: null });

      const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]')!;
      fxButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));

      const pickerInput = host.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]')!;
      pickerInput.value = "summe";
      pickerInput.dispatchEvent(new Event("input", { bubbles: true }));

      const summeItem = host.querySelector<HTMLElement>('[data-testid="formula-function-picker-item-SUMME"]');
      expect(summeItem).toBeTruthy();

      pickerInput.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

      expect(view.textarea.value).toBe("=SUMME()");
      expect(view.textarea.selectionStart).toBe(view.textarea.value.length - 1);
      expect(view.textarea.selectionEnd).toBe(view.textarea.value.length - 1);
    } finally {
      host.remove();
      document.documentElement.lang = prevLang;
    }
  });

  it("uses getLocaleId() to localize results (even if document.lang differs)", () => {
    const prevLang = document.documentElement.lang;
    document.documentElement.lang = "en-US";

    const host = document.createElement("div");
    document.body.appendChild(host);

    try {
      const view = new FormulaBarView(host, { onCommit: () => {} }, { getLocaleId: () => "de-DE" });
      view.setActiveCell({ address: "A1", input: "", value: null });

      const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]')!;
      fxButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));

      // Empty-query list is localized (default/common functions).
      expect(host.querySelector<HTMLElement>('[data-testid="formula-function-picker-item-SUMME"]')).toBeTruthy();

      // Non-empty search also uses the override locale.
      const pickerInput = host.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]')!;
      pickerInput.value = "zähl";
      pickerInput.dispatchEvent(new Event("input", { bubbles: true }));

      expect(host.querySelector<HTMLElement>('[data-testid="formula-function-picker-item-ZÄHLENWENN"]')).toBeTruthy();
    } finally {
      host.remove();
      document.documentElement.lang = prevLang;
    }
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
    expect(pickerInput.getAttribute("aria-activedescendant")).toBe(items[1]!.id);

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

    expect(view.textarea.value).toBe("=1+SUM()");
    // Caret should land inside the parens: `=1+SUM(|)`.
    expect(view.textarea.selectionStart).toBe(view.textarea.value.length - 1);
    expect(view.textarea.selectionEnd).toBe(view.textarea.value.length - 1);

    host.remove();
  });

  it("replaces the current selection when inserting a function", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    view.setActiveCell({ address: "A1", input: "=1+2", value: null });

    // Start editing and select the trailing "2".
    view.focus({ cursor: "end" });
    view.textarea.setSelectionRange(3, 4);
    view.textarea.dispatchEvent(new Event("select"));

    const fxButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-fx-button"]')!;
    // Preserve the textarea selection (the fx button's handler relies on mousedown preventDefault).
    fxButton.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true }));
    fxButton.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    const pickerInput = host.querySelector<HTMLInputElement>('[data-testid="formula-function-picker-input"]')!;
    pickerInput.value = "sum";
    pickerInput.dispatchEvent(new Event("input", { bubbles: true }));

    pickerInput.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

    expect(view.textarea.value).toBe("=1+SUM()");
    // Caret should land inside the parens: `=1+SUM(|)`.
    expect(view.textarea.selectionStart).toBe(view.textarea.value.length - 1);
    expect(view.textarea.selectionEnd).toBe(view.textarea.value.length - 1);

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
    expect(fxButton.getAttribute("aria-expanded")).toBe("false");
    expect(view.textarea.value).toBe("");

    host.remove();
  });
});
