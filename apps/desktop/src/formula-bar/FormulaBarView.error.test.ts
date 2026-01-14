/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";
import { FORMULA_REFERENCE_PALETTE } from "@formula/spreadsheet-frontend";

describe("FormulaBarView error panel", () => {
  it("only shows the error button when the active cell has an error explanation", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });

    const errorButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-button"]');
    expect(errorButton).not.toBeNull();
    expect(view.model.errorExplanation()).toBeNull();
    expect(errorButton?.hidden).toBe(true);

    view.setActiveCell({ address: "A1", input: "=1/0", value: "#DIV/0!" });

    expect(view.model.errorExplanation()).not.toBeNull();
    expect(errorButton?.hidden).toBe(false);

    host.remove();
  });

  it("opens the panel with actionable buttons and sets aria-expanded", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    new FormulaBarView(host, { onCommit: () => {} }).setActiveCell({ address: "A1", input: "=1/0", value: "#DIV/0!" });

    const errorButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-button"]')!;
    expect(errorButton.hidden).toBe(false);
    expect(errorButton.getAttribute("aria-expanded")).toBe("false");

    errorButton.click();

    expect(errorButton.getAttribute("aria-expanded")).toBe("true");
    const panel = host.querySelector<HTMLElement>('[data-testid="formula-error-panel"]')!;
    expect(panel.hidden).toBe(false);
    expect(panel.querySelector('[data-testid="formula-error-fix-ai"]')).not.toBeNull();
    expect(panel.querySelector('[data-testid="formula-error-show-ranges"]')).not.toBeNull();

    host.remove();
  });

  it("invokes onFixFormulaErrorWithAi when clicking Fix with AI", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onFixFormulaErrorWithAi = vi.fn();
    const view = new FormulaBarView(host, { onCommit: () => {}, onFixFormulaErrorWithAi });
    view.setActiveCell({ address: "A1", input: "=1/0", value: "#DIV/0!" });

    const errorButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-button"]')!;
    errorButton.click();

    const fixButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-fix-ai"]')!;
    fixButton.click();

    expect(onFixFormulaErrorWithAi).toHaveBeenCalledTimes(1);
    expect(onFixFormulaErrorWithAi.mock.calls[0]?.[0]).toMatchObject({
      address: "A1",
      draft: "=1/0",
      explanation: { code: "#DIV/0!" },
    });

    host.remove();
  });

  it("resolves named ranges when showing referenced ranges from the error panel", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let lastHighlights: any[] = [];
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onReferenceHighlights: (highlights) => {
        lastHighlights = highlights;
      },
    });

    view.model.setNameResolver((name) =>
      name === "SalesData" ? { sheet: "Sheet1", startRow: 0, startCol: 0, endRow: 1, endCol: 1 } : null
    );
    view.setActiveCell({ address: "A1", input: "=SUM(SalesData)", value: "#DIV/0!" });

    const errorButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-button"]')!;
    errorButton.click();

    const showRanges = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-show-ranges"]')!;
    showRanges.click();

    expect(lastHighlights).toEqual([
      {
        range: { sheet: "Sheet1", startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
        color: FORMULA_REFERENCE_PALETTE[0],
        text: "SalesData",
        index: 0,
        active: false,
      },
    ]);

    host.remove();
  });

  it("closes on Escape and restores focus to the error button", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCancel = vi.fn();
    new FormulaBarView(host, { onCommit: () => {}, onCancel }).setActiveCell({ address: "A1", input: "=1/0", value: "#DIV/0!" });

    const errorButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-button"]')!;
    errorButton.click();

    const panel = host.querySelector<HTMLElement>('[data-testid="formula-error-panel"]')!;
    expect(panel.hidden).toBe(false);

    const esc = new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true });
    panel.dispatchEvent(esc);

    expect(esc.defaultPrevented).toBe(true);
    expect(panel.hidden).toBe(true);
    expect(errorButton.getAttribute("aria-expanded")).toBe("false");
    expect(document.activeElement).toBe(errorButton);
    expect(onCancel).not.toHaveBeenCalled();

    host.remove();
  });

  it("Escape in the error panel does not cancel formula bar editing (it only closes the panel)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onCancel = vi.fn();
    const view = new FormulaBarView(host, { onCommit: () => {}, onCancel });
    view.setActiveCell({ address: "A1", input: "=1/0", value: "#DIV/0!" });

    // Start editing.
    view.textarea.focus();
    expect(view.model.isEditing).toBe(true);

    const errorButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-button"]')!;
    errorButton.click();

    const panel = host.querySelector<HTMLElement>('[data-testid="formula-error-panel"]')!;
    expect(panel.hidden).toBe(false);
    expect(view.model.isEditing).toBe(true);

    const closeButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-close"]')!;
    closeButton.focus();

    const esc = new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true });
    closeButton.dispatchEvent(esc);

    expect(esc.defaultPrevented).toBe(true);
    expect(panel.hidden).toBe(true);
    // Edit mode should remain active; Escape only dismissed the error panel.
    expect(view.model.isEditing).toBe(true);
    expect(onCancel).not.toHaveBeenCalled();

    host.remove();
  });

  it("traps focus with Tab/Shift+Tab inside the error panel (skipping disabled buttons)", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    // No onFixFormulaErrorWithAi callback → Fix button should be disabled.
    new FormulaBarView(host, { onCommit: () => {} }).setActiveCell({ address: "A1", input: "=A1/0", value: "#DIV/0!" });

    const errorButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-button"]')!;
    errorButton.click();

    const panel = host.querySelector<HTMLElement>('[data-testid="formula-error-panel"]')!;
    expect(panel.hidden).toBe(false);

    const fixButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-fix-ai"]')!;
    const showRanges = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-show-ranges"]')!;
    const closeButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-close"]')!;

    expect(fixButton.disabled).toBe(true);
    expect(showRanges.disabled).toBe(false);

    // Opening should focus the first enabled action (Show referenced ranges).
    expect(document.activeElement).toBe(showRanges);

    const tabForward = new KeyboardEvent("keydown", { key: "Tab", bubbles: true, cancelable: true });
    showRanges.dispatchEvent(tabForward);
    expect(tabForward.defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(closeButton);

    const tabForwardWrap = new KeyboardEvent("keydown", { key: "Tab", bubbles: true, cancelable: true });
    closeButton.dispatchEvent(tabForwardWrap);
    expect(tabForwardWrap.defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(showRanges);

    const tabBack = new KeyboardEvent("keydown", { key: "Tab", shiftKey: true, bubbles: true, cancelable: true });
    showRanges.dispatchEvent(tabBack);
    expect(tabBack.defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(closeButton);

    host.remove();
  });

  it("toggles referenced range highlights from the error panel and clears them on close", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let lastHighlights: any[] = [];
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onReferenceHighlights: (highlights) => {
        lastHighlights = highlights;
      },
    });
    view.setActiveCell({ address: "A1", input: "=A1/0", value: "#DIV/0!" });

    const errorButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-button"]')!;
    errorButton.click();

    const panel = host.querySelector<HTMLElement>('[data-testid="formula-error-panel"]')!;
    const showRanges = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-show-ranges"]')!;
    expect(panel.hidden).toBe(false);

    showRanges.click();
    expect(showRanges.getAttribute("aria-pressed")).toBe("true");
    expect(showRanges.textContent).toContain("Hide");
    expect(lastHighlights.length).toBeGreaterThan(0);
    expect(lastHighlights[0]).toMatchObject({
      text: "A1",
      color: FORMULA_REFERENCE_PALETTE[0],
      index: 0,
      active: false,
    });

    // Toggle off.
    showRanges.click();
    expect(showRanges.getAttribute("aria-pressed")).toBe("false");
    expect(showRanges.textContent).toContain("Show");
    expect(lastHighlights).toEqual([]);

    // Toggle on again, then close the panel via Escape; highlights should clear.
    showRanges.click();
    expect(lastHighlights.length).toBeGreaterThan(0);

    const esc = new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true });
    panel.dispatchEvent(esc);
    expect(esc.defaultPrevented).toBe(true);
    expect(panel.hidden).toBe(true);
    expect(lastHighlights).toEqual([]);

    host.remove();
  });

  it("disables referenced range highlighting when the draft is not a formula", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let lastHighlights: any[] = [];
    const view = new FormulaBarView(host, {
      onCommit: () => {},
      onReferenceHighlights: (highlights) => {
        lastHighlights = highlights;
      },
    });
    // Error value, but non-formula input.
    view.setActiveCell({ address: "A1", input: "hello", value: "#DIV/0!" });

    const errorButton = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-button"]')!;
    errorButton.click();

    const showRanges = host.querySelector<HTMLButtonElement>('[data-testid="formula-error-show-ranges"]')!;
    expect(showRanges.disabled).toBe(true);
    expect(showRanges.getAttribute("aria-pressed")).toBe("false");
    expect(showRanges.textContent).toContain("Show");

    // Disabled button should not toggle highlights.
    showRanges.click();
    expect(showRanges.getAttribute("aria-pressed")).toBe("false");
    expect(lastHighlights).toEqual([]);

    host.remove();
  });
});
