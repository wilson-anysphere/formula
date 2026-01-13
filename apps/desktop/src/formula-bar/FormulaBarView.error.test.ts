/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

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
});

