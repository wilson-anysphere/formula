/**
 * @vitest-environment jsdom
 */

import { beforeEach, describe, expect, it } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

const STORAGE_KEY = "formula:ui:formulaBarExpanded";

function queryExpandButton(host: HTMLElement): HTMLButtonElement {
  const button = host.querySelector<HTMLButtonElement>('[data-testid="formula-expand-button"]');
  if (!button) {
    throw new Error("Expected formula expand button to exist");
  }
  return button;
}

describe("FormulaBarView expand/collapse", () => {
  beforeEach(() => {
    try {
      window.sessionStorage.removeItem(STORAGE_KEY);
    } catch {
      // ignore
    }
  });

  it("toggles a CSS class on the root element when clicked", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    const button = queryExpandButton(host);

    expect(view.root.classList.contains("formula-bar--expanded")).toBe(false);
    expect(button.getAttribute("aria-label")).toBe("Expand formula bar");

    button.click();
    expect(view.root.classList.contains("formula-bar--expanded")).toBe(true);
    expect(button.getAttribute("aria-label")).toBe("Collapse formula bar");

    button.click();
    expect(view.root.classList.contains("formula-bar--expanded")).toBe(false);
    expect(button.getAttribute("aria-label")).toBe("Expand formula bar");

    host.remove();
  });

  it("uses a larger max height when expanded", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const view = new FormulaBarView(host, { onCommit: () => {} });
    const button = queryExpandButton(host);

    // Force the editor to "want" to grow beyond the default max height.
    Object.defineProperty(view.textarea, "scrollHeight", {
      configurable: true,
      get: () => 1000,
    });

    view.focus({ cursor: "end" });
    expect(view.textarea.style.height).toBe("140px");

    button.click();
    expect(view.textarea.style.height).toBe("360px");

    button.click();
    expect(view.textarea.style.height).toBe("140px");

    host.remove();
  });
});
