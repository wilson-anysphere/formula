// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";

import { FormulaBarView } from "../FormulaBarView.js";

afterEach(() => {
  document.body.innerHTML = "";
});

describe("FormulaBarView name box dropdown menu", () => {
  it("opens a menu listing items from the provider when clicking ▾", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    new FormulaBarView(host, {
      onCommit: () => {},
      getNameBoxMenuItems: () => [
        { label: "MyRange", reference: "'My Sheet'!A1:B2" },
        { label: "Table1", reference: "Sheet1!A1:C10" },
      ],
    });

    const dropdown = host.querySelector<HTMLButtonElement>(".formula-bar-name-box-dropdown");
    expect(dropdown).toBeInstanceOf(HTMLButtonElement);

    dropdown!.click();

    const overlay = document.querySelector<HTMLDivElement>('[data-testid="name-box-menu"]');
    expect(overlay).toBeInstanceOf(HTMLDivElement);
    expect(overlay?.hidden).toBe(false);
    expect(dropdown?.getAttribute("aria-expanded")).toBe("true");

    const labels = Array.from(overlay!.querySelectorAll<HTMLElement>(".context-menu__label")).map((el) => el.textContent);
    expect(labels).toContain("MyRange");
    expect(labels).toContain("Table1");

    // Close so subsequent tests don't inherit global window listeners.
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", cancelable: true }));
    expect(overlay?.hidden).toBe(true);
  });

  it("opens via ArrowDown when the name box input is focused and selects via Enter", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onGoTo = vi.fn();
    new FormulaBarView(host, {
      onCommit: () => {},
      onGoTo,
      getNameBoxMenuItems: () => [
        { label: "MyRange", reference: "'My Sheet'!A1:B2" },
        { label: "Other", reference: "A1" },
      ],
    });

    const address = host.querySelector<HTMLInputElement>('[data-testid="formula-address"]');
    expect(address).toBeInstanceOf(HTMLInputElement);
    address!.focus();

    address!.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true, cancelable: true }));

    const overlay = document.querySelector<HTMLDivElement>('[data-testid="name-box-menu"]');
    expect(overlay).toBeInstanceOf(HTMLDivElement);
    expect(overlay?.hidden).toBe(false);

    const firstItem = overlay!.querySelector<HTMLButtonElement>(".context-menu__item:not(:disabled)");
    expect(firstItem).toBeInstanceOf(HTMLButtonElement);
    expect(document.activeElement).toBe(firstItem);

    // Enter should activate the focused item and close the menu.
    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", cancelable: true }));
    expect(onGoTo).toHaveBeenCalledWith("'My Sheet'!A1:B2");
    expect(overlay?.hidden).toBe(true);
  });

  it("navigates via onGoTo when selecting a menu item", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onGoTo = vi.fn();
    new FormulaBarView(host, {
      onCommit: () => {},
      onGoTo,
      getNameBoxMenuItems: () => [{ label: "MyRange", reference: "'My Sheet'!A1:B2" }],
    });

    const dropdown = host.querySelector<HTMLButtonElement>(".formula-bar-name-box-dropdown");
    expect(dropdown).toBeInstanceOf(HTMLButtonElement);

    dropdown!.click();
    const overlay = document.querySelector<HTMLDivElement>('[data-testid="name-box-menu"]');
    expect(overlay).toBeInstanceOf(HTMLDivElement);
    expect(overlay?.hidden).toBe(false);

    const item = Array.from(overlay!.querySelectorAll<HTMLButtonElement>(".context-menu__item")).find(
      (btn) => btn.querySelector(".context-menu__label")?.textContent === "MyRange",
    );
    expect(item).toBeInstanceOf(HTMLButtonElement);

    item!.click();

    expect(onGoTo).toHaveBeenCalledTimes(1);
    expect(onGoTo).toHaveBeenCalledWith("'My Sheet'!A1:B2");
    expect(overlay?.hidden).toBe(true);
  });

  it("closes on Escape and restores focus to the name box input", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    new FormulaBarView(host, {
      onCommit: () => {},
      getNameBoxMenuItems: () => [{ label: "MyRange", reference: "'My Sheet'!A1:B2" }],
    });

    const address = host.querySelector<HTMLInputElement>('[data-testid="formula-address"]');
    expect(address).toBeInstanceOf(HTMLInputElement);

    const dropdown = host.querySelector<HTMLButtonElement>(".formula-bar-name-box-dropdown");
    expect(dropdown).toBeInstanceOf(HTMLButtonElement);

    dropdown!.click();
    const overlay = document.querySelector<HTMLDivElement>('[data-testid="name-box-menu"]');
    expect(overlay).toBeInstanceOf(HTMLDivElement);
    expect(overlay?.hidden).toBe(false);

    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", cancelable: true }));

    expect(overlay?.hidden).toBe(true);
    expect(document.activeElement).toBe(address);
  });

  it("shows a disabled placeholder item when the workbook has no names/tables", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    new FormulaBarView(host, {
      onCommit: () => {},
      getNameBoxMenuItems: () => [],
    });

    const dropdown = host.querySelector<HTMLButtonElement>(".formula-bar-name-box-dropdown");
    expect(dropdown).toBeInstanceOf(HTMLButtonElement);

    dropdown!.click();
    const overlay = document.querySelector<HTMLDivElement>('[data-testid="name-box-menu"]');
    expect(overlay).toBeInstanceOf(HTMLDivElement);
    expect(overlay?.hidden).toBe(false);

    const placeholder = Array.from(overlay!.querySelectorAll<HTMLButtonElement>(".context-menu__item")).find(
      (btn) => btn.querySelector(".context-menu__label")?.textContent === "No named ranges",
    );
    expect(placeholder).toBeInstanceOf(HTMLButtonElement);
    expect(placeholder?.disabled).toBe(true);

    window.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", cancelable: true }));
  });
});
