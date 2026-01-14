/**
 * @vitest-environment jsdom
 */

import { describe, expect, it, vi } from "vitest";

import { FormulaBarView } from "./FormulaBarView.js";

describe("FormulaBarView name box dropdown", () => {
  it("opens dropdown and keyboard-selects an item", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    let view: FormulaBarView | null = null;
    const onGoTo = vi.fn((reference: string) => {
      if (reference === "Table1[#All]" && view) {
        // SpreadsheetApp updates the FormulaBarView selection synchronously during navigation.
        view.setActiveCell({ address: "A1", input: "", value: "", nameBox: "A1:D10" });
      }
      return true;
    });
    const provider = {
      getItems: () => [
        {
          kind: "namedRange",
          key: "namedRange:SalesData",
          label: "SalesData",
          reference: "SalesData",
          description: "Sheet1!A1:B2",
        },
        {
          kind: "table",
          key: "table:Table1",
          label: "Table1",
          reference: "Table1[#All]",
          description: "Sheet1!A1:D10",
        },
      ],
    };

    view = new FormulaBarView(host, { onCommit: () => {}, onGoTo }, { nameBoxDropdownProvider: provider });

    const address = host.querySelector<HTMLInputElement>('[data-testid="formula-address"]');
    const dropdown = host.querySelector<HTMLButtonElement>('[data-testid="name-box-dropdown"]');
    const popup = host.querySelector<HTMLDivElement>('[data-testid="formula-name-box-popup"]');
    const list = host.querySelector<HTMLDivElement>('[data-testid="formula-name-box-list"]');

    expect(address).toBeTruthy();
    expect(dropdown).toBeTruthy();
    expect(popup).toBeTruthy();
    expect(list).toBeTruthy();

    dropdown!.dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(popup!.hidden).toBe(false);
    expect(address!.getAttribute("aria-expanded")).toBe("true");
    expect(address!.getAttribute("aria-controls")).toBe(list!.id);

    // Default selection should be the first item.
    let active = list!.querySelector<HTMLElement>('[role="option"][aria-selected="true"]');
    expect(active?.textContent).toContain("SalesData");

    address!.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowDown", bubbles: true, cancelable: true }));
    active = list!.querySelector<HTMLElement>('[role="option"][aria-selected="true"]');
    expect(active?.textContent).toContain("Table1");

    address!.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

    expect(onGoTo).toHaveBeenCalledTimes(1);
    expect(onGoTo).toHaveBeenCalledWith("Table1[#All]");

    expect(address!.value).toBe("A1:D10");
    expect(popup!.hidden).toBe(true);
    expect(address!.getAttribute("aria-expanded")).toBe("false");

    host.remove();
  });

  it("filters items by typing and Escape restores the previous value", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onGoTo = vi.fn();
    const provider = {
      getItems: () => [
        { kind: "namedRange", key: "namedRange:SalesData", label: "SalesData", reference: "SalesData" },
        { kind: "namedRange", key: "namedRange:Costs", label: "Costs", reference: "Costs" },
      ],
    };

    new FormulaBarView(host, { onCommit: () => {}, onGoTo }, { nameBoxDropdownProvider: provider });

    const address = host.querySelector<HTMLInputElement>('[data-testid="formula-address"]')!;
    const dropdown = host.querySelector<HTMLButtonElement>('[data-testid="name-box-dropdown"]')!;
    const popup = host.querySelector<HTMLDivElement>('[data-testid="formula-name-box-popup"]')!;
    const list = host.querySelector<HTMLDivElement>('[data-testid="formula-name-box-list"]')!;

    address.value = "A1";
    dropdown.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(popup.hidden).toBe(false);

    // Type a prefix to filter down to a single item.
    address.value = "Sal";
    address.dispatchEvent(new Event("input", { bubbles: true }));
    const options = list.querySelectorAll<HTMLElement>('[role="option"]');
    expect(options.length).toBe(1);
    expect(options[0]?.textContent).toContain("SalesData");

    // Escape should cancel the dropdown and restore the original address.
    address.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
    expect(popup.hidden).toBe(true);
    expect(address.value).toBe("A1");

    host.remove();
  });

  it("closes on Tab without navigating", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onGoTo = vi.fn(() => true);
    const provider = {
      getItems: () => [{ kind: "namedRange", key: "namedRange:SalesData", label: "SalesData", reference: "SalesData" }],
    };

    new FormulaBarView(host, { onCommit: () => {}, onGoTo }, { nameBoxDropdownProvider: provider });

    const address = host.querySelector<HTMLInputElement>('[data-testid="formula-address"]')!;
    const dropdown = host.querySelector<HTMLButtonElement>('[data-testid="name-box-dropdown"]')!;
    const popup = host.querySelector<HTMLDivElement>('[data-testid="formula-name-box-popup"]')!;

    address.value = "A1";
    dropdown.click();
    expect(popup.hidden).toBe(false);

    // Tab should dismiss the dropdown but not trigger navigation.
    address.dispatchEvent(new KeyboardEvent("keydown", { key: "Tab", bubbles: true, cancelable: true }));
    expect(popup.hidden).toBe(true);
    expect(onGoTo).not.toHaveBeenCalled();
    expect(address.value).toBe("A1");

    host.remove();
  });

  it("closes the dropdown when focus moves outside the name box", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onGoTo = vi.fn(() => true);
    const provider = {
      getItems: () => [{ kind: "namedRange", key: "namedRange:SalesData", label: "SalesData", reference: "SalesData" }],
    };

    new FormulaBarView(host, { onCommit: () => {}, onGoTo }, { nameBoxDropdownProvider: provider });

    const dropdown = host.querySelector<HTMLButtonElement>('[data-testid="name-box-dropdown"]')!;
    const popup = host.querySelector<HTMLDivElement>('[data-testid="formula-name-box-popup"]')!;

    dropdown.click();
    expect(popup.hidden).toBe(false);

    const outside = document.createElement("button");
    document.body.appendChild(outside);
    outside.focus();

    expect(popup.hidden).toBe(true);

    host.remove();
    outside.remove();
  });

  it("shows an empty-state message when the workbook has no names/tables", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onGoTo = vi.fn();
    const provider = {
      getItems: () => [],
    };

    new FormulaBarView(host, { onCommit: () => {}, onGoTo }, { nameBoxDropdownProvider: provider });

    const dropdown = host.querySelector<HTMLButtonElement>('[data-testid="name-box-dropdown"]')!;
    const popup = host.querySelector<HTMLDivElement>('[data-testid="formula-name-box-popup"]')!;
    const list = host.querySelector<HTMLDivElement>('[data-testid="formula-name-box-list"]')!;

    dropdown.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(popup.hidden).toBe(false);
    expect(list.textContent).toContain("No named ranges");
    const emptyOption = list.querySelector<HTMLElement>('[role="option"][aria-disabled="true"]');
    expect(emptyOption).toBeTruthy();
    expect(emptyOption?.textContent).toContain("No named ranges");

    // Escape should close without navigating.
    const address = host.querySelector<HTMLInputElement>('[data-testid="formula-address"]')!;
    address.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true }));
    expect(popup.hidden).toBe(true);
    expect(onGoTo).not.toHaveBeenCalled();

    host.remove();
  });

  it("keeps focus in the name box when selecting an item without a navigation reference", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onGoTo = vi.fn();
    const provider = {
      getItems: () => [{ kind: "namedRange", key: "namedRange:Const", label: "Const", reference: "" }],
    };

    new FormulaBarView(host, { onCommit: () => {}, onGoTo }, { nameBoxDropdownProvider: provider });

    const address = host.querySelector<HTMLInputElement>('[data-testid="formula-address"]')!;
    const dropdown = host.querySelector<HTMLButtonElement>('[data-testid="name-box-dropdown"]')!;
    const popup = host.querySelector<HTMLDivElement>('[data-testid="formula-name-box-popup"]')!;

    dropdown.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(popup.hidden).toBe(false);

    // Select the only item.
    address.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));

    expect(onGoTo).not.toHaveBeenCalled();
    expect(popup.hidden).toBe(true);
    expect(document.activeElement).toBe(address);
    expect(address.value).toBe("Const");
    expect(address.selectionStart).toBe(0);
    expect(address.selectionEnd).toBe(5);

    host.remove();
  });

  it("promotes the most recently selected item into a Recent section", () => {
    const host = document.createElement("div");
    document.body.appendChild(host);

    const onGoTo = vi.fn(() => true);
    const provider = {
      getItems: () => [
        { kind: "namedRange", key: "namedRange:SalesData", label: "SalesData", reference: "SalesData" },
        { kind: "namedRange", key: "namedRange:Costs", label: "Costs", reference: "Costs" },
      ],
    };

    new FormulaBarView(host, { onCommit: () => {}, onGoTo }, { nameBoxDropdownProvider: provider });

    const address = host.querySelector<HTMLInputElement>('[data-testid="formula-address"]')!;
    const dropdown = host.querySelector<HTMLButtonElement>('[data-testid="name-box-dropdown"]')!;
    const popup = host.querySelector<HTMLDivElement>('[data-testid="formula-name-box-popup"]')!;
    const list = host.querySelector<HTMLDivElement>('[data-testid="formula-name-box-list"]')!;

    dropdown.click();
    const initialActive = list.querySelector<HTMLElement>('[role="option"][aria-selected="true"]');
    const initialLabel = (initialActive?.textContent ?? "").trim();
    address.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
    expect(onGoTo).toHaveBeenCalledTimes(1);
    const selectedRef = onGoTo.mock.calls[0]?.[0];
    expect(typeof selectedRef).toBe("string");
    expect(initialLabel).toContain(String(selectedRef));
    expect(popup.hidden).toBe(true);

    // Reopen: the most recently selected item should be surfaced under "Recent" at the top.
    dropdown.click();
    expect(popup.hidden).toBe(false);

    const headings = Array.from(list.querySelectorAll<HTMLElement>(".formula-bar-name-box-group-label")).map((el) => el.textContent);
    expect(headings[0]).toBe("Recent");

    const active = list.querySelector<HTMLElement>('[role="option"][aria-selected="true"]');
    expect(active?.textContent).toContain(String(selectedRef));

    host.remove();
  });
});
