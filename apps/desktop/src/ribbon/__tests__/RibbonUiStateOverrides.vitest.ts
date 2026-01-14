// @vitest-environment jsdom
import React, { act } from "react";
import { createRoot } from "react-dom/client";
import { afterEach, describe, expect, it, vi } from "vitest";

import { Ribbon } from "../Ribbon";
import { setRibbonUiState } from "../ribbonUiState";
import { RIBBON_DISABLED_BY_ID_WHILE_EDITING } from "../ribbonEditingDisabledById";

afterEach(() => {
  act(() => {
    setRibbonUiState({
      pressedById: Object.create(null),
      labelById: Object.create(null),
      disabledById: Object.create(null),
      shortcutById: Object.create(null),
      ariaKeyShortcutsById: Object.create(null),
    });
  });
  document.body.innerHTML = "";
  vi.restoreAllMocks();
});

function renderRibbon() {
  (globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  act(() => {
    root.render(React.createElement(Ribbon, { actions: {} }));
  });
  return { container, root };
}

describe("Ribbon UI state overrides", () => {
  it("updates toggle aria-pressed when pressed overrides change", () => {
    const { container, root } = renderRibbon();
    const bold = container.querySelector<HTMLButtonElement>('[data-command-id="format.toggleBold"]');
    expect(bold).toBeInstanceOf(HTMLButtonElement);
    expect(bold?.getAttribute("aria-pressed")).toBe("false");

    act(() => {
      setRibbonUiState({
        pressedById: { "format.toggleBold": true },
        labelById: Object.create(null),
        disabledById: Object.create(null),
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    expect(bold?.getAttribute("aria-pressed")).toBe("true");
    act(() => root.unmount());
  });

  it("updates number-format dropdown label via label overrides", () => {
    const { container, root } = renderRibbon();
    const numberFormat = container.querySelector<HTMLButtonElement>('[data-command-id="home.number.numberFormat"]');
    expect(numberFormat).toBeInstanceOf(HTMLButtonElement);

    const labelSpan = () => numberFormat?.querySelector(".ribbon-button__label")?.textContent?.trim() ?? "";
    expect(labelSpan()).toBe("General");

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: { "home.number.numberFormat": "Percent" },
        disabledById: Object.create(null),
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    expect(labelSpan()).toBe("Percent");
    act(() => root.unmount());
  });

  it("disables individual dropdown menu items via disabledById overrides", async () => {
    const { container, root } = renderRibbon();

    const clearFormatting = container.querySelector<HTMLButtonElement>('[data-command-id="home.font.clearFormatting"]');
    expect(clearFormatting).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      clearFormatting?.click();
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });

    const menu = container.querySelector<HTMLElement>(".ribbon-dropdown__menu");
    expect(menu).toBeInstanceOf(HTMLElement);
    if (!menu) throw new Error("Missing dropdown menu");

    const menuItemId = "format.clearFormats";
    const clearFormats = menu.querySelector<HTMLButtonElement>(`[data-command-id="${menuItemId}"]`);
    expect(clearFormats).toBeInstanceOf(HTMLButtonElement);
    expect(clearFormats?.disabled).toBe(false);
    expect(clearFormats?.hasAttribute("disabled")).toBe(false);

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: Object.create(null),
        disabledById: { [menuItemId]: true },
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    const updated = menu.querySelector<HTMLButtonElement>(`[data-command-id="${menuItemId}"]`);
    expect(updated).toBeInstanceOf(HTMLButtonElement);
    expect(updated?.disabled).toBe(true);
    expect(updated?.hasAttribute("disabled")).toBe(true);
    expect(menu.querySelector(`.ribbon-dropdown__menuitem:not(:disabled)[data-command-id="${menuItemId}"]`)).toBeNull();

    act(() => root.unmount());
  });

  it("includes shortcut hints in the button title when provided", () => {
    const { container, root } = renderRibbon();
    const copy = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.copy"]');
    expect(copy).toBeInstanceOf(HTMLButtonElement);
    expect(copy?.getAttribute("aria-label")).toBe("Copy");
    expect(copy?.getAttribute("title")).toBe("Copy");

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: Object.create(null),
        disabledById: Object.create(null),
        shortcutById: { "clipboard.copy": "Ctrl+C" },
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    expect(copy?.getAttribute("aria-label")).toBe("Copy");
    expect(copy?.getAttribute("title")).toBe("Copy (Ctrl+C)");
    act(() => root.unmount());
  });

  it("sets aria-keyshortcuts when ariaKeyShortcuts overrides change", () => {
    const { container, root } = renderRibbon();
    const paste = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.paste"]');
    expect(paste).toBeInstanceOf(HTMLButtonElement);
    expect(paste?.getAttribute("aria-keyshortcuts")).toBeNull();

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: Object.create(null),
        disabledById: Object.create(null),
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: {
          "clipboard.paste": "Control+V",
          "clipboard.pasteSpecial.values": "Alt+V",
        },
      });
    });

    expect(paste?.getAttribute("aria-keyshortcuts")).toBe("Control+V");

    act(() => {
      paste?.click();
    });

    const menuItem = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.pasteSpecial.values"]');
    expect(menuItem).toBeInstanceOf(HTMLButtonElement);
    expect(menuItem?.getAttribute("aria-keyshortcuts")).toBe("Alt+V");

    act(() => root.unmount());
  });

  it("includes expected editing-mode disabled command ids", () => {
    const expected = [
      "home.cells.format.rowHeight",
      "home.cells.format.columnWidth",
      "home.cells.insert.insertCells",
      "home.cells.insert.insertSheetRows",
      "home.cells.insert.insertSheetColumns",
      "home.cells.insert.insertSheet",
      "home.cells.delete.deleteCells",
      "home.cells.delete.deleteSheetRows",
      "home.cells.delete.deleteSheetColumns",
      "home.cells.delete.deleteSheet",
      "home.alignment.mergeCenter",
      "home.alignment.mergeCenter.mergeCenter",
      "home.alignment.mergeCenter.mergeAcross",
      "home.alignment.mergeCenter.mergeCells",
      "home.alignment.mergeCenter.unmergeCells",
      "home.number.moreFormats.custom",
      "home.font.subscript",
      "home.font.superscript",
      "home.editing.sortFilter.sortAtoZ",
      "home.editing.sortFilter.sortZtoA",
      "home.editing.sortFilter.customSort",
      "home.editing.sortFilter.filter",
      "home.editing.sortFilter.clear",
      "home.editing.sortFilter.reapply",
      "home.editing.clear",
      "home.editing.clear.clearAll",
      "home.editing.clear.clearFormats",
      "home.editing.clear.clearContents",
      "data.sortFilter.sortAtoZ",
      "data.sortFilter.sortZtoA",
      "data.sortFilter.sort",
      "data.sortFilter.sort.customSort",
      "data.sortFilter.sort.sortAtoZ",
      "data.sortFilter.sort.sortZtoA",
      "data.sortFilter.filter",
      "data.sortFilter.clear",
      "data.sortFilter.reapply",
      "data.sortFilter.advanced",
      "data.sortFilter.advanced.advancedFilter",
      "data.sortFilter.advanced.clearFilter",
    ];

    for (const id of expected) {
      expect(RIBBON_DISABLED_BY_ID_WHILE_EDITING[id]).toBe(true);
    }
  });

  it("applies disabledById overrides to dropdown menu items", () => {
    const { container, root } = renderRibbon();

    const formatDropdown = container.querySelector<HTMLButtonElement>('[data-command-id="home.cells.format"]');
    expect(formatDropdown).toBeInstanceOf(HTMLButtonElement);

    act(() => {
      formatDropdown?.click();
    });

    const rowHeightItemSelector = '[data-command-id="home.cells.format.rowHeight"]';
    const columnWidthItemSelector = '[data-command-id="home.cells.format.columnWidth"]';

    expect(container.querySelector<HTMLButtonElement>(rowHeightItemSelector)).toBeInstanceOf(HTMLButtonElement);
    expect(container.querySelector<HTMLButtonElement>(columnWidthItemSelector)).toBeInstanceOf(HTMLButtonElement);
    expect(container.querySelector<HTMLButtonElement>(rowHeightItemSelector)?.disabled).toBe(false);
    expect(container.querySelector<HTMLButtonElement>(columnWidthItemSelector)?.disabled).toBe(false);

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: Object.create(null),
        disabledById: {
          "home.cells.format.rowHeight": true,
          "home.cells.format.columnWidth": true,
        },
        shortcutById: Object.create(null),
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    expect(container.querySelector<HTMLButtonElement>(rowHeightItemSelector)?.disabled).toBe(true);
    expect(container.querySelector<HTMLButtonElement>(columnWidthItemSelector)?.disabled).toBe(true);

    act(() => root.unmount());
  });
});
