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

  it("shows shortcut hints for dropdown menu items when shortcutById overrides change", async () => {
    const { container, root } = renderRibbon();

    const paste = container.querySelector<HTMLButtonElement>('[data-command-id="clipboard.paste"]');
    expect(paste).toBeInstanceOf(HTMLButtonElement);

    await act(async () => {
      paste?.click();
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
    });

    const menuItemId = "clipboard.pasteSpecial.values";
    const menuItem = container.querySelector<HTMLButtonElement>(`[data-command-id="${menuItemId}"]`);
    expect(menuItem).toBeInstanceOf(HTMLButtonElement);
    expect(menuItem?.dataset.shortcut).toBeUndefined();
    expect(menuItem?.getAttribute("title")).toBe("Paste Values");

    act(() => {
      setRibbonUiState({
        pressedById: Object.create(null),
        labelById: Object.create(null),
        disabledById: Object.create(null),
        shortcutById: { [menuItemId]: "Alt+V" },
        ariaKeyShortcutsById: Object.create(null),
      });
    });

    const updated = container.querySelector<HTMLButtonElement>(`[data-command-id="${menuItemId}"]`);
    expect(updated).toBeInstanceOf(HTMLButtonElement);
    expect(updated?.dataset.shortcut).toBe("Alt+V");
    expect(updated?.getAttribute("title")).toBe("Paste Values (Alt+V)");

    act(() => root.unmount());
  });

  it("includes expected editing-mode disabled command ids", () => {
    const expected = [
      "view.toggleShowFormulas",
      "view.window.freezePanes",
      "view.freezePanes",
      "view.freezeTopRow",
      "view.freezeFirstColumn",
      "view.unfreezePanes",
      "audit.togglePrecedents",
      "audit.toggleDependents",
      "audit.toggleTransitive",
      "formulas.formulaAuditing.tracePrecedents",
      "formulas.formulaAuditing.traceDependents",
      "formulas.formulaAuditing.removeArrows",
      "view.macros.viewMacros",
      "view.macros.viewMacros.run",
      "view.macros.viewMacros.edit",
      "view.macros.viewMacros.delete",
      "view.macros.recordMacro",
      "view.macros.recordMacro.stop",
      "view.macros.useRelativeReferences",
      "developer.code.visualBasic",
      "developer.code.macros",
      "developer.code.macros.run",
      "developer.code.macros.edit",
      "developer.code.recordMacro",
      "developer.code.recordMacro.stop",
      "developer.code.useRelativeReferences",
      "developer.code.macroSecurity",
      "developer.code.macroSecurity.trustCenter",
      "ai.inlineEdit",
      "clipboard.copy",
      "clipboard.cut",
      "clipboard.paste",
      "clipboard.pasteSpecial",
      "clipboard.pasteSpecial.values",
      "clipboard.pasteSpecial.formulas",
      "clipboard.pasteSpecial.formats",
      "clipboard.pasteSpecial.transpose",
      "home.editing.findSelect",
      "edit.find",
      "edit.replace",
      "navigation.goTo",
      "home.cells.insert",
      "home.cells.insert.insertCells",
      "home.cells.insert.insertSheetRows",
      "home.cells.insert.insertSheetColumns",
      "home.cells.insert.insertSheet",
      "home.cells.delete",
      "home.cells.delete.deleteCells",
      "home.cells.delete.deleteSheetRows",
      "home.cells.delete.deleteSheetColumns",
      "home.cells.delete.deleteSheet",
      "home.cells.format",
      "home.cells.format.rowHeight",
      "home.cells.format.columnWidth",
      "home.cells.format.organizeSheets",
      "home.alignment.mergeCenter",
      "home.alignment.mergeCenter.mergeCenter",
      "home.alignment.mergeCenter.mergeAcross",
      "home.alignment.mergeCenter.mergeCells",
      "home.alignment.mergeCenter.unmergeCells",
      "home.number.moreFormats.custom",
      "home.styles.formatAsTable",
      "home.styles.formatAsTable.light",
      "home.styles.formatAsTable.medium",
      "home.styles.formatAsTable.dark",
      "home.styles.formatAsTable.newStyle",
      "home.styles.cellStyles",
      "home.styles.cellStyles.goodBadNeutral",
      "home.styles.cellStyles.dataModel",
      "home.styles.cellStyles.titlesHeadings",
      "home.styles.cellStyles.numberFormat",
      "home.styles.cellStyles.newStyle",
      "view.insertPivotTable",
      "insert.tables.pivotTable.fromTableRange",
      "insert.illustrations.pictures",
      "insert.illustrations.pictures.thisDevice",
      "insert.illustrations.pictures.stockImages",
      "insert.illustrations.pictures.onlinePictures",
      "insert.illustrations.onlinePictures",
      "data.queriesConnections.refreshAll",
      "data.queriesConnections.refreshAll.refresh",
      "data.queriesConnections.refreshAll.refreshAllConnections",
      "data.queriesConnections.refreshAll.refreshAllQueries",
      "data.forecast.whatIfAnalysis",
      "data.forecast.whatIfAnalysis.scenarioManager",
      "data.forecast.whatIfAnalysis.goalSeek",
      "data.forecast.whatIfAnalysis.monteCarlo",
      "data.forecast.whatIfAnalysis.dataTable",
      "formulas.solutions.solver",
      "format.toggleSubscript",
      "format.toggleSuperscript",
      "edit.autoSum",
      "home.editing.autoSum.average",
      "home.editing.autoSum.countNumbers",
      "home.editing.autoSum.max",
      "home.editing.autoSum.min",
      "home.editing.fill",
      "edit.fillDown",
      "edit.fillRight",
      "edit.fillUp",
      "edit.fillLeft",
      "home.editing.fill.series",
      "home.editing.sortFilter",
      "home.editing.sortFilter.customSort",
      "home.editing.clear",
      "format.clearAll",
      "format.clearFormats",
      "edit.clearContents",
      "home.editing.clear.clearComments",
      "home.editing.clear.clearHyperlinks",
      "data.sortFilter.sortAtoZ",
      "data.sortFilter.sortZtoA",
      "data.sortFilter.sort",
      "data.sortFilter.sort.customSort",
      "data.sortFilter.filter",
      "data.sortFilter.clear",
      "data.sortFilter.reapply",
      "data.sortFilter.advanced",
      "data.sortFilter.advanced.advancedFilter",
      "data.sortFilter.advanced.clearFilter",
      "pageLayout.pageSetup.pageSetupDialog",
      "pageLayout.pageSetup.margins",
      "pageLayout.pageSetup.margins.normal",
      "pageLayout.pageSetup.margins.wide",
      "pageLayout.pageSetup.margins.narrow",
      "pageLayout.pageSetup.margins.custom",
      "pageLayout.pageSetup.orientation",
      "pageLayout.pageSetup.orientation.portrait",
      "pageLayout.pageSetup.orientation.landscape",
      "pageLayout.pageSetup.size",
      "pageLayout.pageSetup.size.letter",
      "pageLayout.pageSetup.size.a4",
      "pageLayout.pageSetup.size.more",
      "pageLayout.printArea.setPrintArea",
      "pageLayout.printArea.clearPrintArea",
      "pageLayout.pageSetup.printArea",
      "pageLayout.pageSetup.printArea.set",
      "pageLayout.pageSetup.printArea.clear",
      "pageLayout.pageSetup.printArea.addTo",
      "pageLayout.arrange.bringForward",
      "pageLayout.arrange.sendBackward",
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
