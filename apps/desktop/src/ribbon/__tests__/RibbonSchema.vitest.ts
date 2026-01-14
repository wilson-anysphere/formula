import { describe, expect, it } from "vitest";

import { defaultRibbonSchema } from "../ribbonSchema";

function expectNonEmpty(value: string, label: string) {
  expect(value, `${label} should be present`).toBeTruthy();
  expect(value.trim().length, `${label} should be non-empty`).toBeGreaterThan(0);
}

function expectUniqueIds(ids: string[], label: string) {
  const seen = new Set<string>();
  for (const id of ids) {
    expectNonEmpty(id, `${label} id`);
    expect(seen.has(id), `${label} contains duplicate id: ${id}`).toBe(false);
    seen.add(id);
  }
}

describe("defaultRibbonSchema", () => {
  it("includes the expected Excel-style tabs", () => {
    const tabIds = defaultRibbonSchema.tabs.map((tab) => tab.id);
    const required = ["file", "home", "insert", "pageLayout", "formulas", "data", "review", "view", "developer", "help"];
    for (const id of required) {
      expect(tabIds, `Expected schema to include tab '${id}'`).toContain(id);
    }
  });

  it("exposes panel toggles in the View tab with stable test ids", () => {
    const viewTab = defaultRibbonSchema.tabs.find((tab) => tab.id === "view");
    expect(viewTab, "Expected View tab to exist").toBeTruthy();
    if (!viewTab) return;

    const panelsGroup = viewTab.groups.find((group) => group.id === "view.panels");
    expect(panelsGroup, "Expected View → Panels group to exist").toBeTruthy();
    if (!panelsGroup) return;

    const marketplace = panelsGroup.buttons.find((button) => button.id === "view.togglePanel.marketplace");
    expect(marketplace, "Expected view.togglePanel.marketplace button").toBeTruthy();
    expect(marketplace?.testId).toBe("open-marketplace-panel");
    const versionHistory = panelsGroup.buttons.find((button) => button.id === "view.togglePanel.versionHistory");
    expect(versionHistory, "Expected view.togglePanel.versionHistory button").toBeTruthy();
    expect(versionHistory?.testId).toBe("open-version-history-panel");

    const branchManager = panelsGroup.buttons.find((button) => button.id === "view.togglePanel.branchManager");
    expect(branchManager, "Expected view.togglePanel.branchManager button").toBeTruthy();
    expect(branchManager?.testId).toBe("open-branch-manager-panel");
  });

  it("exposes Page Layout print/export controls with stable test ids", () => {
    const pageLayoutTab = defaultRibbonSchema.tabs.find((tab) => tab.id === "pageLayout");
    expect(pageLayoutTab, "Expected Page Layout tab to exist").toBeTruthy();
    if (!pageLayoutTab) return;

    const pageSetupGroup = pageLayoutTab.groups.find((group) => group.id === "pageLayout.pageSetup");
    expect(pageSetupGroup, "Expected Page Layout → Page Setup group to exist").toBeTruthy();
    if (!pageSetupGroup) return;

    const pageSetupButton = pageSetupGroup.buttons.find((button) => button.id === "pageLayout.pageSetup.pageSetupDialog");
    expect(pageSetupButton, "Expected Page Setup dialog button to exist").toBeTruthy();
    expect(pageSetupButton?.testId).toBe("ribbon-page-setup");

    const printAreaGroup = pageLayoutTab.groups.find((group) => group.id === "pageLayout.printArea");
    expect(printAreaGroup, "Expected Page Layout → Print Area group to exist").toBeTruthy();
    if (!printAreaGroup) return;

    const setPrintArea = printAreaGroup.buttons.find((button) => button.id === "pageLayout.printArea.setPrintArea");
    expect(setPrintArea, "Expected Set Print Area button").toBeTruthy();
    expect(setPrintArea?.testId).toBe("ribbon-set-print-area");

    const clearPrintArea = printAreaGroup.buttons.find((button) => button.id === "pageLayout.printArea.clearPrintArea");
    expect(clearPrintArea, "Expected Clear Print Area button").toBeTruthy();
    expect(clearPrintArea?.testId).toBe("ribbon-clear-print-area");

    const exportGroup = pageLayoutTab.groups.find((group) => group.id === "pageLayout.export");
    expect(exportGroup, "Expected Page Layout → Export group to exist").toBeTruthy();
    if (!exportGroup) return;

    const exportPdf = exportGroup.buttons.find((button) => button.id === "pageLayout.export.exportPdf");
    expect(exportPdf, "Expected Export to PDF button").toBeTruthy();
    expect(exportPdf?.testId).toBe("ribbon-export-pdf");
  });

  it("includes Home → Font subscript/superscript toggles", () => {
    const homeTab = defaultRibbonSchema.tabs.find((tab) => tab.id === "home");
    expect(homeTab, "Expected Home tab to exist").toBeTruthy();
    if (!homeTab) return;

    const fontGroup = homeTab.groups.find((group) => group.id === "home.font");
    expect(fontGroup, "Expected Home → Font group to exist").toBeTruthy();
    if (!fontGroup) return;

    const subscript = fontGroup.buttons.find((button) => button.id === "home.font.subscript");
    expect(subscript, "Expected Home → Font → Subscript button to exist").toBeTruthy();
    expect(subscript?.kind).toBe("toggle");
    expect(subscript?.disabled).not.toBe(true);

    const superscript = fontGroup.buttons.find((button) => button.id === "home.font.superscript");
    expect(superscript, "Expected Home → Font → Superscript button to exist").toBeTruthy();
    expect(superscript?.kind).toBe("toggle");
    expect(superscript?.disabled).not.toBe(true);
  });

  it("keeps File → Info → Manage Workbook menu item ids wired", () => {
    const fileTab = defaultRibbonSchema.tabs.find((tab) => tab.id === "file");
    expect(fileTab, "Expected File tab to exist").toBeTruthy();
    if (!fileTab) return;

    const infoGroup = fileTab.groups.find((group) => group.id === "file.info");
    expect(infoGroup, "Expected File → Info group to exist").toBeTruthy();
    if (!infoGroup) return;

    const manageWorkbook = infoGroup.buttons.find((button) => button.id === "file.info.manageWorkbook");
    expect(manageWorkbook, "Expected File → Info → Manage Workbook dropdown").toBeTruthy();
    expect(manageWorkbook?.kind).toBe("dropdown");
    expect(
      manageWorkbook?.menuItems?.some((item) => item.id === "file.info.manageWorkbook.versions"),
      "Expected Manage Workbook dropdown to include file.info.manageWorkbook.versions",
    ).toBe(true);
    expect(
      manageWorkbook?.menuItems?.some((item) => item.id === "file.info.manageWorkbook.branches"),
      "Expected Manage Workbook dropdown to include file.info.manageWorkbook.branches",
    ).toBe(true);
  });

  it("wires Insert → PivotTable to the CommandRegistry command id (view.insertPivotTable)", () => {
    const insertTab = defaultRibbonSchema.tabs.find((tab) => tab.id === "insert");
    expect(insertTab, "Expected Insert tab to exist").toBeTruthy();
    if (!insertTab) return;

    const tablesGroup = insertTab.groups.find((group) => group.id === "insert.tables");
    expect(tablesGroup, "Expected Insert → Tables group to exist").toBeTruthy();
    if (!tablesGroup) return;

    const pivot = tablesGroup.buttons.find((button) => button.testId === "ribbon-insert-pivot-table");
    expect(pivot, "Expected PivotTable button to exist").toBeTruthy();
    expect(pivot?.id).toBe("view.insertPivotTable");
    expect(pivot?.menuItems?.[0]?.id).toBe("view.insertPivotTable");
    expect(pivot?.menuItems?.[1]?.id).toBe("insert.tables.pivotTable.fromTableRange");
    expect(pivot?.menuItems?.[1]?.disabled).not.toBe(true);
    expect(pivot?.menuItems?.[2]?.disabled).toBe(true);
    expect(pivot?.menuItems?.[3]?.disabled).toBe(true);
  });

  it("does not include legacy icon properties in the schema", () => {
    const offenders: string[] = [];
    for (const tab of defaultRibbonSchema.tabs) {
      for (const group of tab.groups) {
        for (const button of group.buttons) {
          if ("icon" in (button as any)) offenders.push(`button:${button.id}`);
          for (const item of button.menuItems ?? []) {
            if ("icon" in (item as any)) offenders.push(`menuItem:${item.id}`);
          }
        }
      }
    }
    expect(offenders, `Found legacy icon fields: ${offenders.join(", ")}`).toEqual([]);
  });

  it("ensures sibling ids are unique (tabs, groups, buttons, menu items)", () => {
    expectUniqueIds(
      defaultRibbonSchema.tabs.map((tab) => tab.id),
      "tab",
    );

    for (const tab of defaultRibbonSchema.tabs) {
      expectUniqueIds(
        tab.groups.map((group) => group.id),
        `group (tab: ${tab.id})`,
      );

      for (const group of tab.groups) {
        expectUniqueIds(
          // A ribbon group may expose multiple controls that execute the same command id
          // (e.g. legacy variants that keep stable test hooks). Mirror the runtime React
          // key selection (testId when available, else command id) so we still catch
          // collisions that would break rendering.
          group.buttons.map((button) => button.testId ?? button.id),
          `button (tab: ${tab.id}, group: ${group.id})`,
        );

        for (const button of group.buttons) {
          if (!button.menuItems || button.menuItems.length === 0) continue;
          expectUniqueIds(
            button.menuItems.map((item) => item.id),
            `menu item (button: ${button.id})`,
          );
        }
      }
    }
  });

  it("ensures labels + aria-labels are present for all schema items", () => {
    for (const tab of defaultRibbonSchema.tabs) {
      expectNonEmpty(tab.id, "tab.id");
      expectNonEmpty(tab.label, `tab.label (${tab.id})`);
      expect(tab.groups.length, `tab.groups (${tab.id}) should not be empty`).toBeGreaterThan(0);

      for (const group of tab.groups) {
        expectNonEmpty(group.id, `group.id (${tab.id})`);
        expectNonEmpty(group.label, `group.label (${group.id})`);
        expect(group.buttons.length, `group.buttons (${group.id}) should not be empty`).toBeGreaterThan(0);

        for (const button of group.buttons) {
          expectNonEmpty(button.id, `button.id (${group.id})`);
          expectNonEmpty(button.label, `button.label (${button.id})`);
          expectNonEmpty(button.ariaLabel, `button.ariaLabel (${button.id})`);

          for (const item of button.menuItems ?? []) {
            expectNonEmpty(item.id, `menuItem.id (${button.id})`);
            expectNonEmpty(item.label, `menuItem.label (${item.id})`);
            expectNonEmpty(item.ariaLabel, `menuItem.ariaLabel (${item.id})`);
          }
        }
      }
    }
  });

  it("aligns Home → Editing AutoSum/Fill ids with builtin command ids", () => {
    const homeTab = defaultRibbonSchema.tabs.find((tab) => tab.id === "home");
    expect(homeTab, "Expected Home tab to exist").toBeTruthy();
    if (!homeTab) return;

    const editingGroup = homeTab.groups.find((group) => group.id === "home.editing");
    expect(editingGroup, "Expected Home → Editing group to exist").toBeTruthy();
    if (!editingGroup) return;

    const autoSum = editingGroup.buttons.find((button) => button.ariaLabel === "AutoSum");
    expect(autoSum?.id).toBe("edit.autoSum");
    expect(
      autoSum?.menuItems?.some((item) => item.id === "edit.autoSum"),
      "Expected AutoSum dropdown to include edit.autoSum",
    ).toBe(true);
    expect(
      autoSum?.menuItems?.some((item) => item.id === "home.editing.autoSum.sum"),
      "Expected AutoSum dropdown not to include legacy home.editing.autoSum.sum",
    ).toBe(false);

    const fill = editingGroup.buttons.find((button) => button.id === "home.editing.fill");
    expect(fill, "Expected Fill dropdown to exist").toBeTruthy();
    expect(
      fill?.menuItems?.some((item) => item.id === "edit.fillDown"),
      "Expected Fill dropdown to include edit.fillDown",
    ).toBe(true);
    expect(
      fill?.menuItems?.some((item) => item.id === "edit.fillRight"),
      "Expected Fill dropdown to include edit.fillRight",
    ).toBe(true);
    expect(
      fill?.menuItems?.some((item) => item.id === "home.editing.fill.down"),
      "Expected Fill dropdown not to include legacy home.editing.fill.down",
    ).toBe(false);
    expect(
      fill?.menuItems?.some((item) => item.id === "home.editing.fill.right"),
      "Expected Fill dropdown not to include legacy home.editing.fill.right",
    ).toBe(false);
  });
  it("aligns Home → Find & Select ids with builtin command ids", () => {
    const homeTab = defaultRibbonSchema.tabs.find((tab) => tab.id === "home");
    expect(homeTab, "Expected Home tab to exist").toBeTruthy();
    if (!homeTab) return;

    const findGroup = homeTab.groups.find((group) => group.id === "home.find");
    expect(findGroup, "Expected Home → Find group to exist").toBeTruthy();
    if (!findGroup) return;

    const findSelect = findGroup.buttons.find((button) => button.id === "home.editing.findSelect");
    expect(findSelect, "Expected Find & Select dropdown to exist").toBeTruthy();
    expect(findSelect?.kind).toBe("dropdown");
    expect(findSelect?.menuItems?.map((item) => item.id)).toEqual(["edit.find", "edit.replace", "navigation.goTo"]);

    expect(findSelect?.menuItems?.find((item) => item.id === "edit.find")?.testId).toBe("ribbon-find");
    expect(findSelect?.menuItems?.find((item) => item.id === "edit.replace")?.testId).toBe("ribbon-replace");
    expect(findSelect?.menuItems?.find((item) => item.id === "navigation.goTo")?.testId).toBe("ribbon-goto");

    expect(
      findSelect?.menuItems?.some((item) => item.id === "home.editing.findSelect.find"),
      "Expected Find & Select dropdown not to include legacy home.editing.findSelect.find",
    ).toBe(false);
    expect(
      findSelect?.menuItems?.some((item) => item.id === "home.editing.findSelect.replace"),
      "Expected Find & Select dropdown not to include legacy home.editing.findSelect.replace",
    ).toBe(false);
    expect(
      findSelect?.menuItems?.some((item) => item.id === "home.editing.findSelect.goTo"),
      "Expected Find & Select dropdown not to include legacy home.editing.findSelect.goTo",
    ).toBe(false);
  });

  it("aligns View → Window → Freeze Panes menu ids with builtin command ids", () => {
    const viewTab = defaultRibbonSchema.tabs.find((tab) => tab.id === "view");
    expect(viewTab, "Expected View tab to exist").toBeTruthy();
    if (!viewTab) return;

    const windowGroup = viewTab.groups.find((group) => group.id === "view.window");
    expect(windowGroup, "Expected View → Window group to exist").toBeTruthy();
    if (!windowGroup) return;

    const freezePanes = windowGroup.buttons.find((button) => button.id === "view.window.freezePanes");
    expect(freezePanes, "Expected Freeze Panes dropdown to exist").toBeTruthy();
    if (!freezePanes) return;

    expect(freezePanes.kind).toBe("dropdown");
    const ids = freezePanes.menuItems?.map((item) => item.id) ?? [];

    expect(ids).toEqual(
      expect.arrayContaining(["view.freezePanes", "view.freezeTopRow", "view.freezeFirstColumn", "view.unfreezePanes"]),
    );

    // Ensure we don't regress to the old hierarchical ids.
    expect(ids).not.toContain("view.window.freezePanes.freezePanes");
    expect(ids).not.toContain("view.window.freezePanes.freezeTopRow");
    expect(ids).not.toContain("view.window.freezePanes.freezeFirstColumn");
    expect(ids).not.toContain("view.window.freezePanes.unfreeze");
  });
});
