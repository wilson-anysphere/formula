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

  it("exposes Version History + Branch Manager panel toggles in the View tab", () => {
    const viewTab = defaultRibbonSchema.tabs.find((tab) => tab.id === "view");
    expect(viewTab, "Expected View tab to exist").toBeTruthy();
    if (!viewTab) return;

    const panelsGroup = viewTab.groups.find((group) => group.id === "view.panels");
    expect(panelsGroup, "Expected View → Panels group to exist").toBeTruthy();
    if (!panelsGroup) return;

    const versionHistory = panelsGroup.buttons.find((button) => button.id === "open-version-history-panel");
    expect(versionHistory, "Expected open-version-history-panel button").toBeTruthy();
    expect(versionHistory?.testId).toBe("open-version-history-panel");

    const branchManager = panelsGroup.buttons.find((button) => button.id === "open-branch-manager-panel");
    expect(branchManager, "Expected open-branch-manager-panel button").toBeTruthy();
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

  it("does not include emoji glyph icons in the schema", () => {
    const emojiRe = /\p{Extended_Pictographic}/u;
    const offenders: string[] = [];
    for (const tab of defaultRibbonSchema.tabs) {
      for (const group of tab.groups) {
        for (const button of group.buttons) {
          if (typeof button.icon === "string" && emojiRe.test(button.icon)) {
            offenders.push(`button:${button.id} (${button.icon})`);
          }
          for (const item of button.menuItems ?? []) {
            if (typeof item.icon === "string" && emojiRe.test(item.icon)) {
              offenders.push(`menuItem:${item.id} (${item.icon})`);
            }
          }
        }
      }
    }
    expect(offenders, `Found emoji icon glyphs: ${offenders.join(", ")}`).toEqual([]);
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
          group.buttons.map((button) => button.id),
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
});
