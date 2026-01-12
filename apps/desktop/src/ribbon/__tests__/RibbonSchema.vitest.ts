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

  it("keeps the File → Info → Manage Workbook → Version History menu item id wired", () => {
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
