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

