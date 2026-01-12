import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("command palette (Audit toggles)", () => {
  test("shows keybinding hints and toggles auditing mode", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Toggle precedents.
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Toggle Trace Precedents");

    const expectedPrecedentsShortcut = process.platform === "darwin" ? "⌘[" : "Ctrl+[";
    const precedentsItem = page
      .getByTestId("command-palette-list")
      .locator(".command-palette__item")
      .filter({ hasText: "Toggle Trace Precedents" })
      .first();
    await expect(precedentsItem).toBeVisible();
    await expect(precedentsItem.locator(".command-palette__shortcut")).toHaveText(expectedPrecedentsShortcut);
    await precedentsItem.click();
    await expect(page.getByTestId("command-palette")).toBeHidden();

    await page.evaluate(async () => (window as any).__formulaApp.whenIdle());
    const afterPrecedents = await page.evaluate(() => (window as any).__formulaApp.getAuditingHighlights());
    expect(afterPrecedents.mode).toBe("precedents");

    // Toggle dependents (should become BOTH).
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Toggle Trace Dependents");

    const expectedDependentsShortcut = process.platform === "darwin" ? "⌘]" : "Ctrl+]";
    const dependentsItem = page
      .getByTestId("command-palette-list")
      .locator(".command-palette__item")
      .filter({ hasText: "Toggle Trace Dependents" })
      .first();
    await expect(dependentsItem).toBeVisible();
    await expect(dependentsItem.locator(".command-palette__shortcut")).toHaveText(expectedDependentsShortcut);
    await dependentsItem.click();
    await expect(page.getByTestId("command-palette")).toBeHidden();

    await page.evaluate(async () => (window as any).__formulaApp.whenIdle());
    const afterDependents = await page.evaluate(() => (window as any).__formulaApp.getAuditingHighlights());
    expect(afterDependents.mode).toBe("both");
  });
});

