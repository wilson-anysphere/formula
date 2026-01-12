import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("command palette: find/replace/go to", () => {
  test("Find command opens the find dialog", async ({ page }) => {
    await gotoDesktop(page);

    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Find");
    await page.keyboard.press("Enter");

    await expect(page.locator("dialog.find-replace-dialog[open]")).toBeVisible();
  });

  test("Replace command opens the replace dialog", async ({ page }) => {
    await gotoDesktop(page);

    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Replace");
    await page.keyboard.press("Enter");

    const replaceDialog = page.locator("dialog.find-replace-dialog[open]");
    await expect(replaceDialog).toBeVisible();
    await expect(replaceDialog.locator('input[placeholder="Replace with…"]')).toBeVisible();
  });

  test("Replace shortcut opens the replace dialog", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure key events go to the spreadsheet shell.
    await page.evaluate(() => (window as any).__formulaApp.focus());

    const replaceShortcut = process.platform === "darwin" ? "Meta+Alt+F" : "Control+H";
    await page.keyboard.press(replaceShortcut);

    const replaceDialog = page.locator("dialog.find-replace-dialog[open]");
    await expect(replaceDialog).toBeVisible();
    await expect(replaceDialog.locator('input[placeholder="Replace with…"]')).toBeVisible();
  });

  test("Go To command opens the go-to dialog", async ({ page }) => {
    await gotoDesktop(page);

    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Go To");
    await page.keyboard.press("Enter");

    await expect(page.locator("dialog.goto-dialog[open]")).toBeVisible();
  });
});
