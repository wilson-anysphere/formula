import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("command palette: find/replace/go to", () => {
  test("Find command opens the find dialog", async ({ page }) => {
    await gotoDesktop(page);

    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    // The palette also includes spreadsheet function results (e.g. `FIND()`), which can
    // outrank the Find *command* for a naive "Find" query. Use the command id to ensure we
    // execute the command rather than inserting a function template.
    await page.getByTestId("command-palette-input").fill("edit.find");
    await page.keyboard.press("Enter");

    await expect(page.locator("dialog.find-replace-dialog[open]")).toBeVisible();
  });

  test("Replace command opens the replace dialog", async ({ page }) => {
    await gotoDesktop(page);

    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    // See note in the Find test: disambiguate from the spreadsheet `REPLACE()` function.
    await page.getByTestId("command-palette-input").fill("edit.replace");
    await page.keyboard.press("Enter");

    const replaceDialog = page.locator("dialog.find-replace-dialog[open]");
    await expect(replaceDialog).toBeVisible();
    await expect(replaceDialog.locator('input[placeholder="Replace with…"]')).toBeVisible();
  });

  test("Replace shortcut opens the replace dialog", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure key events go to the spreadsheet shell.
    await page.evaluate(() => (window as any).__formulaApp.focus());

    // Avoid using Playwright's `keyboard.press()` here since Ctrl+H / Cmd+Option+F may be
    // intercepted by the browser shell (History, toolbar focus, etc.) in some environments.
    // Dispatching a keydown event directly validates our in-app handler logic deterministically.
    await page.evaluate((isMac) => {
      const evt = new KeyboardEvent("keydown", {
        key: isMac ? "f" : "h",
        metaKey: isMac,
        altKey: isMac,
        ctrlKey: !isMac,
        bubbles: true,
        cancelable: true,
      });
      window.dispatchEvent(evt);
    }, process.platform === "darwin");

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
