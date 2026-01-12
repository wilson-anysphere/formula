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

