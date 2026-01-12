import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("command palette: panel toggles", () => {
  test("toggles AI Chat panel open/closed", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);

    const primary = process.platform === "darwin" ? "Meta" : "Control";

    // Open via command palette.
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("AI Chat");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toBeVisible();

    // Close via command palette.
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("AI Chat");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("panel-aiChat")).toHaveCount(0);
  });

  test("toggles Marketplace panel open/closed", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);

    const primary = process.platform === "darwin" ? "Meta" : "Control";

    // Open via command palette.
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Marketplace");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("panel-marketplace")).toBeVisible();

    // Close via command palette.
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Marketplace");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("panel-marketplace")).toHaveCount(0);
  });
});
