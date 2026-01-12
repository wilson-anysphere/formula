import { expect, test } from "@playwright/test";

test("switching sheets restores grid focus for immediate keyboard editing", async ({ page }) => {
  await page.goto("/?e2e=1");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const sheetSelect = page.getByTestId("sheet-switcher");
  await sheetSelect.selectOption("Sheet2");

  const grid = page.getByTestId("canvas-grid");
  await expect(grid).toBeVisible();
  await expect(grid).toBeFocused();

  await page.keyboard.press("F2");
  await expect(page.getByTestId("cell-editor")).toBeVisible();
});
