import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("shared grid Excel-scale navigation", () => {
  test("can jump to row 1,000,000 and column XFD", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => (window as any).__formulaApp.whenIdle());

    await page.evaluate(() => {
      (window as any).__formulaApp.activateCell({ row: 999_999, col: 0 });
      (window as any).__formulaApp.focus();
    });
    await expect(page.getByTestId("active-cell")).toHaveText("A1000000");

    await page.evaluate(() => {
      (window as any).__formulaApp.activateCell({ row: 0, col: 16_383 });
      (window as any).__formulaApp.focus();
    });
    await expect(page.getByTestId("active-cell")).toHaveText("XFD1");
  });
});

