import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function getActiveCell(page: import("@playwright/test").Page): Promise<{ row: number; col: number }> {
  return page.evaluate(() => (window as any).__formulaApp.getActiveCell());
}

test.describe("keyboard navigation: Tab grid traversal + F6 focus cycling", () => {
  const GRID_MODES = ["legacy", "shared"] as const;

  for (const mode of GRID_MODES) {
    test(`Tab/Shift+Tab moves selection and wraps (${mode})`, async ({ page }) => {
      const url = mode === "legacy" ? "/?grid=legacy&maxRows=3&maxCols=3" : "/?grid=shared";
      await gotoDesktop(page, url);

      const limits = await page.evaluate(() => {
        const app = (window as any).__formulaApp as any;
        const limits = app.limits as { maxRows: number; maxCols: number };
        app.activateCell({ row: 0, col: limits.maxCols - 1 });
        app.focus();
        return limits;
      });

      await expect(page.locator("#grid")).toBeFocused();

      // End-of-row Tab wraps to the first column of the next row.
      await page.keyboard.press("Tab");
      await expect.poll(async () => await getActiveCell(page)).toEqual({ row: 1, col: 0 });
      await expect(page.locator("#grid")).toBeFocused();

      // Shift+Tab at the start of the row wraps to the last column of the previous row.
      await page.keyboard.press("Shift+Tab");
      await expect.poll(async () => await getActiveCell(page)).toEqual({ row: 0, col: limits.maxCols - 1 });
      await expect(page.locator("#grid")).toBeFocused();
    });
  }

  test("F6 / Shift+F6 cycles focus across ribbon, formula bar, sheet tabs, and grid", async ({ page }) => {
    await gotoDesktop(page, "/?grid=legacy");

    const ribbonRoot = page.getByTestId("ribbon-root");
    await expect(ribbonRoot).toBeVisible();

    await page.evaluate(() => (window as any).__formulaApp.focus());
    await expect(page.locator("#grid")).toBeFocused();

    const activeRibbonTab = ribbonRoot.locator('[role="tab"][aria-selected="true"]');
    const sheetTabsActiveTab = page.locator('#sheet-tabs button[role="tab"][tabindex="0"]');

    // Forward cycle: grid -> ribbon -> formula bar -> sheet tabs -> grid.
    await page.keyboard.press("F6");
    await expect(activeRibbonTab).toBeFocused();

    await page.keyboard.press("F6");
    await expect(page.getByTestId("formula-address")).toBeFocused();

    await page.keyboard.press("F6");
    await expect(sheetTabsActiveTab).toBeFocused();

    await page.keyboard.press("F6");
    await expect(page.locator("#grid")).toBeFocused();

    // Reverse cycle: grid -> sheet tabs -> formula bar -> ribbon -> grid.
    await page.keyboard.press("Shift+F6");
    await expect(sheetTabsActiveTab).toBeFocused();

    await page.keyboard.press("Shift+F6");
    await expect(page.getByTestId("formula-address")).toBeFocused();

    await page.keyboard.press("Shift+F6");
    await expect(activeRibbonTab).toBeFocused();

    await page.keyboard.press("Shift+F6");
    await expect(page.locator("#grid")).toBeFocused();
  });
});

