import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("sheet switcher focus", () => {
  test("selecting a sheet from the switcher restores grid focus for immediate keyboard editing", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure A1 is active before switching sheets so F2 editing is deterministic.
    await page.evaluate(() => {
      const app = window.__formulaApp as any;
      app.activateCell({ row: 0, col: 0 });
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    const switcher = page.getByTestId("sheet-switcher");
    await switcher.selectOption("Sheet2", { force: true });

    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getCurrentSheetId())).toBe("Sheet2");

    // Sheet activation should behave like navigation and leave the grid ready for keyboard input.
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    await page.keyboard.press("F2");
    await expect(page.locator("textarea.cell-editor")).toBeVisible();
  });
});
