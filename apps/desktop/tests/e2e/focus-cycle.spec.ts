import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("focus cycling (Excel-style F6)", () => {
  test("F6 / Shift+F6 cycle focus between ribbon, formula bar, grid, sheet tabs, and status bar", async ({ page }) => {
    await gotoDesktop(page);

    // Start from the grid.
    await page.click("#grid", { position: { x: 80, y: 40 } });
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    // Forward cycle (matches apps/desktop/src/main.ts):
    // ribbon -> formula bar -> grid -> sheet tabs -> status bar -> ribbon
    //
    // Starting from the grid, that means: sheet tabs -> status bar -> ribbon -> formula bar -> grid.
    await page.keyboard.press("F6");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeFocused();

    await page.keyboard.press("F6");
    await expect(page.getByTestId("zoom-control")).toBeFocused();

    await page.keyboard.press("F6");
    await expect(page.getByTestId("ribbon-tab-home")).toBeFocused();

    await page.keyboard.press("F6");
    await expect(page.getByTestId("formula-address")).toBeFocused();

    await page.keyboard.press("F6");
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    // One more forward press returns to sheet tabs (wrap).
    await page.keyboard.press("F6");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeFocused();

    // Reverse cycle.
    await page.keyboard.press("Shift+F6");
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    await page.keyboard.press("Shift+F6");
    await expect(page.getByTestId("formula-address")).toBeFocused();

    await page.keyboard.press("Shift+F6");
    await expect(page.getByTestId("ribbon-tab-home")).toBeFocused();

    await page.keyboard.press("Shift+F6");
    await expect(page.getByTestId("zoom-control")).toBeFocused();

    await page.keyboard.press("Shift+F6");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeFocused();
  });
});
