import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("sheet tabs", () => {
  test("switching sheets updates the visible cell values", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure A1 is active before switching sheets so the status bar reflects A1 values.
    await page.click("#grid", { position: { x: 5, y: 5 } });
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();

    // Lazily create Sheet2 by writing a value into it.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    await page.getByTestId("sheet-tab-Sheet2").click();
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("Hello from Sheet2");

    // Switching back restores the original Sheet1 value.
    await page.getByTestId("sheet-tab-Sheet1").click();
    await expect(page.getByTestId("active-value")).toHaveText("Seed");
  });

  test("add sheet button creates and activates the next SheetN tab", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure A1 is active so the status bar is deterministic after the sheet switch.
    await page.click("#grid", { position: { x: 5, y: 5 } });

    const nextSheetId = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const ids = app.getDocument().getSheetIds();
      const existing = new Set((ids.length > 0 ? ids : ["Sheet1"]) as string[]);
      let n = 1;
      while (existing.has(`Sheet${n}`)) n += 1;
      return `Sheet${n}`;
    });

    await page.getByTestId("sheet-add").click();

    const newTab = page.getByTestId(`sheet-tab-${nextSheetId}`);
    await expect(newTab).toBeVisible();
    await expect(newTab).toHaveAttribute("data-active", "true");

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe(nextSheetId);

    // Sheet activation should return focus to the grid so keyboard shortcuts keep working.
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    // Verify the new sheet is functional by writing a value into A1 and observing the status bar update.
    await page.evaluate((sheetId) => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue(sheetId, "A1", `Hello from ${sheetId}`);
    }, nextSheetId);

    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText(`Hello from ${nextSheetId}`);
  });
});
