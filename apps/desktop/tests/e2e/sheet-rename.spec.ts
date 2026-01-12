import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("sheet rename", () => {
  test("double-clicking a tab renames the sheet and marks the document dirty", async ({ page }) => {
    await gotoDesktop(page);

    // Create Sheet2 and reset the dirty state so the rename is what marks the document dirty.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "x");
      app.getDocument().markSaved();
    });

    const sheet2Tab = page.getByTestId("sheet-tab-Sheet2");
    await expect(sheet2Tab).toBeVisible();

    await sheet2Tab.dblclick();

    const input = sheet2Tab.locator("input.sheet-tab__input");
    await expect(input).toBeVisible();
    await input.fill("Data");
    await input.press("Enter");

    // The stable sheet id stays the same (data-testid uses the sheet id), but the display label changes.
    await expect(sheet2Tab.locator(".sheet-tab__name")).toHaveText("Data");
    await expect(sheet2Tab).toHaveAttribute("data-testid", "sheet-tab-Sheet2");

    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty))
      .toBe(true);

    await page.getByTestId("sheet-add").click();
    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty))
      .toBe(true);
  });
});
