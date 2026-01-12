import { expect, test } from "@playwright/test";

import { gotoDesktop, openSheetTabContextMenu } from "./helpers";

test.describe("sheet tab context menu", () => {
  test("hide/unhide sheets and set tab color", async ({ page }) => {
    await gotoDesktop(page);

    // Lazily create Sheet2 by writing a value into it.
    await page.evaluate(() => {
      const app = window.__formulaApp as any;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });

    const sheet2Tab = page.getByTestId("sheet-tab-Sheet2");
    await expect(sheet2Tab).toBeVisible();

    // Reset dirty state so we can assert metadata operations mark the workbook dirty.
    await page.evaluate(() => {
      const doc = (window.__formulaApp as any).getDocument();
      doc.markSaved();
    });
    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getDocument().isDirty)).toBe(false);

    // Hide Sheet2.
    const menu = page.getByTestId("sheet-tab-context-menu");
    await openSheetTabContextMenu(page, "Sheet2");
    await menu.getByRole("button", { name: "Hide", exact: true }).click();

    await expect(sheet2Tab).toHaveCount(0);
    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getCurrentSheetId())).toBe("Sheet1");
    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getDocument().isDirty)).toBe(true);

    // Unhide Sheet2.
    await page.evaluate(() => {
      const doc = (window.__formulaApp as any).getDocument();
      doc.markSaved();
    });
    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getDocument().isDirty)).toBe(false);

    await openSheetTabContextMenu(page, "Sheet1");
    await menu.getByRole("button", { name: "Unhide…", exact: true }).click();
    await menu.getByRole("button", { name: "Sheet2" }).click();

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getDocument().isDirty)).toBe(true);

    // Set a tab color on Sheet2.
    await page.evaluate(() => {
      const doc = (window.__formulaApp as any).getDocument();
      doc.markSaved();
    });
    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getDocument().isDirty)).toBe(false);

    const sheet2TabVisible = page.getByTestId("sheet-tab-Sheet2");
    await openSheetTabContextMenu(page, "Sheet2");
    await menu.getByRole("button", { name: "Tab Color", exact: true }).click();
    await menu.getByRole("button", { name: "Red" }).click();

    await expect(sheet2TabVisible).toHaveAttribute("data-tab-color", "#ff0000");
    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getDocument().isDirty)).toBe(true);
  });
});
