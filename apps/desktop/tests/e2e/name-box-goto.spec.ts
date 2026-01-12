import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("name box go to", () => {
  test("sheet-qualified references resolve sheet display names to stable ids (no phantom sheets)", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const store = (app as any).getWorkbookSheetStore?.();
      if (!store) throw new Error("Missing workbook sheet store");

      // Create a stable-id sheet ("Sheet2") then rename its display name to "Budget".
      app.getDocument().setCellValue("Sheet2", "A1", "BudgetCell");
      store.rename("Sheet2", "Budget");
    });

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    const nameBox = page.getByTestId("formula-address");
    await expect(nameBox).toBeVisible();
    await nameBox.fill("Budget!A1");
    await nameBox.press("Enter");

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("BudgetCell");

    const sheetIds = await page.evaluate(() => (window as any).__formulaApp.getDocument().getSheetIds());
    expect(sheetIds).not.toContain("Budget");
  });
});

