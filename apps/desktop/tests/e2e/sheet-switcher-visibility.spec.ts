import { expect, test } from "@playwright/test";

import { expectSheetPosition, gotoDesktop, openSheetTabContextMenu } from "./helpers";

test.describe("sheet switcher", () => {
  test("only lists visible sheets (hide/unhide)", async ({ page }) => {
    await gotoDesktop(page);

    // Create Sheet2 + Sheet3.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = window.__formulaApp as any;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();

      doc.setCellValue("Sheet2", "A1", "Hello from Sheet2");
      doc.setCellValue("Sheet3", "A1", "Hello from Sheet3");
    });

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();
    await expectSheetPosition(page, { position: 1, total: 3 });

    const switcher = page.getByTestId("sheet-switcher");

    await expect(switcher.locator("option")).toHaveCount(3);
    await expect(switcher.locator("option", { hasText: "Sheet2" })).toHaveCount(1);

    // Make Sheet2 active so hiding it exercises the "active sheet becomes hidden" path.
    await page.getByTestId("sheet-tab-Sheet2").click();
    await expectSheetPosition(page, { position: 2, total: 3 });

    // Hide Sheet2 via context menu.
    let menu = await openSheetTabContextMenu(page, "Sheet2");
    await expect(page.getByTestId("context-menu")).toBeHidden();
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Hide", exact: true }).click();

    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveCount(0);
    await expect(switcher.locator("option")).toHaveCount(2);
    await expect(switcher.locator("option", { hasText: "Sheet2" })).toHaveCount(0);
    {
      const optionLabels = await switcher.locator("option").allTextContents();
      expect(optionLabels).toEqual(["Sheet1", "Sheet3"]);
    }
    {
      // Hiding the active sheet should activate a visible sheet.
      const activeSheetId = await page.evaluate(() => (window.__formulaApp as any).getCurrentSheetId());
      expect(activeSheetId).not.toEqual("Sheet2");
      expect(["Sheet1", "Sheet3"]).toContain(activeSheetId);
      await expect(switcher).toHaveValue(activeSheetId);

      // The sheet position indicator should update to use visible sheets only.
      if (activeSheetId === "Sheet1") {
        await expectSheetPosition(page, { position: 1, total: 2 });
      } else if (activeSheetId === "Sheet3") {
        await expectSheetPosition(page, { position: 2, total: 2 });
      } else {
        throw new Error(`Unexpected active sheet after hide: ${activeSheetId}`);
      }
    }

    // Unhide Sheet2 via context menu on any visible tab.
    menu = await openSheetTabContextMenu(page, "Sheet1");
    await expect(page.getByTestId("context-menu")).toBeHidden();
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Unhide…", exact: true }).click();
    await menu.getByRole("button", { name: "Sheet2" }).click();

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(switcher.locator("option")).toHaveCount(3);
    {
      const optionLabels = await switcher.locator("option").allTextContents();
      expect(optionLabels).toEqual(["Sheet1", "Sheet2", "Sheet3"]);
    }

    const activeAfterUnhide = await page.evaluate(() => (window.__formulaApp as any).getCurrentSheetId());
    if (activeAfterUnhide === "Sheet1") {
      await expectSheetPosition(page, { position: 1, total: 3 });
    } else if (activeAfterUnhide === "Sheet2") {
      await expectSheetPosition(page, { position: 2, total: 3 });
    } else if (activeAfterUnhide === "Sheet3") {
      await expectSheetPosition(page, { position: 3, total: 3 });
    } else {
      throw new Error(`Unexpected active sheet after unhide: ${activeAfterUnhide}`);
    }
  });
}); 
