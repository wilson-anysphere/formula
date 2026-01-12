import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("command palette go to", () => {
  test("typing a cell reference navigates immediately", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("B3");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("active-cell")).toHaveText("B3");
  });

  test("typing a range reference selects the range", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("B3:D4");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("selection-range")).toHaveText("B3:D4");
    await expect(page.getByTestId("active-cell")).toHaveText("B3");
  });

  test("typing a sheet-qualified reference switches sheets", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "C3", "Hello from Sheet2 C3");
    });

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Sheet2!C3");
    await page.keyboard.press("Enter");

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2");
    await expect(page.getByTestId("active-cell")).toHaveText("C3");
    await expect(page.getByTestId("active-value")).toHaveText("Hello from Sheet2 C3");
  });

  test("typing a named range navigates immediately", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const workbook = app.getSearchWorkbook?.();
      if (!workbook || typeof workbook.defineName !== "function") throw new Error("Missing search workbook adapter");

      // Name -> single cell B3 on Sheet1.
      workbook.defineName("MyCell", {
        sheetName: "Sheet1",
        range: { startRow: 2, endRow: 2, startCol: 1, endCol: 1 },
      });
      app.getDocument().setCellValue("Sheet1", "B3", "Hello from named range");
    });

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("MyCell");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("active-cell")).toHaveText("B3");
    await expect(page.getByTestId("active-value")).toHaveText("Hello from named range");
  });

  test("sheet-qualified Go To resolves sheet display names to stable ids (no phantom sheets)", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const store = (app as any).getWorkbookSheetStore?.();
      if (!store) throw new Error("Missing workbook sheet store");

      store.addAfter("Sheet1", { id: "sheet-1", name: "Budget" });
      app.getDocument().setCellValue("sheet-1", "A1", "BudgetCell");
    });

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Budget!A1");
    await page.keyboard.press("Enter");

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("sheet-1");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("BudgetCell");

    // Rename the display name and ensure the new name resolves.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const store = (app as any).getWorkbookSheetStore?.();
      if (!store) throw new Error("Missing workbook sheet store");
      store.rename("sheet-1", "Budget2026");
      app.activateSheet("Sheet1");
      app.activateCell({ sheetId: "Sheet1", row: 0, col: 0 });
    });

    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Budget2026!A1");
    await page.keyboard.press("Enter");
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("sheet-1");

    // Stale name should not create a new DocumentController sheet id.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.activateSheet("Sheet1");
      app.activateCell({ sheetId: "Sheet1", row: 0, col: 0 });
    });

    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Budget!A1");
    await expect(page.locator("li.command-palette__item", { hasText: "Go to Budget!A1" })).toHaveCount(0);
    // Close without executing any command result.
    await page.keyboard.press("Escape");

    const state = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      return { activeSheetId: app.getCurrentSheetId(), sheetIds: app.getDocument().getSheetIds() };
    });
    expect(state.activeSheetId).toBe("Sheet1");
    expect(state.sheetIds).not.toContain("Budget");
  });
});
