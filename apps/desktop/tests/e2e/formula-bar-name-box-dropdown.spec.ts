import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("formula bar name box dropdown", () => {
  test("selecting a named range navigates to its range", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const workbook = app.getSearchWorkbook?.();
      if (!workbook) throw new Error("Missing search workbook adapter");

      // Define a 2x2 named range at B2:C3 (0-based coords).
      workbook.defineName("E2E_NameBoxRange", {
        sheetName: app.getCurrentSheetId(),
        range: { startRow: 1, startCol: 1, endRow: 2, endCol: 2 },
      });
    });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.getByTestId("name-box-dropdown").click();

    const quickPick = page.getByTestId("quick-pick");
    await expect(quickPick).toBeVisible();
    await expect(quickPick).toContainText("E2E_NameBoxRange");

    await quickPick.getByRole("button", { name: /E2E_NameBoxRange/ }).click();
    await expect(quickPick).toBeHidden();

    await expect(page.getByTestId("active-cell")).toHaveText("B2");
    await expect(page.getByTestId("selection-range")).toHaveText("B2:C3");
  });

  test("selecting a table navigates to its full range", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const workbook = app.getSearchWorkbook?.();
      if (!workbook) throw new Error("Missing search workbook adapter");

      // Seed a named range too so the dropdown renders multiple items (table + name),
      // ensuring the UI can list heterogeneous entries.
      workbook.defineName("E2E_NameBoxRange", {
        sheetName: app.getCurrentSheetId(),
        range: { startRow: 1, startCol: 1, endRow: 2, endCol: 2 },
      });

      // Define a 2x2 table at A5:B6 (0-based coords).
      workbook.addTable({
        name: "E2E_NameBoxTable",
        sheetName: app.getCurrentSheetId(),
        startRow: 4,
        startCol: 0,
        endRow: 5,
        endCol: 1,
        columns: ["Col1", "Col2"],
      });
    });

    await page.getByTestId("name-box-dropdown").click();

    const quickPick = page.getByTestId("quick-pick");
    await expect(quickPick).toBeVisible();
    await expect(quickPick).toContainText("E2E_NameBoxTable");

    await quickPick.getByRole("button", { name: /E2E_NameBoxTable/ }).click();
    await expect(quickPick).toBeHidden();

    await expect(page.getByTestId("active-cell")).toHaveText("A5");
    await expect(page.getByTestId("selection-range")).toHaveText("A5:B6");
  });

  test("keyboard navigation (ArrowDown + Enter) selects items in the dropdown", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const workbook = app.getSearchWorkbook?.();
      if (!workbook) throw new Error("Missing search workbook adapter");

      // Ensure the menu is deterministic for keyboard navigation.
      workbook.clearSchema?.();

      // Define two named ranges so we can arrow between them. The menu sorts
      // alphabetically, so A should be focused first.
      workbook.defineName("E2E_Dropdown_A", {
        sheetName: app.getCurrentSheetId(),
        range: { startRow: 1, startCol: 1, endRow: 2, endCol: 2 }, // B2:C3
      });
      workbook.defineName("E2E_Dropdown_B", {
        sheetName: app.getCurrentSheetId(),
        range: { startRow: 4, startCol: 3, endRow: 5, endCol: 4 }, // D5:E6
      });
    });

    await page.getByTestId("name-box-dropdown").click();

    const quickPick = page.getByTestId("quick-pick");
    await expect(quickPick).toBeVisible();

    const first = quickPick.getByRole("button", { name: /E2E_Dropdown_A/ });
    const second = quickPick.getByRole("button", { name: /E2E_Dropdown_B/ });

    await expect(first).toBeFocused();
    await page.keyboard.press("ArrowDown");
    await expect(second).toBeFocused();

    await page.keyboard.press("Enter");
    await expect(quickPick).toBeHidden();

    await expect(page.getByTestId("active-cell")).toHaveText("D5");
    await expect(page.getByTestId("selection-range")).toHaveText("D5:E6");
  });

  test("selecting a named range on another sheet switches sheets and selects the range", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const workbook = app.getSearchWorkbook?.();
      if (!workbook) throw new Error("Missing search workbook adapter");

      workbook.clearSchema?.();

      // Materialize Sheet2 so navigation can activate it.
      app.getDocument().setCellValue("Sheet2", "A10", "CrossSheet");

      workbook.defineName("E2E_CrossSheetRange", {
        sheetName: "Sheet2",
        range: { startRow: 9, startCol: 0, endRow: 10, endCol: 1 }, // A10:B11
      });
    });

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    await page.getByTestId("name-box-dropdown").click();

    const quickPick = page.getByTestId("quick-pick");
    await expect(quickPick).toBeVisible();
    await quickPick.getByRole("button", { name: /E2E_CrossSheetRange/ }).click();
    await expect(quickPick).toBeHidden();

    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
    await expect(page.getByTestId("active-cell")).toHaveText("A10");
    await expect(page.getByTestId("selection-range")).toHaveText("A10:B11");
  });
});
