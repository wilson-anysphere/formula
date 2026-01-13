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

    const menu = page.getByTestId("name-box-menu");
    await expect(menu).toBeVisible();
    await expect(menu).toContainText("E2E_NameBoxRange");

    await menu.getByRole("button", { name: "E2E_NameBoxRange", exact: true }).click();
    await expect(menu).toBeHidden();

    await expect(page.getByTestId("active-cell")).toHaveText("B2");
    await expect(page.getByTestId("selection-range")).toHaveText("B2:C3");
    // Navigating from the dropdown should restore focus to the grid so keyboard shortcuts work.
    await expect.poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id)).toBe("grid");
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

    const menu = page.getByTestId("name-box-menu");
    await expect(menu).toBeVisible();
    await expect(menu).toContainText("E2E_NameBoxTable");

    await menu.getByRole("button", { name: "E2E_NameBoxTable", exact: true }).click();
    await expect(menu).toBeHidden();

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

    const menu = page.getByTestId("name-box-menu");
    await expect(menu).toBeVisible();

    const first = menu.getByRole("button", { name: "E2E_Dropdown_A", exact: true });
    const second = menu.getByRole("button", { name: "E2E_Dropdown_B", exact: true });

    await expect(first).toBeFocused();
    await page.keyboard.press("ArrowDown");
    await expect(second).toBeFocused();

    await page.keyboard.press("Enter");
    await expect(menu).toBeHidden();

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

    const menu = page.getByTestId("name-box-menu");
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "E2E_CrossSheetRange", exact: true }).click();
    await expect(menu).toBeHidden();

    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
    await expect(page.getByTestId("active-cell")).toHaveText("A10");
    await expect(page.getByTestId("selection-range")).toHaveText("A10:B11");
    await expect.poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id)).toBe("grid");
  });

  test("Escape closes the dropdown without changing selection", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const workbook = app.getSearchWorkbook?.();
      if (!workbook) throw new Error("Missing search workbook adapter");

      workbook.clearSchema?.();
      workbook.defineName("E2E_CancelRange", {
        sheetName: app.getCurrentSheetId(),
        range: { startRow: 9, startCol: 0, endRow: 10, endCol: 1 }, // A10:B11
      });
    });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("selection-range")).toHaveText("A1");

    await page.getByTestId("name-box-dropdown").click();

    const menu = page.getByTestId("name-box-menu");
    await expect(menu).toBeVisible();

    await page.keyboard.press("Escape");
    await expect(menu).toBeHidden();

    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("selection-range")).toHaveText("A1");
  });
});
