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
});

