import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("formula bar - fx function picker", () => {
  const modes = ["legacy", "shared"] as const;

  for (const mode of modes) {
    test(`inserts a selected function template into the formula bar with the cursor inside parentheses (${mode})`, async ({
      page,
    }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      // Pick an empty cell so the formula bar starts blank (the default workbook seeds A1 with text).
      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        if (!app) throw new Error("Missing window.__formulaApp");
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        // E1 is outside the seeded used range (ends at D5).
        doc.setCellValue(sheetId, "E1", null);
        app.activateCell({ sheetId, row: 0, col: 4 }, { focus: false, scrollIntoView: false });
      });
      await expect(page.getByTestId("active-cell")).toHaveText("E1");
      // Ensure any background engine/document syncing triggered by the cell update has settled
      // before we start interacting with the formula bar.
      await page.evaluate(() => (window as any).__formulaApp.whenIdle());

      // 1) Focus the formula bar.
      await page.getByTestId("formula-highlight").click();
      const formulaInput = page.getByTestId("formula-input");
      await expect(formulaInput).toBeVisible();
      await expect(formulaInput).toHaveValue("");

      // 2) Click fx to open the function picker.
      await page.getByTestId("formula-fx-button").click();

      // 3) Search/select a known function.
      const pickerInput = page.getByTestId("formula-function-picker-input");
      await expect(pickerInput).toBeVisible();
      await expect(pickerInput).toBeFocused();
      await pickerInput.fill("sum");
      const sumOption = page.getByTestId("formula-function-picker-item-SUM");
      await expect(sumOption).toBeVisible();
      await sumOption.click();

      // Picker should close.
      await expect(page.getByTestId("formula-function-picker")).toBeHidden();

      // 4) Assert formula bar text becomes `=SUM()` with cursor inside parentheses.
      await expect(formulaInput).toHaveValue("=SUM()");
      await expect(formulaInput).toBeFocused();

      await expect
        .poll(() =>
          page.evaluate(() => {
            const el = document.querySelector<HTMLTextAreaElement>('[data-testid="formula-input"]');
            return { start: el?.selectionStart ?? null, end: el?.selectionEnd ?? null };
          }),
        )
        .toEqual({ start: 5, end: 5 });

      await page.keyboard.type("1");
      await expect(formulaInput).toHaveValue("=SUM(1)");
    });
  }
});
