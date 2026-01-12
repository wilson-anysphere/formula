import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("formula bar (keyboard undo)", () => {
  test("Ctrl/Cmd+Z undoes text in the formula bar when focused", async ({ page }) => {
    await gotoDesktop(page);

    // Seed history so a spreadsheet-level undo would be observable if the keybinding
    // were routed incorrectly while the formula bar has focus.
    await page.evaluate(async () => {
      const app = window.__formulaApp as any;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, "A1", "FormulaBarUndoText", { label: "Set A1" });
      app.activateCell({ sheetId, row: 0, col: 0 });
      app.refresh();
      await app.whenIdle();
    });

    // Focus the formula bar and type a single character so the undo stack is stable.
    await page.getByTestId("formula-highlight").click();
    const input = page.getByTestId("formula-input");
    await expect(input).toBeFocused();

    const initial = await input.inputValue();
    await page.keyboard.type("x");
    await expect(input).toHaveValue(`${initial}x`);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Z`);
    await expect(input).toHaveValue(initial);

    // Ensure the spreadsheet edit was not undone.
    const cellValue = await page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A1"));
    expect(cellValue).toBe("FormulaBarUndoText");
  });
});

