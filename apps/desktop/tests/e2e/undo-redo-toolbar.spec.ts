import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("titlebar undo/redo buttons", () => {
  const GRID_MODES = ["legacy", "shared"] as const;

  for (const mode of GRID_MODES) {
    test(`undo/redo works via titlebar (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      const undo = page.getByTestId("undo");
      const redo = page.getByTestId("redo");

      await expect(undo).toBeVisible();
      await expect(redo).toBeVisible();

      // Ensure a clean undo stack regardless of any boot-time document initialization.
      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const doc = app.getDocument();
        doc.applyState(doc.encodeState());
        doc.markSaved();
      });
      await waitForIdle(page);
      const initialValue = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        return app.getDocument().getCell(sheetId, "A1").value;
      });

      await expect(undo).toBeDisabled();
      await expect(redo).toBeDisabled();

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        doc.setCellValue(sheetId, "A1", 42, { label: "Set Cell" });
      });
      await waitForIdle(page);

      await expect(undo).toBeEnabled();
      await expect(redo).toBeDisabled();
      await expect(undo).toHaveAttribute("aria-label", "Undo Set Cell");

      // While editing, undo/redo should be disabled (Excel-like behavior: Ctrl+Z becomes text undo).
      // In view mode the formula bar textarea is hidden; click the highlight to enter edit mode.
      await page.getByTestId("formula-highlight").click();
      const formulaInput = page.getByTestId("formula-input");
      await expect(formulaInput).toBeVisible();
      await expect(formulaInput).toBeFocused();
      await expect(undo).toBeDisabled();
      await expect(redo).toBeDisabled();
      await page.keyboard.press("Escape");
      await expect(undo).toBeEnabled();

      await undo.click();
      await waitForIdle(page);

      await expect(undo).toBeDisabled();
      await expect(redo).toBeEnabled();
      await expect(redo).toHaveAttribute("aria-label", "Redo Set Cell");

      const afterUndo = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        return app.getDocument().getCell(sheetId, "A1").value;
      });
      expect(afterUndo).toEqual(initialValue);

      await redo.click();
      await waitForIdle(page);

      const afterRedo = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        return app.getDocument().getCell(sheetId, "A1").value;
      });
      expect(afterRedo).toBe(42);
    });
  }
});
