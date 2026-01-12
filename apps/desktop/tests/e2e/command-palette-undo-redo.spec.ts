import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("command palette (Undo/Redo)", () => {
  test("shows keybinding hints and runs edit.undo/edit.redo", async ({ page }) => {
    await gotoDesktop(page);

    const before = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));

    await page.evaluate(async () => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, "A1", "PaletteUndoRedo", { label: "Set A1" });
      app.refresh();
      await app.whenIdle();
    });

    const edited = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(edited).toBe("PaletteUndoRedo");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Undo.
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Undo");

    const expectedUndoShortcut = process.platform === "darwin" ? "⌘Z" : "Ctrl+Z";
    const undoItem = page
      .getByTestId("command-palette-list")
      .locator(".command-palette__item")
      .filter({ hasText: "Undo" })
      .first();
    await expect(undoItem).toBeVisible();
    await expect(undoItem.locator(".command-palette__shortcut")).toHaveText(expectedUndoShortcut);
    await undoItem.click();
    await expect(page.getByTestId("command-palette")).toBeHidden();

    const afterUndo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(afterUndo).toBe(before);

    // Redo.
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Redo");

    const expectedRedoShortcut = process.platform === "darwin" ? "⇧⌘Z" : "Ctrl+Y";
    const redoItem = page
      .getByTestId("command-palette-list")
      .locator(".command-palette__item")
      .filter({ hasText: "Redo" })
      .first();
    await expect(redoItem).toBeVisible();
    await expect(redoItem.locator(".command-palette__shortcut")).toHaveText(expectedRedoShortcut);
    await redoItem.click();
    await expect(page.getByTestId("command-palette")).toBeHidden();

    const afterRedo = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(afterRedo).toBe("PaletteUndoRedo");
  });
});

