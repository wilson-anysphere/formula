import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("grid context menu (Undo/Redo)", () => {
  test("Undo is enabled after an edit and reverts A1", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    const before = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, "A1", "ContextMenuEdit", { label: "Set A1" });
      app.refresh();
    });
    await waitForIdle(page);

    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("A1");
      return rect && rect.width > 0 && rect.height > 0;
    });
    const a1 = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));
    await page.locator("#grid").click({ button: "right", position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const undo = menu.getByRole("button", { name: "Undo Set A1" });
    await expect(undo).toBeEnabled();
    const expectedUndoShortcut = process.platform === "darwin" ? "⌘Z" : "Ctrl+Z";
    await expect(undo.locator('span[aria-hidden="true"]')).toHaveText(expectedUndoShortcut);

    const redo = menu.getByRole("button", { name: "Redo" });
    await expect(redo).toBeDisabled();
    const expectedRedoShortcut = process.platform === "darwin" ? "⇧⌘Z" : "Ctrl+Y";
    await expect(redo.locator('span[aria-hidden="true"]')).toHaveText(expectedRedoShortcut);

    await undo.click();
    await waitForIdle(page);

    const after = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(after).toBe(before);
  });
});
