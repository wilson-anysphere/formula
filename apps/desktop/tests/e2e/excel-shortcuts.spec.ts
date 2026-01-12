import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("excel keyboard shortcuts", () => {
  const GRID_MODES = ["legacy", "shared"] as const;

  for (const mode of GRID_MODES) {
    test(`Ctrl/Cmd+D fills down (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();

        doc.setCellValue(sheetId, "A1", 1);
        doc.setCellValue(sheetId, "B1", 2);

        // Select A1:B3.
        app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 0, endCol: 1 } });
      });
      await waitForIdle(page);

      // Ensure the grid has focus for key events.
      await page.click("#grid", { position: { x: 5, y: 5 } });

      const modifier = process.platform === "darwin" ? "Meta" : "Control";
      await page.keyboard.press(`${modifier}+D`);
      await waitForIdle(page);

      const [a2, b2, a3, b3] = await Promise.all([
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A2")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B2")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A3")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B3")),
      ]);

      expect(a2).toBe("1");
      expect(b2).toBe("2");
      expect(a3).toBe("1");
      expect(b3).toBe("2");
    });

    test(`Ctrl/Cmd+; inserts date (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();

        // Use C1 for this test and ensure it's empty.
        doc.setCellValue(sheetId, "C1", null);
        app.activateCell({ row: 0, col: 2 });
      });
      await waitForIdle(page);

      await page.click("#grid", { position: { x: 5, y: 5 } });

      const modifier = process.platform === "darwin" ? "Meta" : "Control";
      await page.keyboard.press(`${modifier}+Semicolon`);
      await waitForIdle(page);

      const c1 = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
      expect(c1).not.toBe("");
    });
  }
});
