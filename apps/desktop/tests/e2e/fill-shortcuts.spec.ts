import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("fill shortcuts (Ctrl/Cmd+D, Ctrl/Cmd+R)", () => {
  const GRID_MODES = ["legacy", "shared"] as const;

  for (const mode of GRID_MODES) {
    test(`Ctrl+D fills down from top row (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        doc.setCellValue(sheetId, "A1", 1);
        app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 0, endCol: 0 } });
        app.focus();
      });
      await waitForIdle(page);

      await page.keyboard.press("ControlOrMeta+D");
      await waitForIdle(page);

      const [a2, a3] = await Promise.all([
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A2")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A3")),
      ]);
      expect(a2).toBe("1");
      expect(a3).toBe("1");
    });

    test(`Ctrl+R fills right from leftmost column (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        doc.setCellValue(sheetId, "A1", 1);
        doc.setCellValue(sheetId, "B1", 2);
        app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 2 } });
        app.focus();
      });
      await waitForIdle(page);

      await page.keyboard.press("ControlOrMeta+R");
      await waitForIdle(page);

      const [b1, c1] = await Promise.all([
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1")),
        page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1")),
      ]);
      expect(b1).toBe("1");
      expect(c1).toBe("1");
    });

    test(`Ctrl+D shifts formulas (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        doc.setCellValue(sheetId, "A1", 1);
        doc.setCellValue(sheetId, "B1", 2);
        doc.setCellInput(sheetId, "C1", "=A1+B1");
        app.selectRange({ range: { startRow: 0, endRow: 2, startCol: 2, endCol: 2 } });
        app.focus();
      });
      await waitForIdle(page);

      await page.keyboard.press("ControlOrMeta+D");
      await waitForIdle(page);

      const [c2Formula, c3Formula] = await Promise.all([
        page.evaluate(
          () =>
            (window as any).__formulaApp.getDocument().getCell((window as any).__formulaApp.getCurrentSheetId(), "C2").formula
        ),
        page.evaluate(
          () =>
            (window as any).__formulaApp.getDocument().getCell((window as any).__formulaApp.getCurrentSheetId(), "C3").formula
        ),
      ]);

      expect(c2Formula).toBe("=A2+B2");
      expect(c3Formula).toBe("=A3+B3");
    });
  }
});
