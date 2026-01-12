import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("formatting shortcuts (Excel presets)", () => {
  const GRID_MODES = ["legacy", "shared"] as const;

  for (const mode of GRID_MODES) {
    test(`Ctrl/Cmd+B/U toggles and Ctrl/Cmd+Shift+$/%/# apply number formats (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
        app.focus();
      });
      await waitForIdle(page);

      // ---- Bold --------------------------------------------------------------
      const initialBold = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        return doc.getCellFormat(sheetId, "A1").font?.bold === true;
      });
      await page.keyboard.press("ControlOrMeta+B");
      await waitForIdle(page);
      let bold = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        return doc.getCellFormat(sheetId, "A1").font?.bold === true;
      });
      expect(bold).toBe(!initialBold);

      await page.keyboard.press("ControlOrMeta+B");
      await waitForIdle(page);
      bold = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        return doc.getCellFormat(sheetId, "A1").font?.bold === true;
      });
      expect(bold).toBe(initialBold);

      // ---- Underline ---------------------------------------------------------
      const initialUnderline = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        return doc.getCellFormat(sheetId, "A1").font?.underline === true;
      });
      await page.keyboard.press("ControlOrMeta+U");
      await waitForIdle(page);
      let underline = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        return doc.getCellFormat(sheetId, "A1").font?.underline === true;
      });
      expect(underline).toBe(!initialUnderline);

      await page.keyboard.press("ControlOrMeta+U");
      await waitForIdle(page);
      underline = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        return doc.getCellFormat(sheetId, "A1").font?.underline === true;
      });
      expect(underline).toBe(initialUnderline);

      // ---- Number format presets --------------------------------------------
      await page.keyboard.press("ControlOrMeta+Shift+Digit4");
      await waitForIdle(page);
      let numberFormat = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        return doc.getCellFormat(sheetId, "A1").numberFormat;
      });
      expect(numberFormat).toBe("$#,##0.00");

      await page.keyboard.press("ControlOrMeta+Shift+Digit5");
      await waitForIdle(page);
      numberFormat = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        return doc.getCellFormat(sheetId, "A1").numberFormat;
      });
      expect(numberFormat).toBe("0%");

      await page.keyboard.press("ControlOrMeta+Shift+Digit3");
      await waitForIdle(page);
      numberFormat = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        return doc.getCellFormat(sheetId, "A1").numberFormat;
      });
      expect(numberFormat).toBe("m/d/yyyy");
    });
  }
});
