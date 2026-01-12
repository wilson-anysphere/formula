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
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        // Ensure a known baseline so shortcut assertions can be deterministic.
        doc.setRangeFormat(sheetId, "A1", { font: { bold: false, underline: false }, numberFormat: null }, { label: "Reset" });
        app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
        app.focus();
      });
      await waitForIdle(page);
      // Ensure we're not in an edit context (formula bar/cell editor), since the shortcut
      // handler intentionally ignores key events coming from editable targets.
      await page.keyboard.press("Escape");
      await waitForIdle(page);
      await page.focus("#grid");
      await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.isEditing())).toBe(false);
      await expect.poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id)).toBe("grid");

      // ---- Bold --------------------------------------------------------------
      await expect.poll(async () => {
        await waitForIdle(page);
        return page.evaluate(() => {
          const app = (window as any).__formulaApp;
          const sheetId = app.getCurrentSheetId();
          const doc = app.getDocument();
          return doc.getCellFormat(sheetId, "A1").font?.bold;
        });
      }).toBe(false);
      await page.keyboard.press("ControlOrMeta+B");
      await expect.poll(async () => {
        await waitForIdle(page);
        return page.evaluate(() => {
          const app = (window as any).__formulaApp;
          const sheetId = app.getCurrentSheetId();
          const doc = app.getDocument();
          return doc.getCellFormat(sheetId, "A1").font?.bold;
        });
      }).toBe(true);

      await page.keyboard.press("ControlOrMeta+B");
      await expect.poll(async () => {
        await waitForIdle(page);
        return page.evaluate(() => {
          const app = (window as any).__formulaApp;
          const sheetId = app.getCurrentSheetId();
          const doc = app.getDocument();
          return doc.getCellFormat(sheetId, "A1").font?.bold;
        });
      }).toBe(false);

      // ---- Underline ---------------------------------------------------------
      await expect.poll(async () => {
        await waitForIdle(page);
        return page.evaluate(() => {
          const app = (window as any).__formulaApp;
          const sheetId = app.getCurrentSheetId();
          const doc = app.getDocument();
          return doc.getCellFormat(sheetId, "A1").font?.underline;
        });
      }).toBe(false);
      await page.keyboard.press("ControlOrMeta+U");
      await expect.poll(async () => {
        await waitForIdle(page);
        return page.evaluate(() => {
          const app = (window as any).__formulaApp;
          const sheetId = app.getCurrentSheetId();
          const doc = app.getDocument();
          return doc.getCellFormat(sheetId, "A1").font?.underline;
        });
      }).toBe(true);

      await page.keyboard.press("ControlOrMeta+U");
      await expect.poll(async () => {
        await waitForIdle(page);
        return page.evaluate(() => {
          const app = (window as any).__formulaApp;
          const sheetId = app.getCurrentSheetId();
          const doc = app.getDocument();
          return doc.getCellFormat(sheetId, "A1").font?.underline;
        });
      }).toBe(false);

      // ---- Number format presets --------------------------------------------
      await page.keyboard.press("ControlOrMeta+Shift+Digit4");
      await expect.poll(async () => {
        await waitForIdle(page);
        return page.evaluate(() => {
          const app = (window as any).__formulaApp;
          const sheetId = app.getCurrentSheetId();
          const doc = app.getDocument();
          return doc.getCellFormat(sheetId, "A1").numberFormat;
        });
      }).toBe("$#,##0.00");

      await page.keyboard.press("ControlOrMeta+Shift+Digit5");
      await expect.poll(async () => {
        await waitForIdle(page);
        return page.evaluate(() => {
          const app = (window as any).__formulaApp;
          const sheetId = app.getCurrentSheetId();
          const doc = app.getDocument();
          return doc.getCellFormat(sheetId, "A1").numberFormat;
        });
      }).toBe("0%");

      await page.keyboard.press("ControlOrMeta+Shift+Digit3");
      await expect.poll(async () => {
        await waitForIdle(page);
        return page.evaluate(() => {
          const app = (window as any).__formulaApp;
          const sheetId = app.getCurrentSheetId();
          const doc = app.getDocument();
          return doc.getCellFormat(sheetId, "A1").numberFormat;
        });
      }).toBe("m/d/yyyy");
    });
  }
});
