import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

async function getA1Underline(page: Page): Promise<boolean> {
  return await page.evaluate(() => {
    const app = (window as any).__formulaApp;
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();
    return doc.getCellFormat(sheetId, "A1").font?.underline ?? false;
  });
}

async function getA1NumberFormat(page: Page): Promise<string | null> {
  return await page.evaluate(() => {
    const app = (window as any).__formulaApp;
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();
    return doc.getCellFormat(sheetId, "A1").numberFormat ?? null;
  });
}

test.describe("formatting shortcuts (more)", () => {
  const GRID_MODES = ["shared", "legacy"] as const;

  for (const mode of GRID_MODES) {
    test(`Ctrl/Cmd+1 opens the Format Cells dialog (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Ensure the grid is focused (keyboard shortcuts are routed globally, but focus helps mimic real usage).
      await page.locator("#grid").focus();

      // Dispatch directly to avoid browser tab-switch shortcuts interfering with Ctrl/Cmd+1.
      await page.evaluate((isMac) => {
        const target = (document.activeElement as HTMLElement | null) ?? window;
        target.dispatchEvent(
          new KeyboardEvent("keydown", {
            key: "1",
            code: "Digit1",
            metaKey: isMac,
            ctrlKey: !isMac,
            bubbles: true,
            cancelable: true,
          }),
        );
      }, process.platform === "darwin");

      const dialog = page.getByTestId("format-cells-dialog");
      await expect(dialog).toBeVisible();

      // Close and ensure focus returns to the grid (Excel-like behavior).
      await page.keyboard.press("Escape");
      await expect(dialog).toBeHidden();
      await expect.poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id)).toBe("grid");
    });

    test(`Ctrl/Cmd+U toggles underline on the selection (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();

        doc.setCellValue(sheetId, "A1", "Hello");
        app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
        app.focus();
      });
      await waitForIdle(page);

      expect(await getA1Underline(page)).toBe(false);

      await page.keyboard.press("ControlOrMeta+U");
      await waitForIdle(page);
      expect(await getA1Underline(page)).toBe(true);

      await page.keyboard.press("ControlOrMeta+U");
      await waitForIdle(page);
      expect(await getA1Underline(page)).toBe(false);
    });

    for (const testCase of [
      { name: "currency", shortcut: "ControlOrMeta+Shift+Digit4", expected: "$#,##0.00" },
      { name: "percent", shortcut: "ControlOrMeta+Shift+Digit5", expected: "0%" },
      { name: "date", shortcut: "ControlOrMeta+Shift+Digit3", expected: "m/d/yyyy" },
    ] as const) {
      test(`Ctrl/Cmd+Shift+${testCase.name} applies number format preset (${mode})`, async ({ page }) => {
        await gotoDesktop(page, `/?grid=${mode}`);

        await page.evaluate(() => {
          const app = (window as any).__formulaApp;
          const sheetId = app.getCurrentSheetId();
          const doc = app.getDocument();

          doc.setCellValue(sheetId, "A1", 1234.56);
          app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
          app.focus();
        });
        await waitForIdle(page);

        await page.keyboard.press(testCase.shortcut);
        await waitForIdle(page);

        expect(await getA1NumberFormat(page)).toBe(testCase.expected);
      });
    }
  }
});
