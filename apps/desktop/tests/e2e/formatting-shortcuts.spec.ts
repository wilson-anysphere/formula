import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

async function waitForIdle(page: Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

async function getA1FontProp(page: Page, prop: "bold" | "italic"): Promise<boolean> {
  return await page.evaluate((propName) => {
    const app = (window as any).__formulaApp;
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();
    const format = doc.getCellFormat(sheetId, "A1");
    return format?.font?.[propName] === true;
  }, prop);
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

test.describe("formatting shortcuts", () => {
  test("Ctrl/Cmd+B toggles bold on the selection", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.setCellValue(sheetId, "A1", "Hello");
      app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
      app.focus();
    });
    await waitForIdle(page);

    // Ensure the grid is focused so keyboard events route through SpreadsheetApp's keydown handler.
    await page.locator("#grid").focus();

    const before = await getA1FontProp(page, "bold");
    await page.keyboard.press("ControlOrMeta+b");
    await waitForIdle(page);

    expect(await getA1FontProp(page, "bold")).toBe(!before);

    await page.keyboard.press("ControlOrMeta+b");
    await waitForIdle(page);
    expect(await getA1FontProp(page, "bold")).toBe(before);
  });

  test("Ctrl+I toggles italic; Cmd+I (macOS) / Ctrl+Shift+A (Win/Linux) opens AI panel without changing formatting", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
      app.focus();
    });
    await waitForIdle(page);

    await page.locator("#grid").focus();

    const initialItalic = await getA1FontProp(page, "italic");

    await page.keyboard.press("Control+I");
    await waitForIdle(page);
    expect(await getA1FontProp(page, "italic")).toBe(!initialItalic);

    await page.keyboard.press("Control+I");
    await waitForIdle(page);
    const italicAfterToggles = await getA1FontProp(page, "italic");
    expect(italicAfterToggles).toBe(initialItalic);

    await expect(page.getByTestId("panel-aiChat")).toHaveCount(0);

    const aiShortcut = process.platform === "darwin" ? "Meta+I" : "Control+Shift+A";
    await page.keyboard.press(aiShortcut);
    await expect(page.getByTestId("panel-aiChat")).toBeVisible();

    expect(await getA1FontProp(page, "italic")).toBe(italicAfterToggles);
  });

  test("Ctrl/Cmd+U toggles underline on the selection", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.setCellValue(sheetId, "A1", "Hello");
      app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
      app.focus();
    });
    await waitForIdle(page);

    // Ensure the grid is focused so keyboard events route through SpreadsheetApp's keydown handler.
    await page.locator("#grid").focus();

    expect(await getA1Underline(page)).toBe(false);

    await page.keyboard.press("ControlOrMeta+U");
    await waitForIdle(page);

    expect(await getA1Underline(page)).toBe(true);

    await page.keyboard.press("ControlOrMeta+U");
    await waitForIdle(page);
    expect(await getA1Underline(page)).toBe(false);
  });

  for (const [name, cfg] of [
    ["Ctrl/Cmd+Shift+$ applies the currency number format preset", { key: "Digit4", expected: "$#,##0.00" }],
    ["Ctrl/Cmd+Shift+% applies the percent number format preset", { key: "Digit5", expected: "0%" }],
    ["Ctrl/Cmd+Shift+# applies the date number format preset", { key: "Digit3", expected: "m/d/yyyy" }],
  ] as const) {
    test(name, async ({ page }) => {
      await gotoDesktop(page);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();

        doc.setCellValue(sheetId, "A1", 123.45);
        app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
        app.focus();
      });
      await waitForIdle(page);

      // Use the physical digit key so the shortcut is stable regardless of keyboard layout (e.g. Shift+4 -> "$").
      await page.locator("#grid").focus();
      await page.keyboard.press(`ControlOrMeta+Shift+${cfg.key}`);
      await waitForIdle(page);

      expect(await getA1NumberFormat(page)).toBe(cfg.expected);
    });
  }
});
