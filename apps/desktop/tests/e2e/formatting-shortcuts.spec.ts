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
    return doc.getCellFormat(sheetId, "A1").font?.[propName] ?? false;
  }, prop);
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

    await page.keyboard.press("ControlOrMeta+B");
    await waitForIdle(page);

    expect(await getA1FontProp(page, "bold")).toBe(true);
  });

  test("Ctrl+I toggles italic; Cmd+I opens AI panel without changing formatting", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
      app.focus();
    });
    await waitForIdle(page);

    expect(await getA1FontProp(page, "italic")).toBe(false);

    await page.keyboard.press("Control+I");
    await waitForIdle(page);
    expect(await getA1FontProp(page, "italic")).toBe(true);

    await page.keyboard.press("Control+I");
    await waitForIdle(page);
    const italicOff = await getA1FontProp(page, "italic");
    expect(italicOff).toBe(false);

    await expect(page.getByTestId("panel-aiChat")).toHaveCount(0);

    await page.keyboard.press("Meta+I");
    await expect(page.getByTestId("panel-aiChat")).toBeVisible();

    expect(await getA1FontProp(page, "italic")).toBe(italicOff);
  });
});

