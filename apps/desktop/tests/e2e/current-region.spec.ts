import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("Ctrl/Cmd+Shift+* (select current region)", () => {
  const GRID_MODES = ["legacy", "shared"] as const;

  for (const mode of GRID_MODES) {
    test(`selects the current region bounding rectangle (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();

        // The app seeds demo data in A1:D5 for other tests, so place our region
        // in E1:G3 to avoid connecting to that data.
        //
        // Seed a 3x3 region with a "hole" at F2. The current-region algorithm should
        // select the bounding rectangle E1:G3 even when the active cell is empty.
        doc.setCellValue(sheetId, "E1", 1);
        doc.setCellValue(sheetId, "E2", 2);
        doc.setCellValue(sheetId, "E3", 3);
        doc.setCellValue(sheetId, "F1", 4);
        doc.setCellValue(sheetId, "F3", 5);
        doc.setCellValue(sheetId, "G1", 6);
        doc.setCellValue(sheetId, "G2", 7);
        doc.setCellValue(sheetId, "G3", 8);
      });
      await waitForIdle(page);

      const gridBox = await page.locator("#grid").boundingBox();
      expect(gridBox).not.toBeNull();

      const f2Rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("F2"));
      expect(f2Rect).not.toBeNull();

      await page.mouse.click(gridBox!.x + f2Rect!.x + f2Rect!.width / 2, gridBox!.y + f2Rect!.y + f2Rect!.height / 2);
      await expect(page.getByTestId("active-cell")).toHaveText("F2");

      const modifier = process.platform === "darwin" ? "Meta" : "Control";
      await page.keyboard.press(`${modifier}+Shift+8`);

      await expect(page.getByTestId("selection-range")).toHaveText("E1:G3");
      await expect(page.getByTestId("active-cell")).toHaveText("F2");
    });

    test(`falls back to the active cell when there is no adjacent region (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      const gridBox = await page.locator("#grid").boundingBox();
      expect(gridBox).not.toBeNull();

      const j10Rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("J10"));
      expect(j10Rect).not.toBeNull();

      await page.mouse.click(
        gridBox!.x + j10Rect!.x + j10Rect!.width / 2,
        gridBox!.y + j10Rect!.y + j10Rect!.height / 2,
      );
      await expect(page.getByTestId("active-cell")).toHaveText("J10");

      const modifier = process.platform === "darwin" ? "Meta" : "Control";
      await page.keyboard.press(`${modifier}+Shift+8`);

      await expect(page.getByTestId("selection-range")).toHaveText("J10");
      await expect(page.getByTestId("active-cell")).toHaveText("J10");
    });
  }
});
