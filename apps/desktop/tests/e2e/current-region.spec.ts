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

        // Seed a 3x3 region with a "hole" at B2. The current-region algorithm should
        // select the bounding rectangle A1:C3 even when the active cell is empty.
        doc.setCellValue(sheetId, "A1", 1);
        doc.setCellValue(sheetId, "A2", 2);
        doc.setCellValue(sheetId, "A3", 3);
        doc.setCellValue(sheetId, "B1", 4);
        doc.setCellValue(sheetId, "B3", 5);
        doc.setCellValue(sheetId, "C1", 6);
        doc.setCellValue(sheetId, "C2", 7);
        doc.setCellValue(sheetId, "C3", 8);
      });
      await waitForIdle(page);

      const gridBox = await page.locator("#grid").boundingBox();
      expect(gridBox).not.toBeNull();

      const b2Rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("B2"));
      expect(b2Rect).not.toBeNull();

      await page.mouse.click(gridBox!.x + b2Rect!.x + b2Rect!.width / 2, gridBox!.y + b2Rect!.y + b2Rect!.height / 2);
      await expect(page.getByTestId("active-cell")).toHaveText("B2");

      const modifier = process.platform === "darwin" ? "Meta" : "Control";
      await page.keyboard.press(`${modifier}+Shift+8`);

      await expect(page.getByTestId("selection-range")).toHaveText("A1:C3");
      await expect(page.getByTestId("active-cell")).toHaveText("B2");
    });
  }
});

