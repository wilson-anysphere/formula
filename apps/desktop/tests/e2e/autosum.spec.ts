import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("AutoSum (Alt+=)", () => {
  const modes = ["legacy", "shared"] as const;

  for (const mode of modes) {
    test(`inserts SUM formula below a vertical range (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        doc.setCellValue(sheetId, "A1", 1);
        doc.setCellValue(sheetId, "A2", 2);
        doc.setCellValue(sheetId, "A3", 3);
        app.selectRange({
          range: {
            startRow: 0,
            endRow: 2,
            startCol: 0,
            endCol: 0,
          },
        });
      });
      await waitForIdle(page);

      await page.keyboard.press("Alt+Equal");
      await waitForIdle(page);

      await expect(page.getByTestId("active-cell")).toHaveText("A4");

      const result = await page.evaluate(async () => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        return {
          formula: doc.getCell(sheetId, "A4").formula,
          value: await app.getCellValueA1("A4"),
          undoLabel: doc.undoLabel,
        };
      });

      expect(result.formula).toBe("=SUM(A1:A3)");
      expect(result.value).toBe("6");
      expect(result.undoLabel).toBe("AutoSum");
    });
  }
});

