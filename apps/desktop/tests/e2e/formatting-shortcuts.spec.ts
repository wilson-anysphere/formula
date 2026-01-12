import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("formatting shortcuts", () => {
  test("Ctrl/Cmd+B toggles bold on the selection", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      doc.setCellValue("Sheet1", "A1", "Hello");
    });

    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("A1");
      return rect && typeof rect.x === "number" && rect.width > 0 && rect.height > 0;
    });

    const rect = (await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"))) as {
      x: number;
      y: number;
      width: number;
      height: number;
    };

    // Select A1.
    await page.locator("#grid").click({
      position: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
    });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+B`);

    await expect
      .poll(async () => {
        return await page.evaluate(() => {
          const app = (window as any).__formulaApp;
          const doc = app.getDocument();
          const cell = doc.getCell("Sheet1", "A1");
          const style = doc.styleTable.get(cell.styleId);
          return style.font?.bold === true;
        });
      })
      .toBe(true);
  });
});

