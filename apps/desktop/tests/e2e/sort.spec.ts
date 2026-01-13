import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("sort", () => {
  test("sort() formula spills values in sorted order", async ({ page }) => {
    await gotoDesktop(page);

    // Seed unsorted values in A1:A3 and insert the SORT formula.
    await page.evaluate(() => {
      const app = window.__formulaApp as any;
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();
      doc.setCellValue(sheetId, "A1", 3);
      doc.setCellValue(sheetId, "A2", 1);
      doc.setCellValue(sheetId, "A3", 2);
      doc.setCellInput(sheetId, "C1", "=SORT(A1:A3)");
    });
    await page.evaluate(() => (window.__formulaApp as any).whenIdle());

    const values = await page.evaluate(async () => {
      const app = window.__formulaApp as any;
      const c1 = await app.getCellValueA1("C1");
      const c2 = await app.getCellValueA1("C2");
      const c3 = await app.getCellValueA1("C3");
      return [c1, c2, c3];
    });

    expect(values).toEqual(["1", "2", "3"]);
  });
});
