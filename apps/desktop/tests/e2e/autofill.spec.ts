import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

async function dragFromTo(page: import("@playwright/test").Page, from: { x: number; y: number }, to: { x: number; y: number }) {
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y);
  await page.mouse.up();
}

test.describe("grid autofill (fill handle)", () => {
  test("fills series + shifts formulas", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellValue(sheetId, "A2", 2);
      app.selectRange({ range: { startRow: 0, endRow: 1, startCol: 0, endCol: 0 } });
    });
    await waitForIdle(page);

    const gridBox = await page.locator("#grid").boundingBox();
    expect(gridBox).not.toBeNull();

    const handle = await page.evaluate(() => (window as any).__formulaApp.getFillHandleRect());
    expect(handle).not.toBeNull();

    const a4Rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A4"));
    expect(a4Rect).not.toBeNull();

    await dragFromTo(
      page,
      { x: gridBox!.x + handle!.x + handle!.width / 2, y: gridBox!.y + handle!.y + handle!.height / 2 },
      { x: gridBox!.x + a4Rect!.x + a4Rect!.width / 2, y: gridBox!.y + a4Rect!.y + a4Rect!.height / 2 }
    );
    await waitForIdle(page);

    const [a3, a4] = await Promise.all([
      page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A3")),
      page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A4"))
    ]);
    expect(a3).toBe("3");
    expect(a4).toBe("4");

    const undoLabel = await page.evaluate(() => (window as any).__formulaApp.getDocument().undoLabel);
    expect(undoLabel).toBe("Fill");

    // Formula case.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.setCellInput(sheetId, "B1", "=A1*2");
      app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 1, endCol: 1 } });
    });
    await waitForIdle(page);

    const handleB1 = await page.evaluate(() => (window as any).__formulaApp.getFillHandleRect());
    expect(handleB1).not.toBeNull();

    const b3Rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("B3"));
    expect(b3Rect).not.toBeNull();

    await dragFromTo(
      page,
      { x: gridBox!.x + handleB1!.x + handleB1!.width / 2, y: gridBox!.y + handleB1!.y + handleB1!.height / 2 },
      { x: gridBox!.x + b3Rect!.x + b3Rect!.width / 2, y: gridBox!.y + b3Rect!.y + b3Rect!.height / 2 }
    );
    await waitForIdle(page);

    const [b2Value, b3Value, b2Formula, b3Formula] = await Promise.all([
      page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B2")),
      page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B3")),
      page.evaluate(() => (window as any).__formulaApp.getDocument().getCell((window as any).__formulaApp.getCurrentSheetId(), "B2").formula),
      page.evaluate(() => (window as any).__formulaApp.getDocument().getCell((window as any).__formulaApp.getCurrentSheetId(), "B3").formula),
    ]);

    expect(b2Formula).toBe("=A2*2");
    expect(b3Formula).toBe("=A3*2");
    expect(b2Value).toBe("4");
    expect(b3Value).toBe("6");
  });
});
