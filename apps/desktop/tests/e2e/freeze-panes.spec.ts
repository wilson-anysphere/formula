import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

function colToName(col0: number): string {
  let value = col0 + 1;
  let name = "";
  while (value > 0) {
    const rem = (value - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

test.describe("freeze panes", () => {
  test("Freeze Panes at B3 keeps rows 1-2 and col A frozen while scrolling", async ({ page }) => {
    await gotoDesktop(page, "/");

    const grid = page.locator("#grid");

    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("B3");
      return rect && typeof rect.x === "number" && rect.width > 0 && rect.height > 0;
    });

    const rects = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      return {
        a1: app.getCellRectA1("A1"),
        a2: app.getCellRectA1("A2"),
        b3: app.getCellRectA1("B3"),
      };
    });

    expect(rects.a1).toBeTruthy();
    expect(rects.a2).toBeTruthy();
    expect(rects.b3).toBeTruthy();

    const b3Rect = rects.b3 as { x: number; y: number; width: number; height: number };

    // Select B3.
    await grid.click({ position: { x: b3Rect.x + b3Rect.width / 2, y: b3Rect.y + b3Rect.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("B3");

    // Open command palette and run "Freeze Panes".
    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await page.keyboard.type("Freeze Panes");
    // Bonus: verify category grouping renders in the list (stable UI affordance).
    await expect(page.getByTestId("command-palette-list")).toContainText("View");
    await page.keyboard.press("Enter");

    await expect.poll(async () => {
      return await page.evaluate(() => (window as any).__formulaApp.getFrozen());
    }).toEqual({ frozenRows: 2, frozenCols: 1 });

    // Capture frozen cell rects before scrolling (these should remain stable).
    const a1RectBefore = rects.a1 as { x: number; y: number; width: number; height: number };
    const a2RectBefore = rects.a2 as { x: number; y: number; width: number; height: number };

    const cellWidth = a1RectBefore.width;
    const cellHeight = a1RectBefore.height;
    const rowHeaderWidth = a1RectBefore.x;
    const colHeaderHeight = a1RectBefore.y;

    // Scroll deep into the sheet in both directions.
    await grid.hover({ position: { x: rowHeaderWidth + cellWidth * 2, y: colHeaderHeight + cellHeight * 6 } });
    await page.mouse.wheel(0, 200 * cellHeight);
    await page.mouse.wheel(20 * cellWidth, 0);

    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().y)).toBeGreaterThan(0);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().x)).toBeGreaterThan(0);

    const a1RectAfterScroll = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));
    const a2RectAfterScroll = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A2"));

    expect(a1RectAfterScroll).toEqual(expect.anything());
    expect(a2RectAfterScroll).toEqual(expect.anything());

    // Frozen cells should not move as scroll changes.
    expect(a1RectAfterScroll.x).toBeCloseTo(a1RectBefore.x, 1);
    expect(a1RectAfterScroll.y).toBeCloseTo(a1RectBefore.y, 1);
    expect(a2RectAfterScroll.x).toBeCloseTo(a2RectBefore.x, 1);
    expect(a2RectAfterScroll.y).toBeCloseTo(a2RectBefore.y, 1);

    // Clicking within the frozen region should still select A1/A2.
    await grid.click({
      position: { x: a1RectAfterScroll.x + a1RectAfterScroll.width / 2, y: a1RectAfterScroll.y + a1RectAfterScroll.height / 2 },
    });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await grid.click({
      position: { x: a2RectAfterScroll.x + a2RectAfterScroll.width / 2, y: a2RectAfterScroll.y + a2RectAfterScroll.height / 2 },
    });
    await expect(page.getByTestId("active-cell")).toHaveText("A2");

    // Clicking in the scrollable region should select the scrolled cell, not a frozen one.
    const clickX = rowHeaderWidth + cellWidth + cellWidth / 2;
    const clickY = colHeaderHeight + 2 * cellHeight + cellHeight / 2;

    const { scroll, frozen } = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      return { scroll: app.getScroll(), frozen: app.getFrozen() };
    });

    const frozenWidth = frozen.frozenCols * cellWidth;
    const frozenHeight = frozen.frozenRows * cellHeight;
    const localX = clickX - rowHeaderWidth;
    const localY = clickY - colHeaderHeight;
    const sheetX = clickX < rowHeaderWidth + frozenWidth ? localX : scroll.x + localX;
    const sheetY = clickY < colHeaderHeight + frozenHeight ? localY : scroll.y + localY;
    const expectedCol = Math.floor(sheetX / cellWidth);
    const expectedRow = Math.floor(sheetY / cellHeight);
    const expectedA1 = `${colToName(expectedCol)}${expectedRow + 1}`;

    await grid.click({ position: { x: clickX, y: clickY } });
    await expect(page.getByTestId("active-cell")).toHaveText(expectedA1);
  });
});
