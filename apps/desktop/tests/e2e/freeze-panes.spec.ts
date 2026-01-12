import { expect, test } from "@playwright/test";

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
    await page.goto("/");

    const grid = page.locator("#grid");

    // Select B3 (accounting for headers; SpreadsheetApp uses fixed 100x24 cells).
    await grid.click({ position: { x: 48 + 100 + 50, y: 24 + 2 * 24 + 12 } });
    await expect(page.getByTestId("active-cell")).toHaveText("B3");

    // Open command palette and run "Freeze Panes".
    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await page.keyboard.type("Freeze Panes");
    await page.keyboard.press("Enter");

    await expect.poll(async () => {
      return await page.evaluate(() => (window as any).__formulaApp.getFrozen());
    }).toEqual({ frozenRows: 2, frozenCols: 1 });

    // Scroll deep into the sheet in both directions.
    await grid.hover({ position: { x: 200, y: 200 } });
    await page.mouse.wheel(0, 200 * 24);
    await page.mouse.wheel(20 * 100, 0);

    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().y)).toBeGreaterThan(0);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().x)).toBeGreaterThan(0);

    // Clicking within the frozen region should still select A1/A2.
    await grid.click({ position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await grid.click({ position: { x: 60, y: 60 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A2");

    // Clicking in the scrollable region should select the scrolled cell, not a frozen one.
    const clickX = 48 + 100 + 50;
    const clickY = 24 + 2 * 24 + 12;

    const { scroll, frozen } = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      return { scroll: app.getScroll(), frozen: app.getFrozen() };
    });

    const frozenWidth = frozen.frozenCols * 100;
    const frozenHeight = frozen.frozenRows * 24;
    const localX = clickX - 48;
    const localY = clickY - 24;
    const sheetX = clickX < 48 + frozenWidth ? localX : scroll.x + localX;
    const sheetY = clickY < 24 + frozenHeight ? localY : scroll.y + localY;
    const expectedCol = Math.floor(sheetX / 100);
    const expectedRow = Math.floor(sheetY / 24);
    const expectedA1 = `${colToName(expectedCol)}${expectedRow + 1}`;

    await grid.click({ position: { x: clickX, y: clickY } });
    await expect(page.getByTestId("active-cell")).toHaveText(expectedA1);
  });
});
