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
    await page.goto("/?e2e=1");

    await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });
    await page.waitForFunction(() => (window as any).__gridApi?.getCellRect, null, { timeout: 30_000 });

    const selectionCanvas = page.getByTestId("canvas-grid-selection");

    // Click B3 (grid includes 1 header row + 1 header col).
    const b3Rect = await page.evaluate(() => (window as any).__gridApi.getCellRect(3, 2));
    expect(b3Rect).not.toBeNull();
    await selectionCanvas.click({
      position: { x: b3Rect.x + b3Rect.width / 2, y: b3Rect.y + b3Rect.height / 2 },
    });

    await expect(page.getByTestId("active-address")).toHaveText("B3");

    // Run Freeze Panes via command palette.
    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await page.keyboard.type("Freeze Panes");
    await page.keyboard.press("Enter");

    // Grid frozen counts include the header row/col.
    await expect
      .poll(async () => {
        return await page.evaluate(() => {
          const api = (window as any).__gridApi;
          const viewport = api?.getViewportState?.();
          return viewport ? { frozenRows: viewport.frozenRows, frozenCols: viewport.frozenCols } : null;
        });
      })
      .toEqual({ frozenRows: 3, frozenCols: 2 });

    await expect(page.getByText("Frozen: 2 rows, 1 cols")).toBeVisible();

    const a1Before = await page.evaluate(() => (window as any).__gridApi.getCellRect(1, 1));
    const a2Before = await page.evaluate(() => (window as any).__gridApi.getCellRect(2, 1));
    expect(a1Before).not.toBeNull();
    expect(a2Before).not.toBeNull();

    // Scroll deep into the sheet in both directions.
    await page.evaluate(() => {
      (window as any).__gridApi.scrollToCell(200, 50, { align: "center" });
    });

    await expect
      .poll(async () => await page.evaluate(() => (window as any).__gridApi.getScroll().y))
      .toBeGreaterThan(0);
    await expect
      .poll(async () => await page.evaluate(() => (window as any).__gridApi.getScroll().x))
      .toBeGreaterThan(0);

    const a1After = await page.evaluate(() => (window as any).__gridApi.getCellRect(1, 1));
    const a2After = await page.evaluate(() => (window as any).__gridApi.getCellRect(2, 1));
    expect(a1After).not.toBeNull();
    expect(a2After).not.toBeNull();

    // Frozen rows/cols should not move as scroll changes.
    expect(a1After.x).toBeCloseTo(a1Before.x, 2);
    expect(a1After.y).toBeCloseTo(a1Before.y, 2);
    expect(a2After.x).toBeCloseTo(a2Before.x, 2);
    expect(a2After.y).toBeCloseTo(a2Before.y, 2);

    // Clicking within the frozen region should still select A1/A2.
    await selectionCanvas.click({
      position: { x: a1After.x + a1After.width / 2, y: a1After.y + a1After.height / 2 },
    });
    await expect(page.getByTestId("active-address")).toHaveText("A1");

    await selectionCanvas.click({
      position: { x: a2After.x + a2After.width / 2, y: a2After.y + a2After.height / 2 },
    });
    await expect(page.getByTestId("active-address")).toHaveText("A2");

    // Clicking in the scrollable region should select the correct cell.
    const targetRow = 200;
    const targetCol = 50;
    const targetRect = await page.evaluate(
      ({ row, col }) => (window as any).__gridApi.getCellRect(row, col),
      { row: targetRow, col: targetCol },
    );
    expect(targetRect).not.toBeNull();

    await selectionCanvas.click({
      position: { x: targetRect.x + targetRect.width / 2, y: targetRect.y + targetRect.height / 2 },
    });

    const headerRows = 1;
    const headerCols = 1;
    const expectedRow0 = targetRow - headerRows;
    const expectedCol0 = targetCol - headerCols;
    const expected = `${colToName(expectedCol0)}${expectedRow0 + 1}`;
    await expect(page.getByTestId("active-address")).toHaveText(expected);
  });
});

