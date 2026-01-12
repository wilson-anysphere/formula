import { expect, test } from "@playwright/test";

test("dragging the fill handle extends a numeric series", async ({ page }) => {
  await page.goto("/?e2e");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });
  await page.waitForFunction(() => (window as any).__gridApi != null);

  const grid = page.getByTestId("canvas-grid-selection");
  const input = page.getByTestId("formula-input");

  const clickCell = async (row: number, col: number) => {
    const position = await page.evaluate(
      ({ row, col }) => {
        const api = (window as any).__gridApi;
        const rect = api.getCellRect(row, col);
        return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
      },
      { row, col }
    );
    await grid.click({ position });
  };

  // Seed A1=1, A2=2.
  await clickCell(1, 1);
  await expect(page.getByTestId("active-address")).toHaveText("A1");
  await input.fill("1");
  await input.press("Enter");

  await clickCell(2, 1);
  await expect(page.getByTestId("active-address")).toHaveText("A2");
  await input.fill("2");
  await input.press("Enter");

  // Select A1:A2 and drag the fill handle down to A4.
  await page.evaluate(() => {
    (window as any).__gridApi.setSelectionRange({ startRow: 1, endRow: 3, startCol: 1, endCol: 2 });
  });

  const { start, end } = await page.evaluate(() => {
    const api = (window as any).__gridApi;
    const handleRect = api.getFillHandleRect();
    if (!handleRect) throw new Error("Fill handle not visible");
    const targetRect = api.getCellRect(4, 1); // A4
    return {
      start: { x: handleRect.x + handleRect.width / 2, y: handleRect.y + handleRect.height / 2 },
      end: { x: targetRect.x + targetRect.width / 2, y: targetRect.y + targetRect.height / 2 }
    };
  });

  const box = await grid.boundingBox();
  if (!box) throw new Error("Grid not visible");

  await page.mouse.move(box.x + start.x, box.y + start.y);
  await page.mouse.down();
  await page.mouse.move(box.x + end.x, box.y + end.y);
  await page.mouse.up();

  // Verify A3=3, A4=4.
  await clickCell(3, 1);
  await expect(page.getByTestId("formula-bar-value")).toHaveText("3");
  await clickCell(4, 1);
  await expect(page.getByTestId("formula-bar-value")).toHaveText("4");
});
