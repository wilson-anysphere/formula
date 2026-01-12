import { expect, test } from "@playwright/test";

test("fill handle shifts full-row/column references using engine semantics", async ({ page }) => {
  test.setTimeout(120_000);

  await page.goto("/?e2e=1");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });
  await page.waitForFunction(() => (window as any).__gridApi != null);

  const grid = page.getByTestId("grid");
  const selectionCanvas = grid.locator("canvas").nth(2);
  await expect(selectionCanvas).toBeVisible();

  const box = await selectionCanvas.boundingBox();
  expect(box).not.toBeNull();

  const getCellRect = async (row0: number, col0: number) =>
    page.evaluate(
      ({ row, col }) => (window as any).__gridApi.getCellRect(row, col),
      // +1 to account for the frozen header row/col.
      { row: row0 + 1, col: col0 + 1 }
    );

  const cellCenter = async (row0: number, col0: number) => {
    const rect = await getCellRect(row0, col0);
    expect(rect).not.toBeNull();
    return { x: box!.x + rect!.x + rect!.width / 2, y: box!.y + rect!.y + rect!.height / 2 };
  };

  const fillHandleCenter = async () => {
    await page.waitForFunction(() => (window as any).__gridApi.getFillHandleRect() != null);
    const rect = await page.evaluate(() => (window as any).__gridApi.getFillHandleRect());
    expect(rect).not.toBeNull();
    return { x: box!.x + rect!.x + rect!.width / 2, y: box!.y + rect!.y + rect!.height / 2 };
  };

  const input = page.getByTestId("formula-input");

  // Seed D1 with a full-column reference and fill right to E1.
  const d1 = await cellCenter(0, 3);
  const e1 = await cellCenter(0, 4);

  await page.mouse.click(d1.x, d1.y);
  await expect(page.getByTestId("active-address")).toHaveText("D1");
  await expect(input).toHaveValue("");
  await input.fill("=SUM(A:A)");
  await input.press("Enter");
  await expect(input).toHaveValue("=SUM(A:A)");

  // Exit formula editing mode so the grid is in "default" interaction mode (fill handle enabled).
  await page.getByTestId("engine-status").click();
  await expect(input).not.toBeFocused();

  await page.mouse.click(d1.x, d1.y);
  await expect(page.getByTestId("active-address")).toHaveText("D1");

  const d1Handle = await fillHandleCenter();
  await page.mouse.move(d1Handle.x, d1Handle.y);
  await page.mouse.down();
  await page.mouse.move(e1.x, d1Handle.y);
  await page.mouse.up();

  await page.mouse.click(e1.x, e1.y);
  await expect(page.getByTestId("active-address")).toHaveText("E1");
  await expect(input).toHaveValue("=SUM(B:B)");

  // Seed D3 with a full-row reference and fill down to D4.
  const d3 = await cellCenter(2, 3);
  const d4 = await cellCenter(3, 3);

  await page.mouse.click(d3.x, d3.y);
  await expect(page.getByTestId("active-address")).toHaveText("D3");
  await expect(input).toHaveValue("");
  await input.fill("=SUM(1:1)");
  await input.press("Enter");
  await expect(input).toHaveValue("=SUM(1:1)");

  await page.getByTestId("engine-status").click();
  await expect(input).not.toBeFocused();

  await page.mouse.click(d3.x, d3.y);
  await expect(page.getByTestId("active-address")).toHaveText("D3");

  const d3Handle = await fillHandleCenter();
  await page.mouse.move(d3Handle.x, d3Handle.y);
  await page.mouse.down();
  await page.mouse.move(d4.x, d4.y);
  await page.mouse.up();

  await page.mouse.click(d4.x, d4.y);
  await expect(page.getByTestId("active-address")).toHaveText("D4");
  await expect(input).toHaveValue("=SUM(2:2)");

  // Spill postfix: B6="=A6#" -> fill left to A6 should become "=#REF!" (spill range operator dropped).
  const b6 = await cellCenter(5, 1);
  const a6 = await cellCenter(5, 0);

  await page.mouse.click(b6.x, b6.y);
  await expect(page.getByTestId("active-address")).toHaveText("B6");
  await expect(input).toHaveValue("");
  await input.fill("=A6#");
  await input.press("Enter");
  await expect(input).toHaveValue("=A6#");

  await page.getByTestId("engine-status").click();
  await expect(input).not.toBeFocused();

  await page.mouse.click(b6.x, b6.y);
  await expect(page.getByTestId("active-address")).toHaveText("B6");

  const b6Handle = await fillHandleCenter();
  await page.mouse.move(b6Handle.x, b6Handle.y);
  await page.mouse.down();
  await page.mouse.move(a6.x, b6Handle.y);
  await page.mouse.up();

  await page.mouse.click(a6.x, a6.y);
  await expect(page.getByTestId("active-address")).toHaveText("A6");
  await expect(input).toHaveValue("=#REF!");
});
