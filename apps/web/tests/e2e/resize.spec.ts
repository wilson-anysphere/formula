import { expect, test } from "@playwright/test";

test("column resize drag + double-click auto-fit updates layout", async ({ page }) => {
  await page.goto("/?e2e=1");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  await page.waitForFunction(() => {
    const api = (window as any).__gridApi;
    return api && typeof api.getCellRect === "function" && typeof api.getColWidth === "function";
  });

  await page.evaluate(() => {
    (window as any).__gridApi.scrollTo(0, 0);
  });

  const selectionCanvas = page.getByTestId("canvas-grid-selection");
  const canvasBox = await selectionCanvas.boundingBox();
  expect(canvasBox).not.toBeNull();

  const beforeColB = await page.evaluate(() => (window as any).__gridApi.getCellRect(0, 2));
  expect(beforeColB).not.toBeNull();

  const colARect = await page.evaluate(() => (window as any).__gridApi.getCellRect(0, 1));
  expect(colARect).not.toBeNull();

  const boundaryX = canvasBox!.x + colARect!.x + colARect!.width;
  const boundaryY = canvasBox!.y + colARect!.y + colARect!.height / 2;

  // Drag the A|B boundary to widen column A.
  await page.mouse.move(boundaryX, boundaryY);
  await page.mouse.down();
  await page.mouse.move(boundaryX + 80, boundaryY);
  await page.mouse.up();

  const afterColB = await page.evaluate(() => (window as any).__gridApi.getCellRect(0, 2));
  expect(afterColB).not.toBeNull();
  expect(afterColB!.x).toBeGreaterThan(beforeColB!.x + 50);

  const widthAfterDrag = await page.evaluate(() => (window as any).__gridApi.getColWidth(1));

  const colARectAfterDrag = await page.evaluate(() => (window as any).__gridApi.getCellRect(0, 1));
  expect(colARectAfterDrag).not.toBeNull();

  const boundaryXAfterDrag = canvasBox!.x + colARectAfterDrag!.x + colARectAfterDrag!.width;
  const boundaryYAfterDrag = canvasBox!.y + colARectAfterDrag!.y + colARectAfterDrag!.height / 2;

  // Double click the boundary to auto-fit column A.
  await page.mouse.click(boundaryXAfterDrag, boundaryYAfterDrag, { clickCount: 2 });

  const widthAfterAutoFit = await page.evaluate(() => (window as any).__gridApi.getColWidth(1));
  expect(widthAfterAutoFit).not.toBe(widthAfterDrag);
});

test("row resize drag + double-click auto-fit updates layout", async ({ page }) => {
  await page.goto("/?e2e=1");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  await page.waitForFunction(() => {
    const api = (window as any).__gridApi;
    return api && typeof api.getCellRect === "function" && typeof api.getRowHeight === "function";
  });

  await page.evaluate(() => {
    (window as any).__gridApi.scrollTo(0, 0);
  });

  const selectionCanvas = page.getByTestId("canvas-grid-selection");
  const canvasBox = await selectionCanvas.boundingBox();
  expect(canvasBox).not.toBeNull();

  const beforeRow2 = await page.evaluate(() => (window as any).__gridApi.getCellRect(2, 1));
  expect(beforeRow2).not.toBeNull();

  const row1HeaderRect = await page.evaluate(() => (window as any).__gridApi.getCellRect(1, 0));
  expect(row1HeaderRect).not.toBeNull();

  const boundaryX = canvasBox!.x + row1HeaderRect!.x + row1HeaderRect!.width / 2;
  const boundaryY = canvasBox!.y + row1HeaderRect!.y + row1HeaderRect!.height;

  // Drag the row 1|2 boundary to increase row 1 height.
  await page.mouse.move(boundaryX, boundaryY);
  await page.mouse.down();
  await page.mouse.move(boundaryX, boundaryY + 60);
  await page.mouse.up();

  const afterRow2 = await page.evaluate(() => (window as any).__gridApi.getCellRect(2, 1));
  expect(afterRow2).not.toBeNull();
  expect(afterRow2!.y).toBeGreaterThan(beforeRow2!.y + 40);

  const heightAfterDrag = await page.evaluate(() => (window as any).__gridApi.getRowHeight(1));

  const row1HeaderRectAfterDrag = await page.evaluate(() => (window as any).__gridApi.getCellRect(1, 0));
  expect(row1HeaderRectAfterDrag).not.toBeNull();

  const boundaryXAfterDrag = canvasBox!.x + row1HeaderRectAfterDrag!.x + row1HeaderRectAfterDrag!.width / 2;
  const boundaryYAfterDrag = canvasBox!.y + row1HeaderRectAfterDrag!.y + row1HeaderRectAfterDrag!.height;

  // Double click the boundary to auto-fit row 1.
  await page.mouse.click(boundaryXAfterDrag, boundaryYAfterDrag, { clickCount: 2 });

  const heightAfterAutoFit = await page.evaluate(() => (window as any).__gridApi.getRowHeight(1));
  expect(heightAfterAutoFit).not.toBe(heightAfterDrag);
});
