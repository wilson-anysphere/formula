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
