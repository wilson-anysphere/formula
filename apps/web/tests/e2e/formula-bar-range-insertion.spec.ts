import { expect, test } from "@playwright/test";

test("inserts dragged range into formula bar at cursor", async ({ page }) => {
  await page.goto("/");

  const input = page.getByTestId("formula-input");
  await input.click();
  await input.type("=SUM(");

  const grid = page.getByTestId("grid");
  await expect(grid).toBeVisible();

  // CanvasGrid renders three canvases (grid/content/selection); the selection canvas handles pointer events.
  const selectionCanvas = grid.locator("canvas").nth(2);
  await expect(selectionCanvas).toBeVisible({ timeout: 30_000 });

  const box = await selectionCanvas.boundingBox();
  expect(box).not.toBeNull();

  // Defaults from `VirtualScrollManager`: col width = 100, row height = 21.
  const headerWidth = 100;
  const headerHeight = 21;
  const colWidth = 100;
  const rowHeight = 21;

  const startX = box!.x + headerWidth + colWidth / 2;
  const startY = box!.y + headerHeight + rowHeight / 2;
  const endY = startY + rowHeight;

  await page.mouse.move(startX, startY);
  await page.mouse.down();
  await page.mouse.move(startX, endY);
  await page.mouse.up();

  await expect(input).toHaveValue("=SUM(A1:A2");
});
