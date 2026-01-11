import { expect, test } from "@playwright/test";

test("copies and pastes a rectangular grid selection via TSV clipboard payload", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const selectionCanvas = page.getByTestId("canvas-grid-selection");
  await expect(selectionCanvas).toBeVisible({ timeout: 30_000 });

  const box = await selectionCanvas.boundingBox();
  expect(box).not.toBeNull();

  // Defaults from `VirtualScrollManager`: col width = 100, row height = 21.
  const headerWidth = 100;
  const headerHeight = 21;
  const colWidth = 100;
  const rowHeight = 21;

  const a1X = box!.x + headerWidth + colWidth / 2;
  const a1Y = box!.y + headerHeight + rowHeight / 2;
  const a2Y = a1Y + rowHeight;

  // Select A1:A2 (workbook initializes A1=1, A2=2).
  await page.mouse.move(a1X, a1Y);
  await page.mouse.down();
  await page.mouse.move(a1X, a2Y);
  await page.mouse.up();

  await page.keyboard.press("ControlOrMeta+C");

  const c1X = a1X + colWidth * 2;
  await selectionCanvas.click({ position: { x: c1X - box!.x, y: a1Y - box!.y } });
  await expect(page.getByTestId("active-address")).toHaveText("C1");

  await page.keyboard.press("ControlOrMeta+V");

  await expect(page.getByTestId("formula-bar-value")).toHaveText("1");

  await selectionCanvas.click({ position: { x: c1X - box!.x, y: a2Y - box!.y } });
  await expect(page.getByTestId("active-address")).toHaveText("C2");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("2");
});
