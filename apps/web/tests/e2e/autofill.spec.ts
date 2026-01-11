import { expect, test } from "@playwright/test";

test("dragging the fill handle repeats values and shifts formulas", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const grid = page.getByTestId("grid");
  const selectionCanvas = grid.locator("canvas").nth(2);
  await expect(selectionCanvas).toBeVisible();

  const box = await selectionCanvas.boundingBox();
  expect(box).not.toBeNull();

  // Defaults from `VirtualScrollManager`: col width = 100, row height = 21.
  const headerWidth = 100;
  const headerHeight = 21;
  const colWidth = 100;
  const rowHeight = 21;

  const cellCenter = (row0: number, col0: number) => ({
    x: box!.x + headerWidth + col0 * colWidth + colWidth / 2,
    y: box!.y + headerHeight + row0 * rowHeight + rowHeight / 2
  });

  const fillHandlePoint = (row0: number, col0: number) => ({
    // Slightly inside the 8x8 handle that straddles the cell bottom-right corner.
    x: box!.x + headerWidth + (col0 + 1) * colWidth + 2,
    y: box!.y + headerHeight + (row0 + 1) * rowHeight + 2
  });

  const input = page.getByTestId("formula-input");

  // Seed A1=1, A2=2.
  await page.mouse.click(cellCenter(0, 0).x, cellCenter(0, 0).y);
  await expect(page.getByTestId("active-address")).toHaveText("A1");
  await input.fill("1");
  await input.press("Enter");

  await page.mouse.click(cellCenter(1, 0).x, cellCenter(1, 0).y);
  await expect(page.getByTestId("active-address")).toHaveText("A2");
  await input.fill("2");
  await input.press("Enter");

  // Select A1:A2.
  await page.mouse.move(cellCenter(0, 0).x, cellCenter(0, 0).y);
  await page.mouse.down();
  await page.mouse.move(cellCenter(1, 0).x, cellCenter(1, 0).y);
  await page.mouse.up();

  // Drag the fill handle down to A4.
  const a2Handle = fillHandlePoint(1, 0);
  const a4Center = cellCenter(3, 0);
  await page.mouse.move(a2Handle.x, a2Handle.y);
  await page.mouse.down();
  await page.mouse.move(a2Handle.x, a4Center.y);
  await page.mouse.up();

  // Series fill (1,2,3,4...).
  await page.mouse.click(cellCenter(2, 0).x, cellCenter(2, 0).y);
  await expect(page.getByTestId("active-address")).toHaveText("A3");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("3");

  await page.mouse.click(cellCenter(3, 0).x, cellCenter(3, 0).y);
  await expect(page.getByTestId("active-address")).toHaveText("A4");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("4");

  // Formula shifting: B1 = A1*2 -> fill down to B3 shifts the referenced row.
  await page.mouse.click(cellCenter(0, 1).x, cellCenter(0, 1).y);
  await expect(page.getByTestId("active-address")).toHaveText("B1");
  await input.fill("=A1*2");
  await input.press("Enter");

  // Exit formula editing mode so grid interactions are in "default" mode.
  await page.getByTestId("engine-status").click();
  await expect(input).not.toBeFocused();

  // Re-select B1 so the fill handle is visible.
  await page.mouse.click(cellCenter(0, 1).x, cellCenter(0, 1).y);

  const b1Handle = fillHandlePoint(0, 1);
  const b3Center = cellCenter(2, 1);
  await page.mouse.move(b1Handle.x, b1Handle.y);
  await page.mouse.down();
  await page.mouse.move(b1Handle.x, b3Center.y);
  await page.mouse.up();

  await page.mouse.click(cellCenter(1, 1).x, cellCenter(1, 1).y);
  await expect(page.getByTestId("active-address")).toHaveText("B2");
  await expect(input).toHaveValue("=A2*2");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("4");

  await page.mouse.click(cellCenter(2, 1).x, cellCenter(2, 1).y);
  await expect(page.getByTestId("active-address")).toHaveText("B3");
  await expect(input).toHaveValue("=A3*2");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("6");
});
