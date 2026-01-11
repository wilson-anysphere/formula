import { expect, test } from "@playwright/test";

test("dragging the fill handle fills series and shifts formulas", async ({ page }) => {
  // `?e2e=1` exposes the grid API on `window.__gridApi` for robust coordinate-based interactions.
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

  const fillHandlePoint = async (row0: number, col0: number) => {
    const rect = await getCellRect(row0, col0);
    expect(rect).not.toBeNull();
    // Slightly inside the 8x8 handle that straddles the cell bottom-right corner.
    return { x: box!.x + rect!.x + rect!.width + 2, y: box!.y + rect!.y + rect!.height + 2 };
  };

  const input = page.getByTestId("formula-input");

  // Seed A1=1, A2=2.
  const a1Center = await cellCenter(0, 0);
  await page.mouse.click(a1Center.x, a1Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("A1");
  await input.fill("1");
  await input.press("Enter");

  const a2Center = await cellCenter(1, 0);
  await page.mouse.click(a2Center.x, a2Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("A2");
  await input.fill("2");
  await input.press("Enter");

  // Select A1:A2.
  await page.mouse.move(a1Center.x, a1Center.y);
  await page.mouse.down();
  await page.mouse.move(a2Center.x, a2Center.y);
  await page.mouse.up();

  // Drag the fill handle down to A4.
  const a2Handle = await fillHandlePoint(1, 0);
  const a4Center = await cellCenter(3, 0);
  await page.mouse.move(a2Handle.x, a2Handle.y);
  await page.mouse.down();
  await page.mouse.move(a2Handle.x, a4Center.y);
  await page.mouse.up();

  // Series fill (1,2,3,4...).
  const a3Center = await cellCenter(2, 0);
  await page.mouse.click(a3Center.x, a3Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("A3");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("3");

  await page.mouse.click(a4Center.x, a4Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("A4");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("4");

  // Formula shifting: B1 = A1*2 -> fill down to B3 shifts the referenced row.
  const b1Center = await cellCenter(0, 1);
  await page.mouse.click(b1Center.x, b1Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("B1");
  // Ensure the formula bar has synced to B1 before starting to edit. Under heavy
  // Playwright parallelism, the async cell sync can race with `fill()` and
  // append the new formula onto the old one (e.g. `=A1+A2=A1*2`).
  await expect(input).toHaveValue("=A1+A2");
  await input.fill("=A1*2");
  await input.press("Enter");
  await expect(input).toHaveValue("=A1*2");

  // Exit formula editing mode so grid interactions are in "default" mode.
  await page.getByTestId("engine-status").click();
  await expect(input).not.toBeFocused();

  // Re-select B1 so the fill handle is visible.
  await page.mouse.click(b1Center.x, b1Center.y);

  const b1Handle = await fillHandlePoint(0, 1);
  const b3Center = await cellCenter(2, 1);
  await page.mouse.move(b1Handle.x, b1Handle.y);
  await page.mouse.down();
  await page.mouse.move(b1Handle.x, b3Center.y);
  await page.mouse.up();

  const b2Center = await cellCenter(1, 1);
  await page.mouse.click(b2Center.x, b2Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("B2");
  await expect(input).toHaveValue("=A2*2");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("4");

  await page.mouse.click(b3Center.x, b3Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("B3");
  await expect(input).toHaveValue("=A3*2");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("6");
});
