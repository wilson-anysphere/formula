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

  const fillHandleCenter = async () => {
    const rect = await page.evaluate(() => (window as any).__gridApi.getFillHandleRect());
    expect(rect).not.toBeNull();
    return { x: box!.x + rect!.x + rect!.width / 2, y: box!.y + rect!.y + rect!.height / 2 };
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
  const a2Handle = await fillHandleCenter();
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

  const b1Handle = await fillHandleCenter();
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

  // Horizontal series fill: C1=1, D1=2 -> fill right to F1.
  const c1Center = await cellCenter(0, 2);
  await page.mouse.click(c1Center.x, c1Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("C1");
  await input.fill("1");
  await input.press("Enter");

  const d1Center = await cellCenter(0, 3);
  await page.mouse.click(d1Center.x, d1Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("D1");
  await input.fill("2");
  await input.press("Enter");

  await page.mouse.move(c1Center.x, c1Center.y);
  await page.mouse.down();
  await page.mouse.move(d1Center.x, d1Center.y);
  await page.mouse.up();

  const d1Handle = await fillHandleCenter();
  const f1Center = await cellCenter(0, 5);
  await page.mouse.move(d1Handle.x, d1Handle.y);
  await page.mouse.down();
  await page.mouse.move(f1Center.x, d1Handle.y);
  await page.mouse.up();

  const e1Center = await cellCenter(0, 4);
  await page.mouse.click(e1Center.x, e1Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("E1");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("3");

  await page.mouse.click(f1Center.x, f1Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("F1");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("4");

  // Fill up: H5=10, H6=12 -> fill up to H3.
  const h5Center = await cellCenter(4, 7);
  await page.mouse.click(h5Center.x, h5Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("H5");
  await input.fill("10");
  await input.press("Enter");

  const h6Center = await cellCenter(5, 7);
  await page.mouse.click(h6Center.x, h6Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("H6");
  await input.fill("12");
  await input.press("Enter");

  await page.mouse.move(h5Center.x, h5Center.y);
  await page.mouse.down();
  await page.mouse.move(h6Center.x, h6Center.y);
  await page.mouse.up();

  const h6Handle = await fillHandleCenter();
  const h3Center = await cellCenter(2, 7);
  await page.mouse.move(h6Handle.x, h6Handle.y);
  await page.mouse.down();
  await page.mouse.move(h6Handle.x, h3Center.y);
  await page.mouse.up();

  await page.mouse.click(h3Center.x, h3Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("H3");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("6");

  const h4Center = await cellCenter(3, 7);
  await page.mouse.click(h4Center.x, h4Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("H4");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("8");

  // Multi-column series: J1:K2 = [[1,2],[3,4]] -> fill down to J4:K4.
  const j1Center = await cellCenter(0, 9);
  await page.mouse.click(j1Center.x, j1Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("J1");
  await input.fill("1");
  await input.press("Enter");

  const k1Center = await cellCenter(0, 10);
  await page.mouse.click(k1Center.x, k1Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("K1");
  await input.fill("2");
  await input.press("Enter");

  const j2Center = await cellCenter(1, 9);
  await page.mouse.click(j2Center.x, j2Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("J2");
  await input.fill("3");
  await input.press("Enter");

  const k2Center = await cellCenter(1, 10);
  await page.mouse.click(k2Center.x, k2Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("K2");
  await input.fill("4");
  await input.press("Enter");

  await page.mouse.move(j1Center.x, j1Center.y);
  await page.mouse.down();
  await page.mouse.move(k2Center.x, k2Center.y);
  await page.mouse.up();

  const k2Handle = await fillHandleCenter();
  const k4Center = await cellCenter(3, 10);
  await page.mouse.move(k2Handle.x, k2Handle.y);
  await page.mouse.down();
  await page.mouse.move(k2Handle.x, k4Center.y);
  await page.mouse.up();

  const j3Center = await cellCenter(2, 9);
  await page.mouse.click(j3Center.x, j3Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("J3");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("5");

  const k3Center = await cellCenter(2, 10);
  await page.mouse.click(k3Center.x, k3Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("K3");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("6");

  const j4Center = await cellCenter(3, 9);
  await page.mouse.click(j4Center.x, j4Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("J4");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("7");

  await page.mouse.click(k4Center.x, k4Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("K4");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("8");

  // Text series: G8="Item 1", G9="Item 3" -> fill down to G11.
  const g8Center = await cellCenter(7, 6);
  await page.mouse.click(g8Center.x, g8Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("G8");
  await expect(input).toHaveValue("");
  await input.fill("Item 1");
  await input.press("Enter");
  await expect(input).toHaveValue("Item 1");

  const g9Center = await cellCenter(8, 6);
  await page.mouse.click(g9Center.x, g9Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("G9");
  await expect(input).toHaveValue("");
  await input.fill("Item 3");
  await input.press("Enter");
  await expect(input).toHaveValue("Item 3");

  await page.mouse.move(g8Center.x, g8Center.y);
  await page.mouse.down();
  await page.mouse.move(g9Center.x, g9Center.y);
  await page.mouse.up();

  const g9Handle = await fillHandleCenter();
  const g11Center = await cellCenter(10, 6);
  await page.mouse.move(g9Handle.x, g9Handle.y);
  await page.mouse.down();
  await page.mouse.move(g9Handle.x, g11Center.y);
  await page.mouse.up();

  const g10Center = await cellCenter(9, 6);
  await page.mouse.click(g10Center.x, g10Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("G10");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("Item 5");

  await page.mouse.click(g11Center.x, g11Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("G11");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("Item 7");

  // Fill left: F2=1, G2=2 -> fill left to D2.
  const f2Center = await cellCenter(1, 5);
  await page.mouse.click(f2Center.x, f2Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("F2");
  await expect(input).toHaveValue("");
  await input.fill("1");
  await input.press("Enter");

  const g2Center = await cellCenter(1, 6);
  await page.mouse.click(g2Center.x, g2Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("G2");
  await expect(input).toHaveValue("");
  await input.fill("2");
  await input.press("Enter");

  await page.mouse.move(f2Center.x, f2Center.y);
  await page.mouse.down();
  await page.mouse.move(g2Center.x, g2Center.y);
  await page.mouse.up();

  const g2Handle = await fillHandleCenter();
  const d2Center = await cellCenter(1, 3);
  await page.mouse.move(g2Handle.x, g2Handle.y);
  await page.mouse.down();
  await page.mouse.move(d2Center.x, g2Handle.y);
  await page.mouse.up();

  await page.mouse.click(d2Center.x, d2Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("D2");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("-1");

  const e2Center = await cellCenter(1, 4);
  await page.mouse.click(e2Center.x, e2Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("E2");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("0");

  // Formula fill right: set A10=1, B10 =A10+1, then fill right to D10.
  const a10Center = await cellCenter(9, 0);
  await page.mouse.click(a10Center.x, a10Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("A10");
  await expect(input).toHaveValue("");
  await input.fill("1");
  await input.press("Enter");

  const b10Center = await cellCenter(9, 1);
  await page.mouse.click(b10Center.x, b10Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("B10");
  await expect(input).toHaveValue("");
  await input.fill("=A10+1");
  await input.press("Enter");
  await expect(input).toHaveValue("=A10+1");

  await page.getByTestId("engine-status").click();
  await expect(input).not.toBeFocused();

  await page.mouse.click(b10Center.x, b10Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("B10");

  const b10Handle = await fillHandleCenter();
  const d10Center = await cellCenter(9, 3);
  await page.mouse.move(b10Handle.x, b10Handle.y);
  await page.mouse.down();
  await page.mouse.move(d10Center.x, b10Handle.y);
  await page.mouse.up();

  const c10Center = await cellCenter(9, 2);
  await page.mouse.click(c10Center.x, c10Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("C10");
  await expect(input).toHaveValue("=B10+1");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("3");

  await page.mouse.click(d10Center.x, d10Center.y);
  await expect(page.getByTestId("active-address")).toHaveText("D10");
  await expect(input).toHaveValue("=C10+1");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("4");
});
