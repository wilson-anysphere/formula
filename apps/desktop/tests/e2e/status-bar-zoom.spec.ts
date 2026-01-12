import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 60_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("status bar zoom", () => {
  test("zoom control is disabled in legacy grid mode", async ({ page }) => {
    await gotoDesktop(page, "/?grid=legacy");
    await waitForIdle(page);
    await expect(page.getByTestId("zoom-control")).toBeDisabled();
    await expect(page.getByTestId("zoom-control")).toHaveValue("100");
  });

  test("zoom control updates shared grid zoom + cell rects", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);

    const zoomControl = page.getByTestId("zoom-control");
    await expect(zoomControl).not.toBeDisabled();
    await expect(zoomControl).toHaveValue("100");

    const before = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));
    if (!before) throw new Error("Missing A1 rect at zoom 1");

    await zoomControl.selectOption("200");

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getZoom())).toBe(2);
    await expect(zoomControl).toHaveValue("200");

    const after = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));
    if (!after) throw new Error("Missing A1 rect after zoom change");

    // Allow some tolerance due to device pixel ratio rounding, but ensure we actually zoomed.
    expect(after.width).toBeGreaterThan(before.width * 1.5);
    expect(after.height).toBeGreaterThan(before.height * 1.5);
  });

  test("ctrl+wheel zoom gesture updates status bar zoom", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);

    const zoomControl = page.getByTestId("zoom-control");
    await expect(zoomControl).toHaveValue("100");

    const gridBox = await page.locator("#grid").boundingBox();
    if (!gridBox) throw new Error("Missing grid bounding box");
    await page.mouse.move(gridBox.x + gridBox.width / 2, gridBox.y + gridBox.height / 2);

    await page.keyboard.down("Control");
    // Large delta to ensure a visible zoom change regardless of platform wheel scale.
    await page.mouse.wheel(0, -600);
    await page.keyboard.up("Control");

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getZoom())).toBeGreaterThan(1.2);
    await expect(zoomControl).not.toHaveValue("100");
  });
});
