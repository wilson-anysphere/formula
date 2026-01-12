import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("shared grid zoom persistence", () => {
  // This test reloads the desktop shell multiple times; give it extra headroom on slower CI runners.
  test.describe.configure({ timeout: 120_000 });

  test("persists zoom across reloads", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");

    // Start from a clean storage state so prior runs (or dev sessions) don't influence the test.
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    const zoomControl = page.getByTestId("zoom-control");

    // Wait for main.ts to wire the status bar control and sync it to the app zoom.
    await expect(zoomControl).toHaveValue("100");

    await zoomControl.selectOption("75");

    await page.waitForFunction(() => {
      const zoom = (window as any).__formulaApp?.getZoom?.();
      return typeof zoom === "number" && Math.abs(zoom - 0.75) < 0.01;
    });

    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    await expect(zoomControl).toHaveValue("75");

    const zoomAfter = await page.evaluate(() => (window as any).__formulaApp.getZoom());
    expect(zoomAfter).toBeCloseTo(0.75, 2);
  });
});
