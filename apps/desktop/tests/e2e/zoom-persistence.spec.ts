import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("shared grid zoom persistence", () => {
  // This test reloads the desktop shell multiple times; give it extra headroom on slower CI runners.
  test.describe.configure({ timeout: 120_000 });

  test("persists zoom across reloads", async ({ page }) => {
    // Start from a clean storage state so prior runs (or dev sessions) don't influence the test.
    //
    // Use `sessionStorage` as a per-tab guard so the cleanup runs only on the first navigation;
    // otherwise we'd clear the persisted zoom again before the reload assertion.
    await page.addInitScript(() => {
      try {
        if (sessionStorage.getItem("__formula_zoom_persistence_cleared") !== "1") {
          localStorage.clear();
          sessionStorage.setItem("__formula_zoom_persistence_cleared", "1");
        }
      } catch {
        // ignore storage errors (disabled/quota/etc.)
      }
    });

    await gotoDesktop(page, "/?grid=shared");

    const zoomControl = page.getByTestId("zoom-control");

    // Wait for main.ts to wire the status bar control and sync it to the app zoom.
    await expect(zoomControl).toHaveValue("100");

    await zoomControl.selectOption("75");

    await page.waitForFunction(() => {
      const zoom = (window.__formulaApp as any)?.getZoom?.();
      return typeof zoom === "number" && Math.abs(zoom - 0.75) < 0.01;
    });

    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    await expect(zoomControl).toHaveValue("75");

    const zoomAfter = await page.evaluate(() => (window.__formulaApp as any).getZoom());
    expect(zoomAfter).toBeCloseTo(0.75, 2);
  });
});
