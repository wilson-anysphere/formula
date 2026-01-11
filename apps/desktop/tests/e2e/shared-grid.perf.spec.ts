import { expect, test } from "@playwright/test";

// Opt-in perf smoke test. Run with `PERF_TESTS=1 pnpm -C apps/desktop test:e2e`.
test.skip(process.env.PERF_TESTS !== "1", "PERF_TESTS not enabled");

test("shared grid scroll perf smoke", async ({ page }) => {
  await page.goto("/?grid=shared");

  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
  await page.evaluate(() => (window as any).__formulaApp.setGridPerfStatsEnabled(true));

  const grid = page.locator("#grid");
  const box = await grid.boundingBox();
  if (!box) throw new Error("Missing grid bounding box");

  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);

  // Generate a handful of large scroll events to exercise rendering.
  for (let i = 0; i < 12; i++) {
    await page.mouse.wheel(0, 2000);
  }

  // Give the renderer a frame to settle.
  await page.waitForTimeout(50);

  const stats = await page.evaluate(() => (window as any).__formulaApp.getGridPerfStats());
  expect(stats).toBeTruthy();

  const lastFrameMs = (stats as any).lastFrameMs;
  expect(typeof lastFrameMs).toBe("number");
  // This threshold is intentionally generous for shared CI runners; it catches catastrophic regressions.
  expect(lastFrameMs).toBeLessThan(250);
});

