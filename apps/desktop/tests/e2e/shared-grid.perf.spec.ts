import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

// Opt-in perf smoke test. Run with `PERF_TESTS=1 pnpm -C apps/desktop test:e2e`.
test.skip(process.env.PERF_TESTS !== "1", "PERF_TESTS not enabled");

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test("shared grid scroll perf smoke", async ({ page }) => {
  await gotoDesktop(page, "/?grid=shared");

  await waitForIdle(page);
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

test("shared grid million-row deep scroll perf smoke (A900001)", async ({ page }) => {
  await gotoDesktop(page, "/?grid=shared");

  await waitForIdle(page);
  await page.evaluate(() => (window as any).__formulaApp.setGridPerfStatsEnabled(true));

  const grid = page.locator("#grid");
  const gridBox = await grid.boundingBox();
  if (!gridBox) throw new Error("Missing grid bounding box");

  // Jump to a deep row (0-based doc coords); should trigger scrollCellIntoView + render at a large scroll offset.
  await page.evaluate(() => (window as any).__formulaApp.activateCell({ row: 900_000, col: 0 }));
  await waitForIdle(page);

  // Give the renderer a frame to settle.
  await page.waitForTimeout(100);

  await expect
    .poll(async () => {
      return await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    })
    .toBeGreaterThan(1_000_000);

  await expect
    .poll(async () => {
      const rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A900001"));
      if (!rect) return false;
      const typedRect = rect as { x: number; y: number; width: number; height: number };
      // `getCellRectA1` returns viewport-relative coords. Ensure the cell overlaps the visible grid region.
      return (
        typedRect.x < gridBox.width &&
        typedRect.y < gridBox.height &&
        typedRect.x + typedRect.width > 0 &&
        typedRect.y + typedRect.height > 0
      );
    })
    .toBe(true);

  const stats = await page.evaluate(() => (window as any).__formulaApp.getGridPerfStats());
  expect(stats).toBeTruthy();

  const lastFrameMs = (stats as any).lastFrameMs;
  expect(typeof lastFrameMs).toBe("number");
  // Intentionally generous; this should catch catastrophic O(maxRows) regressions without being too flaky in CI.
  expect(lastFrameMs).toBeLessThan(500);
});
