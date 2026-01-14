import { test, expect } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("canvas chart overlay", () => {
  test("renders charts on the canvas overlay (no .chart-object DOM nodes) and keeps them anchored while scrolling", async ({ page }) => {
    // Canvas charts are the default; this spec asserts the default path renders charts via the
    // drawings overlay without legacy DOM chart hosts.
    await gotoDesktop(page, "/");

    const chartId = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const result = app.addChart({
        chart_type: "bar",
        data_range: "A2:B5",
        title: "Canvas Chart",
        position: "C1",
      });
      return result.chart_id as string;
    });

    await expect(page.locator(".chart-object")).toHaveCount(0);

    const before = await page.evaluate((chartId) => {
      const app = (window as any).__formulaApp;
      return {
        scroll: app.getScroll(),
        rect: app.getChartViewportRect(chartId),
      };
    }, chartId);

    expect(before.rect).not.toBeNull();
    const beforeRect = before.rect as { left: number; top: number; width: number; height: number };
    expect(Number.isFinite(beforeRect.left)).toBe(true);
    expect(Number.isFinite(beforeRect.top)).toBe(true);

    // Vertical scroll should move the chart by the scroll delta.
    const grid = page.locator("#grid");
    await grid.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 240);

    await expect
      .poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().y))
      .toBeGreaterThan(before.scroll.y);

    const after = await page.evaluate((chartId) => {
      const app = (window as any).__formulaApp;
      return {
        scroll: app.getScroll(),
        rect: app.getChartViewportRect(chartId),
      };
    }, chartId);

    expect(after.rect).not.toBeNull();
    const afterRect = after.rect as { left: number; top: number; width: number; height: number };
    const deltaScrollY = after.scroll.y - before.scroll.y;
    expect(Math.abs((beforeRect.top - afterRect.top) - deltaScrollY)).toBeLessThan(1);
    expect(Math.abs(beforeRect.left - afterRect.left)).toBeLessThan(1);

    // Horizontal scroll should move the chart by the scroll delta.
    await page.mouse.wheel(240, 0);

    await expect
      .poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().x))
      .toBeGreaterThan(after.scroll.x);

    const afterX = await page.evaluate((chartId) => {
      const app = (window as any).__formulaApp;
      return {
        scroll: app.getScroll(),
        rect: app.getChartViewportRect(chartId),
      };
    }, chartId);

    expect(afterX.rect).not.toBeNull();
    const afterXRect = afterX.rect as { left: number; top: number; width: number; height: number };
    const deltaScrollX = afterX.scroll.x - after.scroll.x;
    expect(deltaScrollX).toBeGreaterThan(0);
    expect(Math.abs((afterRect.left - afterXRect.left) - deltaScrollX)).toBeLessThan(1);
    expect(Math.abs(afterRect.top - afterXRect.top)).toBeLessThan(1);
  });
});
