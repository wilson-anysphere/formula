import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

const LAYOUT_KEY = "formula.layout.workbook.local-workbook.v1";

test.describe("split view", () => {
  test("secondary pane mounts a real grid with independent scroll + zoom and persists state", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");

    // Start from a clean persisted layout so the test is deterministic.
    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(3);

    const primaryScrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    const secondaryScrollBefore = Number((await secondary.getAttribute("data-scroll-y")) ?? 0);

    await secondary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 600);

    await expect.poll(async () => Number((await secondary.getAttribute("data-scroll-y")) ?? 0)).toBeGreaterThan(secondaryScrollBefore);

    const primaryScrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(primaryScrollAfter).toBe(primaryScrollBefore);

    await expect
      .poll(async () => {
        return await page.evaluate((key) => {
          const raw = localStorage.getItem(key);
          if (!raw) return 0;
          const layout = JSON.parse(raw);
          return layout?.splitView?.panes?.secondary?.scrollY ?? 0;
        }, LAYOUT_KEY);
      })
      .toBeGreaterThan(0);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    const zoomBefore = Number((await secondary.getAttribute("data-zoom")) ?? 1);

    await page.keyboard.down(modifier);
    await page.mouse.wheel(0, -200);
    await page.keyboard.up(modifier);

    await expect.poll(async () => Number((await secondary.getAttribute("data-zoom")) ?? 1)).not.toBe(zoomBefore);

    await expect
      .poll(async () => {
        return await page.evaluate((key) => {
          const raw = localStorage.getItem(key);
          if (!raw) return 1;
          const layout = JSON.parse(raw);
          return layout?.splitView?.panes?.secondary?.zoom ?? 1;
        }, LAYOUT_KEY);
      })
      .not.toBe(1);

    const persisted = await page.evaluate((key) => {
      const raw = localStorage.getItem(key);
      if (!raw) return null;
      return JSON.parse(raw);
    }, LAYOUT_KEY);
    expect(persisted?.splitView?.direction).toBe("vertical");

    const persistedScrollY = persisted?.splitView?.panes?.secondary?.scrollY ?? 0;
    const persistedZoom = persisted?.splitView?.panes?.secondary?.zoom ?? 1;
    expect(persistedScrollY).toBeGreaterThan(0);
    expect(persistedZoom).not.toBe(1);

    // Reload and ensure split state + scroll/zoom restore.
    await page.reload();
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await expect(page.locator("#grid-secondary")).toBeVisible();
    await expect(page.locator("#grid-secondary canvas")).toHaveCount(3);

    await expect
      .poll(async () => Number((await page.locator("#grid-secondary").getAttribute("data-scroll-y")) ?? 0))
      .toBeCloseTo(persistedScrollY, 1);
    await expect
      .poll(async () => Number((await page.locator("#grid-secondary").getAttribute("data-zoom")) ?? 1))
      .toBeCloseTo(persistedZoom, 2);
  });
});

test.describe("split view / shared grid zoom", () => {
  test("Ctrl/Cmd+wheel zoom changes grid geometry", async ({ page }) => {
    await page.goto("/?grid=shared");

    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("B1");
      return rect && typeof rect.x === "number" && rect.width > 0 && rect.height > 0;
    });

    const rectsBefore = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      return {
        a1: app.getCellRectA1("A1"),
        b1: app.getCellRectA1("B1"),
      };
    });

    expect(rectsBefore.a1).toBeTruthy();
    expect(rectsBefore.b1).toBeTruthy();

    const a1Before = rectsBefore.a1 as { x: number; y: number; width: number; height: number };
    const b1Before = rectsBefore.b1 as { x: number; y: number; width: number; height: number };

    // Dispatch a ctrl+wheel event directly (avoid Playwright actionability checks around
    // visibility/stability; we only care that the handler updates zoom + geometry).
    await page.evaluate(() => {
      const grid = document.querySelector("#grid");
      if (!grid) throw new Error("Missing #grid");
      grid.dispatchEvent(
        new WheelEvent("wheel", {
          deltaY: -100,
          deltaMode: 0,
          ctrlKey: true,
          bubbles: true,
          cancelable: true,
          // Note: client coords don't matter for this assertion (we only assert geometry changes).
          clientX: 0,
          clientY: 0,
        })
      );
    });

    await expect
      .poll(async () => {
        const rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("B1"));
        return rect?.x ?? null;
      })
      .toBeGreaterThan(b1Before.x);
  });
});
