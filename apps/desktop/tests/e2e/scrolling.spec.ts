import { expect, test } from "@playwright/test";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("grid scrolling + virtualization", () => {
  test("wheel scroll down reaches far rows and clicking selects correct cell", async ({ page }) => {
    await page.goto("/");

    // Seed A200 (0-based row 199) with a sentinel string.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.getDocument().setCellValue(sheetId, { row: 199, col: 0 }, "Bottom");
      app.refresh();
    });
    await waitForIdle(page);

    // Wheel-scroll so that row 200 is near the top of the viewport.
    // (cellHeight is currently fixed at 24px in SpreadsheetApp.)
    const grid = page.locator("#grid");
    await grid.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 199 * 24);

    await expect.poll(async () => {
      return await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    }).toBeGreaterThan(0);

    // Click within A200.
    await grid.click({ position: { x: 60, y: 24 + 12 } });

    await expect(page.getByTestId("active-cell")).toHaveText("A200");
    await expect(page.getByTestId("active-value")).toHaveText("Bottom");
  });

  test("ArrowDown navigation auto-scrolls to keep the active cell visible", async ({ page }) => {
    await page.goto("/");
    const grid = page.locator("#grid");

    // Focus A1 (account for headers).
    await grid.click({ position: { x: 60, y: 40 } });

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scrollBefore).toBe(0);

    for (let i = 0; i < 200; i += 1) {
      await page.keyboard.press("ArrowDown");
    }

    await expect(page.getByTestId("active-cell")).toHaveText("A201");
    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scrollAfter).toBeGreaterThan(scrollBefore);
  });

  test("name box Go To scrolls and updates selection", async ({ page }) => {
    await page.goto("/");

    const address = page.getByTestId("formula-address");
    await address.click();
    await address.fill("A500");
    await address.press("Enter");

    await expect(page.getByTestId("active-cell")).toHaveText("A500");
    const scroll = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scroll).toBeGreaterThan(0);
  });

  test("name box Go To range scrolls and selects the full range", async ({ page }) => {
    await page.goto("/");

    const address = page.getByTestId("formula-address");
    await address.click();
    await address.fill("A500:C505");
    await address.press("Enter");

    await expect(page.getByTestId("active-cell")).toHaveText("A500");
    await expect(page.getByTestId("selection-range")).toHaveText("A500:C505");

    const scroll = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scroll).toBeGreaterThan(0);

    const drawn = await page.evaluate(() => (window as any).__formulaApp.getLastSelectionDrawn());
    expect(drawn).toBeTruthy();
    expect(drawn.ranges.length).toBeGreaterThan(0);
    expect(drawn.ranges[0].rect.width).toBeGreaterThan(0);
    expect(drawn.ranges[0].rect.height).toBeGreaterThan(0);
  });

  test("wheel scroll right reaches far columns and clicking selects correct cell", async ({ page }) => {
    await page.goto("/");

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.getDocument().setCellValue(sheetId, { row: 0, col: 100 }, "FarX");
      app.refresh();
    });
    await waitForIdle(page);

    const grid = page.locator("#grid");
    await grid.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(100 * 100, 0);

    await expect.poll(async () => {
      return await page.evaluate(() => (window as any).__formulaApp.getScroll().x);
    }).toBeGreaterThan(0);

    await grid.click({ position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("CW1");
    await expect(page.getByTestId("active-value")).toHaveText("FarX");
  });

  test("scrollbar track click scrolls without changing the selection", async ({ page }) => {
    await page.goto("/");
    const grid = page.locator("#grid");

    // Focus A1.
    await grid.click({ position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scrollBefore).toBe(0);

    const track = page.getByTestId("scrollbar-track-y");
    const box = await track.boundingBox();
    expect(box).toBeTruthy();
    // Click near the bottom of the track (page down).
    await page.mouse.click(box!.x + box!.width / 2, box!.y + box!.height - 4);

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scrollAfter).toBeGreaterThan(scrollBefore);
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
  });

  test("horizontal scrollbar track click scrolls without changing the selection", async ({ page }) => {
    await page.goto("/");
    const grid = page.locator("#grid");

    await grid.click({ position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().x);
    expect(scrollBefore).toBe(0);

    const track = page.getByTestId("scrollbar-track-x");
    const box = await track.boundingBox();
    expect(box).toBeTruthy();
    await page.mouse.click(box!.x + box!.width - 4, box!.y + box!.height / 2);

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().x);
    expect(scrollAfter).toBeGreaterThan(scrollBefore);
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
  });

  test("charts remain anchored to sheet coordinates while scrolling", async ({ page }) => {
    await page.goto("/");

    const before = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const el = document.querySelector<HTMLElement>('[data-testid="chart-object"]');
      if (!el) return null;
      return {
        scroll: app.getScroll(),
        left: Number.parseFloat(el.style.left),
        top: Number.parseFloat(el.style.top),
      };
    });
    expect(before).not.toBeNull();
    expect(Number.isFinite(before!.left)).toBe(true);
    expect(Number.isFinite(before!.top)).toBe(true);

    const grid = page.locator("#grid");
    await grid.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 240);

    await expect.poll(async () => {
      return await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    }).toBeGreaterThan(before!.scroll.y);

    const after = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const el = document.querySelector<HTMLElement>('[data-testid="chart-object"]');
      if (!el) return null;
      return {
        scroll: app.getScroll(),
        left: Number.parseFloat(el.style.left),
        top: Number.parseFloat(el.style.top),
      };
    });
    expect(after).not.toBeNull();

    const deltaScroll = after!.scroll.y - before!.scroll.y;
    // Chart DOM nodes are positioned in sheet space minus scroll offsets.
    expect(Math.abs((before!.top - after!.top) - deltaScroll)).toBeLessThan(1);
    expect(Math.abs(before!.left - after!.left)).toBeLessThan(1);

    // Horizontal scroll should move the chart's x position by the scroll delta.
    await page.mouse.wheel(240, 0);

    const afterX = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const el = document.querySelector<HTMLElement>('[data-testid="chart-object"]');
      if (!el) return null;
      return {
        scroll: app.getScroll(),
        left: Number.parseFloat(el.style.left),
        top: Number.parseFloat(el.style.top),
      };
    });
    expect(afterX).not.toBeNull();

    const deltaScrollX = afterX!.scroll.x - after!.scroll.x;
    expect(deltaScrollX).toBeGreaterThan(0);
    expect(Math.abs((after!.left - afterX!.left) - deltaScrollX)).toBeLessThan(1);
    expect(Math.abs(after!.top - afterX!.top)).toBeLessThan(1);
  });
});
