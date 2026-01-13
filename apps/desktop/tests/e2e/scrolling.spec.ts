import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  // Vite may occasionally trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-wait.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
      await page.evaluate(() => (window as any).__formulaApp.whenIdle());
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      throw err;
    }
  }
}

test.describe("grid scrolling + virtualization", () => {
  test("wheel scroll down reaches far rows and clicking selects correct cell", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

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
    await gotoDesktop(page);
    await waitForIdle(page);
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

  test("ArrowRight navigation auto-scrolls to keep the active cell visible", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);
    const grid = page.locator("#grid");

    await grid.click({ position: { x: 60, y: 40 } });

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().x);
    expect(scrollBefore).toBe(0);

    for (let i = 0; i < 50; i += 1) {
      await page.keyboard.press("ArrowRight");
    }

    await expect(page.getByTestId("active-cell")).toHaveText("AY1");
    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().x);
    expect(scrollAfter).toBeGreaterThan(scrollBefore);
  });

  test("programmatic selection sync can opt out of scrolling (split-pane)", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);
    const grid = page.locator("#grid");

    // Scroll down so A1 is offscreen.
    await grid.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 300 * 24);
    await expect.poll(async () => {
      return await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    }).toBeGreaterThan(0);

    // Select a visible cell at the current scroll offset so we can verify the active cell
    // updates even when we suppress scroll+focus.
    await grid.click({ position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).not.toHaveText("A1");

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll());

    // Simulate a split-pane selection mirror: update selection without scrolling the pane.
    await page.evaluate(() => {
      (window as any).__formulaApp.activateCell({ row: 0, col: 0 }, { scrollIntoView: false, focus: false });
    });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll());
    expect(scrollAfter.x).toBeCloseTo(scrollBefore.x, 5);
    expect(scrollAfter.y).toBeCloseTo(scrollBefore.y, 5);
  });

  test("programmatic selection sync can opt out of focusing the grid (split-pane)", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    const grid = page.locator("#grid");

    // Scroll down so A1 is offscreen (to also ensure focus suppression doesn't accidentally scroll).
    await grid.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 300 * 24);
    await expect.poll(async () => {
      return await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    }).toBeGreaterThan(0);

    // Focus an element outside the grid (name box / formula address).
    const address = page.getByTestId("formula-address");
    await address.click();
    await expect(address).toBeFocused();

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll());

    // Update selection without focusing the grid.
    await page.evaluate(() => {
      (window as any).__formulaApp.activateCell({ row: 0, col: 0 }, { scrollIntoView: false, focus: false });
    });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(address).toBeFocused();

    // Also ensure range selection updates do not steal focus.
    await page.evaluate(() => {
      (window as any).__formulaApp.selectRange(
        { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 } },
        { scrollIntoView: false, focus: false }
      );
    });

    await expect(page.getByTestId("selection-range")).toHaveText("A1:B2");
    await expect(address).toBeFocused();

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll());
    expect(scrollAfter.x).toBeCloseTo(scrollBefore.x, 5);
    expect(scrollAfter.y).toBeCloseTo(scrollBefore.y, 5);
  });

  test("programmatic selection sync can opt out of focusing the grid in shared-grid mode (split-pane)", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);

    const grid = page.locator("#grid");

    // Scroll down so A1 is offscreen (to also ensure focus suppression doesn't accidentally scroll).
    await grid.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 300 * 24);
    await expect.poll(async () => {
      return await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    }).toBeGreaterThan(0);

    // Focus an element outside the grid (name box / formula address).
    const address = page.getByTestId("formula-address");
    await address.click();
    await expect(address).toBeFocused();

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll());

    await page.evaluate(() => {
      (window as any).__formulaApp.activateCell({ row: 0, col: 0 }, { scrollIntoView: false, focus: false });
    });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(address).toBeFocused();

    await page.evaluate(() => {
      (window as any).__formulaApp.selectRange(
        { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 } },
        { scrollIntoView: false, focus: false }
      );
    });

    await expect(page.getByTestId("selection-range")).toHaveText("A1:B2");
    await expect(address).toBeFocused();

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll());
    expect(scrollAfter.x).toBeCloseTo(scrollBefore.x, 5);
    expect(scrollAfter.y).toBeCloseTo(scrollBefore.y, 5);
  });

  test("programmatic range selection sync can opt out of scrolling (split-pane)", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);
    const grid = page.locator("#grid");

    // Scroll down so A1 is offscreen.
    await grid.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 300 * 24);
    await expect.poll(async () => {
      return await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    }).toBeGreaterThan(0);

    // Select a visible cell at the current scroll offset so we can verify the selection
    // updates even when we suppress scroll+focus.
    await grid.click({ position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).not.toHaveText("A1");

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll());

    // Update selection without scrolling the pane.
    await page.evaluate(() => {
      (window as any).__formulaApp.selectRange(
        { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 } },
        { scrollIntoView: false, focus: false }
      );
    });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("selection-range")).toHaveText("A1:B2");

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll());
    expect(scrollAfter.x).toBeCloseTo(scrollBefore.x, 5);
    expect(scrollAfter.y).toBeCloseTo(scrollBefore.y, 5);
  });

  test("programmatic selection sync can opt out of scrolling in shared-grid mode (split-pane)", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);
    const grid = page.locator("#grid");

    // Scroll down so A1 is offscreen.
    await grid.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 300 * 24);
    await expect.poll(async () => {
      return await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    }).toBeGreaterThan(0);

    // Select a visible cell at the current scroll offset so we can verify the active cell
    // updates even when we suppress scroll+focus.
    await grid.click({ position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).not.toHaveText("A1");

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll());

    await page.evaluate(() => {
      (window as any).__formulaApp.activateCell({ row: 0, col: 0 }, { scrollIntoView: false, focus: false });
    });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll());
    expect(scrollAfter.x).toBeCloseTo(scrollBefore.x, 5);
    expect(scrollAfter.y).toBeCloseTo(scrollBefore.y, 5);
  });

  test("programmatic range selection sync can opt out of scrolling in shared-grid mode (split-pane)", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);
    const grid = page.locator("#grid");

    // Scroll down so A1 is offscreen.
    await grid.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 300 * 24);
    await expect.poll(async () => {
      return await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    }).toBeGreaterThan(0);

    // Select a visible cell at the current scroll offset so we can verify the selection
    // updates even when we suppress scroll+focus.
    await grid.click({ position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).not.toHaveText("A1");

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll());

    await page.evaluate(() => {
      (window as any).__formulaApp.selectRange(
        { range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 } },
        { scrollIntoView: false, focus: false }
      );
    });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("selection-range")).toHaveText("A1:B2");

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll());
    expect(scrollAfter.x).toBeCloseTo(scrollBefore.x, 5);
    expect(scrollAfter.y).toBeCloseTo(scrollBefore.y, 5);
  });

  test("name box Go To scrolls and updates selection", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    const address = page.getByTestId("formula-address");
    await address.click();
    await address.fill("A500");
    await address.press("Enter");

    await expect(page.getByTestId("active-cell")).toHaveText("A500");
    const scroll = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scroll).toBeGreaterThan(0);
  });

  test("name box Go To range scrolls and selects the full range", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

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

  test("name box Go To resolves sheet display names and ignores stale names after rename", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");

      // Ensure the sheet exists in the metadata store with a user-facing display name.
      // `id` is the stable sheet id used by the DocumentController.
      const store = (app as any).getWorkbookSheetStore?.();
      if (!store) throw new Error("Missing workbook sheet store");
      store.addAfter("Sheet1", { id: "sheet-1", name: "Budget" });

      // Materialize the sheet in the DocumentController.
      app.getDocument().setCellValue("sheet-1", "A1", "BudgetCell");

      // Start from Sheet1 so sheet-qualified Go To must resolve by display name.
      app.activateSheet("Sheet1");
      app.activateCell({ sheetId: "Sheet1", row: 0, col: 0 });
    });
    await waitForIdle(page);

    const address = page.getByTestId("formula-address");
    await address.click();
    await address.fill("Budget!A1");
    await address.press("Enter");
    await waitForIdle(page);

    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    expect(await page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("sheet-1");

    // Rename the display name and ensure the new name resolves.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const store = (app as any).getWorkbookSheetStore?.();
      if (!store) throw new Error("Missing workbook sheet store");
      store.rename("sheet-1", "Budget2026");

      // Switch away so we can ensure stale-name Go To doesn't switch sheets.
      app.activateSheet("Sheet1");
    });
    await waitForIdle(page);

    await address.click();
    await address.fill("Budget2026!A1");
    await address.press("Enter");
    await waitForIdle(page);
    expect(await page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("sheet-1");

    // Stale display name should not create a phantom sheet or change selection.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      app.activateSheet("Sheet1");
      app.activateCell({ sheetId: "Sheet1", row: 0, col: 0 });
    });
    await waitForIdle(page);

    await address.click();
    await address.fill("Budget!A1");
    await address.press("Enter");
    await waitForIdle(page);

    const state = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      return { activeSheetId: app.getCurrentSheetId(), sheetIds: app.getDocument().getSheetIds() };
    });
    expect(state.activeSheetId).toBe("Sheet1");
    expect(state.sheetIds).not.toContain("Budget");
  });

  test("wheel scroll right reaches far columns and clicking selects correct cell", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

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
    await gotoDesktop(page);
    await waitForIdle(page);
    const grid = page.locator("#grid");

    // Focus A1.
    await grid.click({ position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scrollBefore).toBe(0);

    const track = page.getByTestId("scrollbar-track-y");
    await expect(track).toBeVisible();
    const box = await track.boundingBox();
    expect(box).toBeTruthy();
    // Click near the bottom of the track (page down).
    await page.mouse.click(box!.x + box!.width / 2, box!.y + box!.height - 4);

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scrollAfter).toBeGreaterThan(scrollBefore);
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
  });

  test("horizontal scrollbar track click scrolls without changing the selection", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);
    const grid = page.locator("#grid");

    await grid.click({ position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().x);
    expect(scrollBefore).toBe(0);

    const track = page.getByTestId("scrollbar-track-x");
    await expect(track).toBeVisible();
    const box = await track.boundingBox();
    expect(box).toBeTruthy();
    await page.mouse.click(box!.x + box!.width - 4, box!.y + box!.height / 2);

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().x);
    expect(scrollAfter).toBeGreaterThan(scrollBefore);
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
  });

  test("scrollbar thumb drag scrolls without changing the selection", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);
    const grid = page.locator("#grid");

    await grid.click({ position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scrollBefore).toBe(0);

    const thumb = page.getByTestId("scrollbar-thumb-y");
    await expect(thumb).toBeVisible();
    const box = await thumb.boundingBox();
    expect(box).toBeTruthy();

    await page.mouse.move(box!.x + box!.width / 2, box!.y + box!.height / 2);
    await page.mouse.down();
    await page.mouse.move(box!.x + box!.width / 2, box!.y + box!.height / 2 + 60);
    await page.mouse.up();

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scrollAfter).toBeGreaterThan(scrollBefore);
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
  });

  test("horizontal scrollbar thumb drag scrolls without changing the selection", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);
    const grid = page.locator("#grid");

    await grid.click({ position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().x);
    expect(scrollBefore).toBe(0);

    const thumb = page.getByTestId("scrollbar-thumb-x");
    await expect(thumb).toBeVisible();
    const box = await thumb.boundingBox();
    expect(box).toBeTruthy();

    await page.mouse.move(box!.x + box!.width / 2, box!.y + box!.height / 2);
    await page.mouse.down();
    await page.mouse.move(box!.x + box!.width / 2 + 60, box!.y + box!.height / 2);
    await page.mouse.up();

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().x);
    expect(scrollAfter).toBeGreaterThan(scrollBefore);
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
  });

  test("charts remain anchored to sheet coordinates while scrolling", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    const before = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const charts = typeof app?.listCharts === "function" ? app.listCharts() : [];
      const chart = Array.isArray(charts) ? charts[0] : null;
      if (!chart) return null;

      const rect = (app as any).chartAnchorToViewportRect?.(chart.anchor) ?? null;
      if (!rect) return null;
      return {
        scroll: app.getScroll(),
        left: rect.left,
        top: rect.top,
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
      const charts = typeof app?.listCharts === "function" ? app.listCharts() : [];
      const chart = Array.isArray(charts) ? charts[0] : null;
      if (!chart) return null;

      const rect = (app as any).chartAnchorToViewportRect?.(chart.anchor) ?? null;
      if (!rect) return null;
      return {
        scroll: app.getScroll(),
        left: rect.left,
        top: rect.top,
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
      const charts = typeof app?.listCharts === "function" ? app.listCharts() : [];
      const chart = Array.isArray(charts) ? charts[0] : null;
      if (!chart) return null;

      const rect = (app as any).chartAnchorToViewportRect?.(chart.anchor) ?? null;
      if (!rect) return null;
      return {
        scroll: app.getScroll(),
        left: rect.left,
        top: rect.top,
      };
    });
    expect(afterX).not.toBeNull();

    const deltaScrollX = afterX!.scroll.x - after!.scroll.x;
    expect(deltaScrollX).toBeGreaterThan(0);
    expect(Math.abs((after!.left - afterX!.left) - deltaScrollX)).toBeLessThan(1);
    expect(Math.abs(after!.top - afterX!.top)).toBeLessThan(1);
  });
});
