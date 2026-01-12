import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

async function waitForGridCanvasesToBeSized(
  page: import("@playwright/test").Page,
  rootSelector: string,
): Promise<void> {
  // Canvas sizing happens asynchronously (ResizeObserver + rAF). Ensure the renderer
  // has produced non-zero backing buffers before attempting wheel/drag interactions.
  await page.waitForFunction(
    (selector) => {
      const root = document.querySelector(selector);
      if (!root) return false;
      const canvases = root.querySelectorAll("canvas");
      if (canvases.length === 0) return false;
      return Array.from(canvases).every((c) => (c as HTMLCanvasElement).width > 0 && (c as HTMLCanvasElement).height > 0);
    },
    rootSelector,
    { timeout: 10_000 },
  );
}

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
        await page.waitForLoadState("load");
        continue;
      }
      throw err;
    }
  }
}

test.describe("split view", () => {
  const LAYOUT_KEY = "formula.layout.workbook.local-workbook.v1";

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
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    const primaryScrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    const secondaryScrollBefore = Number((await secondary.getAttribute("data-scroll-y")) ?? 0);

    await secondary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 600);

    await expect
      .poll(async () => Number((await secondary.getAttribute("data-scroll-y")) ?? 0))
      .toBeGreaterThan(secondaryScrollBefore);

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
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    await expect
      .poll(async () => Number((await page.locator("#grid-secondary").getAttribute("data-scroll-y")) ?? 0))
      .toBeCloseTo(persistedScrollY, 1);
    await expect
      .poll(async () => Number((await page.locator("#grid-secondary").getAttribute("data-zoom")) ?? 1))
      .toBeCloseTo(persistedZoom, 2);
  });

  test("secondary pane supports clipboard shortcuts + Delete key", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page, "/?grid=shared");

    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(3);

    // Focus/select A1 in secondary pane.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Type a value and commit (Enter moves selection down).
    await page.keyboard.press("h");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("ello");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await expect(page.getByTestId("active-cell")).toHaveText("A2");
    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("hello");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Copy A1 and paste into B1, all while focus remains in secondary pane.
    await page.keyboard.press("ArrowUp");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect.poll(() => page.evaluate(() => document.activeElement?.id)).toBe("grid-secondary");

    await page.keyboard.press(`${modifier}+C`);
    await waitForIdle(page);
    await expect.poll(() => page.evaluate(() => document.activeElement?.id)).toBe("grid-secondary");

    await page.keyboard.press("ArrowRight");
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    await page.keyboard.press(`${modifier}+V`);
    await waitForIdle(page);
    await expect.poll(() => page.evaluate(() => document.activeElement?.id)).toBe("grid-secondary");

    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1"))).toBe("hello");

    // Cut B1 and paste into C1.
    await page.keyboard.press(`${modifier}+X`);
    await waitForIdle(page);
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1"))).toBe("");
    await expect.poll(() => page.evaluate(() => document.activeElement?.id)).toBe("grid-secondary");

    await page.keyboard.press("ArrowRight");
    await expect(page.getByTestId("active-cell")).toHaveText("C1");
    await page.keyboard.press(`${modifier}+V`);
    await waitForIdle(page);
    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"))).toBe("hello");

    // Delete clears the active cell in secondary pane.
    await page.keyboard.press("Delete");
    await waitForIdle(page);
    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"))).toBe("");
  });

  test("selection is global across panes without cross-pane scrolling", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");

    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(3);
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Scroll the primary pane so A1 is offscreen.
    const primary = page.locator("#grid");
    await primary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(100 * 100, 0);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().x)).toBeGreaterThan(0);
    await page.mouse.wheel(0, 200 * 24);
    await expect
      .poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().y))
      .toBeGreaterThan(0);

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll());

    // Click B2 in the secondary pane (account for headers: row header ~48px, col header ~24px).
    await secondary.click({ position: { x: 48 + 100 + 12, y: 24 + 24 + 12 } });

    await expect(page.getByTestId("active-cell")).toHaveText("B2");
    await expect(page.getByTestId("formula-address")).toHaveValue("B2");
    // The primary pane should also mirror the selection state (even if the cell is offscreen).
    await expect(page.locator("#grid").getByTestId("canvas-grid-a11y-active-cell")).toContainText("Cell B2");
    await expect(page.locator("#grid").getByTestId("canvas-grid-a11y-status")).toContainText("Selection B2");
    await expect(secondary.getByTestId("canvas-grid-a11y-status")).toContainText("Selection B2");
    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll());
    expect(Math.abs(scrollAfter.x - scrollBefore.x)).toBeLessThan(0.1);
    expect(Math.abs(scrollAfter.y - scrollBefore.y)).toBeLessThan(0.1);
  });

  test("primary selection sync does not scroll the secondary pane", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(3);
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Scroll the secondary pane so A1 is offscreen.
    await secondary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(100 * 100, 0);
    await expect.poll(async () => Number((await secondary.getAttribute("data-scroll-x")) ?? 0)).toBeGreaterThan(0);
    await page.mouse.wheel(0, 200 * 24);
    await expect.poll(async () => Number((await secondary.getAttribute("data-scroll-y")) ?? 0)).toBeGreaterThan(0);

    const secondaryScrollBefore = {
      x: Number((await secondary.getAttribute("data-scroll-x")) ?? 0),
      y: Number((await secondary.getAttribute("data-scroll-y")) ?? 0),
    };

    const rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("B2"));
    if (!rect) throw new Error("Missing B2 rect");
    await page.locator("#grid").click({ position: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } });

    await expect(page.getByTestId("active-cell")).toHaveText("B2");
    await expect(page.getByTestId("formula-address")).toHaveValue("B2");
    // The secondary pane should mirror selection state without being scrolled to reveal it.
    await expect(secondary.getByTestId("canvas-grid-a11y-active-cell")).toContainText("Cell B2");
    await expect(secondary.getByTestId("canvas-grid-a11y-status")).toContainText("Selection B2");
    const secondaryScrollAfter = {
      x: Number((await secondary.getAttribute("data-scroll-x")) ?? 0),
      y: Number((await secondary.getAttribute("data-scroll-y")) ?? 0),
    };
    expect(Math.abs(secondaryScrollAfter.x - secondaryScrollBefore.x)).toBeLessThan(0.1);
    expect(Math.abs(secondaryScrollAfter.y - secondaryScrollBefore.y)).toBeLessThan(0.1);
  });

  test("primary keyboard navigation sync does not scroll the secondary pane", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(3);
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Scroll the secondary pane away from the top-left so selection mirroring does not auto-scroll it.
    await secondary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(100 * 100, 0);
    await page.mouse.wheel(0, 200 * 24);
    await expect.poll(async () => Number((await secondary.getAttribute("data-scroll-x")) ?? 0)).toBeGreaterThan(0);
    await expect.poll(async () => Number((await secondary.getAttribute("data-scroll-y")) ?? 0)).toBeGreaterThan(0);

    const secondaryScrollBefore = {
      x: Number((await secondary.getAttribute("data-scroll-x")) ?? 0),
      y: Number((await secondary.getAttribute("data-scroll-y")) ?? 0),
    };

    // Focus the primary pane and move selection via keyboard.
    const rectA1 = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));
    if (!rectA1) throw new Error("Missing A1 rect");
    await page.click("#grid", { position: { x: rectA1.x + rectA1.width / 2, y: rectA1.y + rectA1.height / 2 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("ArrowDown");
    await expect(page.getByTestId("active-cell")).toHaveText("A2");
    await expect(page.getByTestId("formula-address")).toHaveValue("A2");
    await expect(secondary.getByTestId("canvas-grid-a11y-active-cell")).toContainText("Cell A2");

    const secondaryScrollAfter = {
      x: Number((await secondary.getAttribute("data-scroll-x")) ?? 0),
      y: Number((await secondary.getAttribute("data-scroll-y")) ?? 0),
    };
    expect(Math.abs(secondaryScrollAfter.x - secondaryScrollBefore.x)).toBeLessThan(0.1);
    expect(Math.abs(secondaryScrollAfter.y - secondaryScrollBefore.y)).toBeLessThan(0.1);
  });

  test("secondary drag selection preserves active cell semantics and does not scroll primary", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(3);
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Scroll the primary pane so we're verifying drag selection from secondary does not disturb it.
    const primary = page.locator("#grid");
    await primary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(100 * 100, 0);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().x)).toBeGreaterThan(0);
    await page.mouse.wheel(0, 200 * 24);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().y)).toBeGreaterThan(0);

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll());

    const secondaryBox = await secondary.boundingBox();
    if (!secondaryBox) throw new Error("Missing secondary grid bounding box");

    // Drag-select from D4 -> B2 in the secondary pane.
    // Coords are derived from the default desktop grid geometry:
    // - row header width ~48px
    // - col header height ~24px
    // - default col width 100px
    // - default row height 24px
    const start = { x: 48 + 3 * 100 + 12, y: 24 + 3 * 24 + 12 }; // D4
    const end = { x: 48 + 1 * 100 + 12, y: 24 + 1 * 24 + 12 }; // B2

    await page.mouse.move(secondaryBox.x + start.x, secondaryBox.y + start.y);
    await page.mouse.down();
    await page.mouse.move(secondaryBox.x + end.x, secondaryBox.y + end.y);
    await page.mouse.up();

    // Shared-grid mouse drag keeps the *anchor* cell active (D4 here) even though the range normalizes to B2:D4.
    await expect(page.getByTestId("selection-range")).toHaveText("B2:D4");
    await expect(page.getByTestId("active-cell")).toHaveText("D4");
    await expect(page.getByTestId("formula-address")).toHaveValue("D4");
    await expect(page.locator("#grid").getByTestId("canvas-grid-a11y-active-cell")).toContainText("Cell D4");
    await expect(page.locator("#grid").getByTestId("canvas-grid-a11y-status")).toContainText("Selection B2:D4");
    await expect(secondary.getByTestId("canvas-grid-a11y-status")).toContainText("Selection B2:D4");

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll());
    expect(Math.abs(scrollAfter.x - scrollBefore.x)).toBeLessThan(0.1);
    expect(Math.abs(scrollAfter.y - scrollBefore.y)).toBeLessThan(0.1);
  });

  test("secondary multi-range selection syncs to primary without scrolling", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(3);
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Scroll the primary pane so the selected cells are offscreen (sync should not scroll it back).
    const primary = page.locator("#grid");
    await primary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(100 * 100, 0);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().x)).toBeGreaterThan(0);
    await page.mouse.wheel(0, 200 * 24);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().y)).toBeGreaterThan(0);

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll());

    // Create a multi-range selection from the secondary pane: A1, then Ctrl/Cmd+click C3.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    const modifier: "Control" | "Meta" = process.platform === "darwin" ? "Meta" : "Control";
    await secondary.click({ position: { x: 48 + 2 * 100 + 12, y: 24 + 2 * 24 + 12 }, modifiers: [modifier] }); // C3

    await expect(page.getByTestId("selection-range")).toHaveText("2 ranges");
    await expect(page.getByTestId("active-cell")).toHaveText("C3");
    await expect(page.getByTestId("formula-address")).toHaveValue("C3");
    await expect(page.locator("#grid").getByTestId("canvas-grid-a11y-active-cell")).toContainText("Cell C3");

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll());
    expect(Math.abs(scrollAfter.x - scrollBefore.x)).toBeLessThan(0.1);
    expect(Math.abs(scrollAfter.y - scrollBefore.y)).toBeLessThan(0.1);
  });

  test("primary multi-range selection sync does not scroll the secondary pane", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(3);
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Scroll the secondary pane so A1 is offscreen; syncing selection from primary should not scroll it back.
    await secondary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(100 * 100, 0);
    await expect.poll(async () => Number((await secondary.getAttribute("data-scroll-x")) ?? 0)).toBeGreaterThan(0);
    await page.mouse.wheel(0, 200 * 24);
    await expect.poll(async () => Number((await secondary.getAttribute("data-scroll-y")) ?? 0)).toBeGreaterThan(0);

    const secondaryScrollBefore = {
      x: Number((await secondary.getAttribute("data-scroll-x")) ?? 0),
      y: Number((await secondary.getAttribute("data-scroll-y")) ?? 0),
    };

    const primary = page.locator("#grid");
    const rectA1 = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));
    const rectC3 = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("C3"));
    if (!rectA1 || !rectC3) throw new Error("Missing cell rects for primary multi-range selection");

    await primary.click({ position: { x: rectA1.x + rectA1.width / 2, y: rectA1.y + rectA1.height / 2 } });
    const modifier: "Control" | "Meta" = process.platform === "darwin" ? "Meta" : "Control";
    await primary.click({
      position: { x: rectC3.x + rectC3.width / 2, y: rectC3.y + rectC3.height / 2 },
      modifiers: [modifier],
    });

    await expect(page.getByTestId("selection-range")).toHaveText("2 ranges");
    await expect(page.getByTestId("active-cell")).toHaveText("C3");
    await expect(page.getByTestId("formula-address")).toHaveValue("C3");
    await expect(secondary.getByTestId("canvas-grid-a11y-active-cell")).toContainText("Cell C3");

    const secondaryScrollAfter = {
      x: Number((await secondary.getAttribute("data-scroll-x")) ?? 0),
      y: Number((await secondary.getAttribute("data-scroll-y")) ?? 0),
    };
    expect(Math.abs(secondaryScrollAfter.x - secondaryScrollBefore.x)).toBeLessThan(0.1);
    expect(Math.abs(secondaryScrollAfter.y - secondaryScrollBefore.y)).toBeLessThan(0.1);
  });

  test("primary pane persists + restores scroll/zoom (parity with secondary)", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");

    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(3);
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Set primary zoom + scroll while split view is active. These should be stored under
    // layout.splitView.panes.primary.
    const primaryViewport = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.setZoom(1.5);
      app.setScroll(0, 400);
      return { scrollY: app.getScroll().y, zoom: app.getZoom() };
    });

    await expect
      .poll(async () => {
        return await page.evaluate((key) => {
          const raw = localStorage.getItem(key);
          if (!raw) return 0;
          const layout = JSON.parse(raw);
          return layout?.splitView?.panes?.primary?.scrollY ?? 0;
        }, LAYOUT_KEY);
      })
      .toBeCloseTo(primaryViewport.scrollY, 0);

    await expect
      .poll(async () => {
        return await page.evaluate((key) => {
          const raw = localStorage.getItem(key);
          if (!raw) return 1;
          const layout = JSON.parse(raw);
          return layout?.splitView?.panes?.primary?.zoom ?? 1;
        }, LAYOUT_KEY);
      })
      .toBeCloseTo(primaryViewport.zoom, 3);

    // Scroll the secondary pane to a different offset so we can assert both restore independently.
    const secondaryScrollBefore = Number((await secondary.getAttribute("data-scroll-y")) ?? 0);
    await secondary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 600);
    await expect
      .poll(async () => Number((await secondary.getAttribute("data-scroll-y")) ?? 0))
      .toBeGreaterThan(secondaryScrollBefore);
    const secondaryScrollY = Number((await secondary.getAttribute("data-scroll-y")) ?? 0);

    // Wait for the debounced layout persistence to flush the secondary scrollY before reloading.
    await expect
      .poll(async () => {
        return await page.evaluate((key) => {
          const raw = localStorage.getItem(key);
          if (!raw) return 0;
          const layout = JSON.parse(raw);
          return layout?.splitView?.panes?.secondary?.scrollY ?? 0;
        }, LAYOUT_KEY);
      })
      .toBeCloseTo(secondaryScrollY, 1);

    // Reload and ensure both panes restore.
    await page.reload();
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await expect(page.locator("#grid-secondary")).toBeVisible();

    await expect
      .poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().y))
      .toBeCloseTo(primaryViewport.scrollY, 0);
    await expect
      .poll(async () => await page.evaluate(() => (window as any).__formulaApp.getZoom()))
      .toBeCloseTo(primaryViewport.zoom, 3);

    await expect
      .poll(async () => Number((await page.locator("#grid-secondary").getAttribute("data-scroll-y")) ?? 0))
      .toBeCloseTo(secondaryScrollY, 1);
  });

  const modes = ["legacy", "shared"] as const;

  for (const mode of modes) {
    test(`dragging a range in the secondary pane inserts it into the formula bar (commit + cancel) (${mode})`, async ({
      page,
    }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Seed numeric inputs in A1 and A2 (so SUM has a visible result).
      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        doc.setCellValue(sheetId, "A1", 1);
        doc.setCellValue(sheetId, "A2", 2);
      });
      await waitForIdle(page);

      await page.getByTestId("split-vertical").click();
      await expect(page.getByTestId("grid-secondary")).toBeVisible();
      await waitForGridCanvasesToBeSized(page, "#grid-secondary");
      const secondary = page.locator("#grid-secondary");
      const secondaryStatus = secondary.getByTestId("canvas-grid-a11y-status");
      await waitForGridCanvasesToBeSized(page, "#grid-secondary");

      const secondary = page.locator("#grid-secondary");
      const secondaryStatus = secondary.getByTestId("canvas-grid-a11y-status");

      // Select C1 in the primary pane (same offsets as formula-bar.spec.ts).
      await page.click("#grid", { position: { x: 260, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("C1");

      // Start editing in the formula bar.
      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=SUM(");

      // Drag select A1:A2 in the secondary pane to insert a range reference.
      const gridBox = await page.locator("#grid-secondary").boundingBox();
      if (!gridBox) throw new Error("Missing grid-secondary bounding box");

      await page.mouse.move(gridBox.x + 60, gridBox.y + 40);
      await page.mouse.down();
      await page.mouse.move(gridBox.x + 60, gridBox.y + 64);
      await page.mouse.up();

      await expect(input).toHaveValue("=SUM(A1:A2");
      await expect(secondaryStatus).toContainText("Selection A1:A2");
      await expect(input).toBeFocused();

      // Cancel should clear the split-view range selection overlay and not apply the edit.
      await page.keyboard.press("Escape");
      await waitForIdle(page);

      const { c1FormulaAfterCancel } = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const doc = app.getDocument();
        const sheetId = app.getCurrentSheetId();
        return { c1FormulaAfterCancel: doc.getCell(sheetId, "C1").formula };
      });
      expect(c1FormulaAfterCancel).toBeNull();
      await expect(secondaryStatus).toContainText("Selection C1");

      // Start editing again.
      await page.getByTestId("formula-highlight").click();
      await expect(input).toBeVisible();
      await input.fill("=SUM(");

      // Drag-select again; focus should return to the formula bar so typing continues.
      await page.mouse.move(gridBox.x + 60, gridBox.y + 40);
      await page.mouse.down();
      await page.mouse.move(gridBox.x + 60, gridBox.y + 64);
      await page.mouse.up();

      await expect(input).toHaveValue("=SUM(A1:A2");
      await expect(input).toBeFocused();

      // Commit the formula; the split-view transient range selection overlay should clear.
      await page.keyboard.type(")");
      await page.keyboard.press("Enter");
      await waitForIdle(page);

      const { c1Formula } = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const doc = app.getDocument();
        const sheetId = app.getCurrentSheetId();
        return { c1Formula: doc.getCell(sheetId, "C1").formula };
      });
      expect(c1Formula).toBe("=SUM(A1:A2)");

      const c1Value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"));
      expect(c1Value).toBe("3");

      await expect(secondaryStatus).toContainText("Selection C1");
    });
  }

  test("secondary pane supports in-place editing without scrolling the primary pane", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload();
    await waitForDesktopReady(page);
    await waitForIdle(page);

    // Enable split view.
    await page.getByTestId("split-vertical").click();
    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(3);

    // Scroll the primary pane away from the origin so selection sync bugs are detectable.
    const primary = page.locator("#grid");
    await primary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(100 * 100, 0);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().x)).toBeGreaterThan(0);
    await page.mouse.wheel(0, 200 * 24);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().y)).toBeGreaterThan(0);

    const primaryScrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll());

    // Click C2 in the secondary pane (account for headers: row header ~48px, col header ~24px).
    await secondary.click({ position: { x: 48 + 2 * 100 + 12, y: 24 + 1 * 24 + 12 } });
    await expect(page.getByTestId("active-cell")).toHaveText("C2");
    await expect(page.getByTestId("formula-address")).toHaveValue("C2");

    // Start typing to begin editing (Excel semantics).
    await page.keyboard.press("h");
    const editor = secondary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("ello");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C2"))).toBe("hello");

    const primaryScrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll());
    expect(Math.abs(primaryScrollAfter.x - primaryScrollBefore.x)).toBeLessThan(0.1);
    expect(Math.abs(primaryScrollAfter.y - primaryScrollBefore.y)).toBeLessThan(0.1);
  });
});

test.describe("split view / shared grid zoom", () => {
  test("Ctrl/Cmd+wheel zoom changes grid geometry", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");

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
        }),
      );
    });

    await expect
      .poll(async () => {
        const rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("B1"));
        return rect?.x ?? null;
      })
      .toBeGreaterThan(b1Before.x);
  });

  test("secondary pane column resize updates primary pane geometry", async ({ page }) => {
    await page.goto("/?grid=shared");

    const secondaryGrid = page.getByTestId("grid-secondary");

    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("B1");
      return rect && typeof rect.x === "number" && rect.width > 0 && rect.height > 0;
    });

    // Enable split view.
    await page.getByTestId("split-vertical").click();
    await expect(secondaryGrid).toBeVisible();

    // Wait for the secondary grid canvases to mount + size.
    await page.waitForFunction(() => {
      const canvas = document.querySelector<HTMLCanvasElement>('[data-testid="grid-secondary"] canvas');
      return Boolean(canvas && canvas.width > 0 && canvas.height > 0);
    });

    const before = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("B1"));
    if (!before) throw new Error("Missing B1 rect");

    // Drag the boundary between columns A and B in the *secondary* header row to make column A wider.
    const boundaryX = before.x;
    const boundaryY = before.y / 2;

    // Use locator-relative hovers so Playwright will auto-scroll the target point into view.
    await secondaryGrid.hover({ position: { x: boundaryX, y: boundaryY } });
    await page.mouse.down();
    await secondaryGrid.hover({ position: { x: boundaryX + 80, y: boundaryY } });
    await page.mouse.up();

    await page.waitForFunction(
      (threshold) => {
        const rect = (window as any).__formulaApp.getCellRectA1("B1");
        return rect && rect.x > threshold;
      },
      before.x + 30,
    );

    const after = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("B1"));
    if (!after) throw new Error("Missing B1 rect after resize");
    expect(after.x).toBeGreaterThan(before.x + 30);
  });
});
