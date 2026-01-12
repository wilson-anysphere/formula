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

    await expect
      .poll(async () => Number((await page.locator("#grid-secondary").getAttribute("data-scroll-y")) ?? 0))
      .toBeCloseTo(persistedScrollY, 1);
    await expect
      .poll(async () => Number((await page.locator("#grid-secondary").getAttribute("data-zoom")) ?? 1))
      .toBeCloseTo(persistedZoom, 2);
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

    // Scroll the primary pane so A1 is offscreen.
    const primary = page.locator("#grid");
    await primary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 200 * 24);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().y)).toBeGreaterThan(0);

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);

    // Click B2 in the secondary pane (account for headers: row header ~48px, col header ~24px).
    await secondary.click({ position: { x: 48 + 100 + 12, y: 24 + 24 + 12 } });

    await expect(page.getByTestId("active-cell")).toHaveText("B2");
    await expect(page.getByTestId("formula-address")).toHaveValue("B2");
    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(Math.abs(scrollAfter - scrollBefore)).toBeLessThan(0.1);
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

    // Scroll the secondary pane so A1 is offscreen.
    await secondary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 200 * 24);
    await expect.poll(async () => Number((await secondary.getAttribute("data-scroll-y")) ?? 0)).toBeGreaterThan(0);

    const secondaryScrollBefore = Number((await secondary.getAttribute("data-scroll-y")) ?? 0);

    const rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("B2"));
    if (!rect) throw new Error("Missing B2 rect");
    await page.locator("#grid").click({ position: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } });

    await expect(page.getByTestId("active-cell")).toHaveText("B2");
    await expect(page.getByTestId("formula-address")).toHaveValue("B2");
    const secondaryScrollAfter = Number((await secondary.getAttribute("data-scroll-y")) ?? 0);
    expect(Math.abs(secondaryScrollAfter - secondaryScrollBefore)).toBeLessThan(0.1);
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

    // Scroll the primary pane so we're verifying drag selection from secondary does not disturb it.
    const primary = page.locator("#grid");
    await primary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 200 * 24);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().y)).toBeGreaterThan(0);

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);

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

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(Math.abs(scrollAfter - scrollBefore)).toBeLessThan(0.1);
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

    // Scroll the primary pane so the selected cells are offscreen (sync should not scroll it back).
    const primary = page.locator("#grid");
    await primary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 200 * 24);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().y)).toBeGreaterThan(0);

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);

    // Create a multi-range selection from the secondary pane: A1, then Ctrl/Cmd+click C3.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    const modifier: "Control" | "Meta" = process.platform === "darwin" ? "Meta" : "Control";
    await secondary.click({ position: { x: 48 + 2 * 100 + 12, y: 24 + 2 * 24 + 12 }, modifiers: [modifier] }); // C3

    await expect(page.getByTestId("selection-range")).toHaveText("2 ranges");
    await expect(page.getByTestId("active-cell")).toHaveText("C3");
    await expect(page.getByTestId("formula-address")).toHaveValue("C3");

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(Math.abs(scrollAfter - scrollBefore)).toBeLessThan(0.1);
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

    // Scroll the secondary pane so A1 is offscreen; syncing selection from primary should not scroll it back.
    await secondary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 200 * 24);
    await expect.poll(async () => Number((await secondary.getAttribute("data-scroll-y")) ?? 0)).toBeGreaterThan(0);

    const secondaryScrollBefore = Number((await secondary.getAttribute("data-scroll-y")) ?? 0);

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

    const secondaryScrollAfter = Number((await secondary.getAttribute("data-scroll-y")) ?? 0);
    expect(Math.abs(secondaryScrollAfter - secondaryScrollBefore)).toBeLessThan(0.1);
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

    const secondaryBox = await secondaryGrid.boundingBox();
    if (!secondaryBox) throw new Error("Missing secondary grid bounding box");

    const before = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("B1"));
    if (!before) throw new Error("Missing B1 rect");

    // Drag the boundary between columns A and B in the *secondary* header row to make column A wider.
    const boundaryX = before.x;
    const boundaryY = before.y / 2;

    await page.mouse.move(secondaryBox.x + boundaryX, secondaryBox.y + boundaryY);
    await page.mouse.down();
    await page.mouse.move(secondaryBox.x + boundaryX + 80, secondaryBox.y + boundaryY, { steps: 4 });
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

  test("clipboard + delete shortcuts work while focus is in the secondary pane", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page, "/?grid=shared");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Enable split view.
    await page.getByTestId("split-vertical").click();
    const secondary = page.getByTestId("grid-secondary");
    await expect(secondary).toBeVisible();

    // Select A1 in the secondary pane and enter a value.
    await secondary.click({ position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("F2");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await editor.fill("Secondary");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    // Re-focus the secondary grid and copy the cell value.
    await secondary.click({ position: { x: 60, y: 40 } });
    await page.keyboard.press(`${modifier}+C`);
    await waitForIdle(page);

    // Paste into B1 using the secondary pane focus.
    await secondary.click({ position: { x: 160, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    await page.keyboard.press(`${modifier}+V`);
    await waitForIdle(page);

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1"))).toBe("Secondary");

    // Delete clears the selection.
    await secondary.click({ position: { x: 160, y: 40 } });
    await page.keyboard.press("Delete");
    await waitForIdle(page);
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("B1"))).toBe("");
  });
});
