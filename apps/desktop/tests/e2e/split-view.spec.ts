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
      await page.waitForFunction(() => Boolean((window.__formulaApp as any)?.whenIdle), null, { timeout: 10_000 });
      await page.evaluate(() => (window.__formulaApp as any).whenIdle());
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

async function dragFromTo(
  page: import("@playwright/test").Page,
  from: { x: number; y: number },
  to: { x: number; y: number },
): Promise<void> {
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y);
  await page.mouse.up();
}

async function dragInLocator(
  page: import("@playwright/test").Page,
  locator: import("@playwright/test").Locator,
  from: { x: number; y: number },
  to: { x: number; y: number },
): Promise<void> {
  // Use locator-relative coordinates so the drag stays correct even if the page layout shifts
  // (e.g. formula bar resizing while entering edit / range selection modes).
  await locator.hover({ position: from });
  await page.mouse.down();
  await locator.hover({ position: to });
  await page.mouse.up();
}

function rectCenter(rect: { x: number; y: number; width: number; height: number }): { x: number; y: number } {
  return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
}

function applyOffset(point: { x: number; y: number }, offset: { x: number; y: number }): { x: number; y: number } {
  return { x: point.x + offset.x, y: point.y + offset.y };
}

test.describe("split view", () => {
  const LAYOUT_KEY = "formula.layout.workbook.local-workbook.v1";

  test("secondary pane mounts a real grid with independent scroll + zoom and persists state", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");

    // Start from a clean persisted layout so the test is deterministic.
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(4);
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
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await expect(page.locator("#grid-secondary")).toBeVisible();
    await expect(page.locator("#grid-secondary canvas")).toHaveCount(4);
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    await expect
      .poll(async () => Number((await page.locator("#grid-secondary").getAttribute("data-scroll-y")) ?? 0))
      .toBeCloseTo(persistedScrollY, 1);
    await expect
      .poll(async () => Number((await page.locator("#grid-secondary").getAttribute("data-zoom")) ?? 1))
      .toBeCloseTo(persistedZoom, 2);
  });

  test("horizontal split view mounts secondary pane with independent scroll and persists direction", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");

    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-horizontal").click();
    await expect(page.locator("#grid-split")).toHaveAttribute("data-split-direction", "horizontal");

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(4);
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
          if (!raw) return "none";
          const layout = JSON.parse(raw);
          return layout?.splitView?.direction ?? "none";
        }, LAYOUT_KEY);
      })
      .toBe("horizontal");

    // Reload and ensure horizontal split restores.
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await expect(page.locator("#grid-secondary")).toBeVisible();
    await expect(page.locator("#grid-split")).toHaveAttribute("data-split-direction", "horizontal");
  });

  test("splitter drag updates split ratio (clamped) and restores it on reload", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");

    // Start from a clean persisted layout so the test is deterministic.
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-horizontal").click();
    await expect(page.locator("#grid-split")).toHaveAttribute("data-split-direction", "horizontal");

    const gridSplit = page.locator("#grid-split");
    const gridSplitter = page.locator("#grid-splitter");
    await expect(gridSplitter).toBeVisible();

    // Splitter drag should *not* emit layout changes on every pointermove (that would trigger
    // `renderLayout()` and cause jank). Expect a single layout change per completed drag.
    await page.evaluate(() => {
      const controller = window.__layoutController as any;
      (window as any).__splitRatioChangeCount = 0;
      if (!controller || typeof controller.on !== "function") return;
      controller.on("change", () => {
        (window as any).__splitRatioChangeCount += 1;
      });
    });

    const getChangeCount = async () => {
      return await page.evaluate(() => (window as any).__splitRatioChangeCount ?? 0);
    };

    const getInMemoryRatio = async () => {
      return await page.evaluate(() => (window.__layoutController as any)?.layout?.splitView?.ratio ?? 0);
    };

    const getPersistedRatio = async () => {
      return await page.evaluate((key) => {
        const raw = localStorage.getItem(key);
        if (!raw) return 0;
        const layout = JSON.parse(raw);
        return layout?.splitView?.ratio ?? 0;
      }, LAYOUT_KEY);
    };

    // Ensure the split view layout is persisted before we start asserting on localStorage.
    await expect.poll(getPersistedRatio).toBeGreaterThanOrEqual(0.1);

    const splitBox = await gridSplit.boundingBox();
    if (!splitBox) throw new Error("Missing #grid-split bounding box");

    // Drag the splitter near the very top of the split region; ratio should clamp to 0.1.
    let splitterBox = await gridSplitter.boundingBox();
    if (!splitterBox) throw new Error("Missing #grid-splitter bounding box");

    const splitterCenter = { x: splitterBox.x + splitterBox.width / 2, y: splitterBox.y + splitterBox.height / 2 };

    // Clicking the splitter without dragging should not change the ratio (avoid accidental drift).
    const ratioBeforeClick = await getInMemoryRatio();
    const persistedBeforeClick = await getPersistedRatio();
    const changeCountBeforeClick = await getChangeCount();
    await page.mouse.move(splitterCenter.x, splitterCenter.y);
    await page.mouse.down();
    await page.mouse.up();
    expect(await getInMemoryRatio()).toBeCloseTo(ratioBeforeClick, 6);
    expect(await getPersistedRatio()).toBeCloseTo(persistedBeforeClick, 6);
    expect(await getChangeCount()).toBe(changeCountBeforeClick);

    await dragFromTo(page, splitterCenter, { x: splitterCenter.x, y: splitBox.y + Math.max(1, splitBox.height * 0.01) });

    await expect.poll(getInMemoryRatio).toBeCloseTo(0.1, 3);
    await expect.poll(getPersistedRatio).toBeCloseTo(0.1, 3);
    await expect.poll(getChangeCount).toBe(1);

    // Drag to a mid-range ratio and ensure it is reflected both in-memory and in persisted layout.
    const targetRatio = 0.73;
    splitterBox = await gridSplitter.boundingBox();
    if (!splitterBox) throw new Error("Missing #grid-splitter bounding box after clamp drag");

    const splitterCenter2 = { x: splitterBox.x + splitterBox.width / 2, y: splitterBox.y + splitterBox.height / 2 };
    await dragFromTo(page, splitterCenter2, { x: splitterCenter2.x, y: splitBox.y + splitBox.height * targetRatio });

    await expect.poll(getInMemoryRatio).toBeCloseTo(targetRatio, 1);
    await expect.poll(getPersistedRatio).toBeCloseTo(targetRatio, 1);
    await expect.poll(getChangeCount).toBe(2);

    const inMemoryRatio = await getInMemoryRatio();
    const persistedRatio = await getPersistedRatio();

    // Sanity: the ratio should always be clamped within [0.1, 0.9].
    expect(inMemoryRatio).toBeGreaterThanOrEqual(0.1);
    expect(inMemoryRatio).toBeLessThanOrEqual(0.9);
    expect(persistedRatio).toBeGreaterThanOrEqual(0.1);
    expect(persistedRatio).toBeLessThanOrEqual(0.9);

    // Reload and ensure the ratio restores.
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await expect(page.locator("#grid-split")).toHaveAttribute("data-split-direction", "horizontal");
    await expect.poll(getPersistedRatio).toBeCloseTo(persistedRatio, 3);
    await expect.poll(getInMemoryRatio).toBeCloseTo(inMemoryRatio, 3);
  });

  test("secondary pane supports clipboard shortcuts + Delete key", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page, "/?grid=shared");

    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(4);

    // Focus/select A1 in secondary pane.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Type a value and commit (Enter moves selection down).
    await page.keyboard.press("h");
    const editor = secondary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await expect(page.getByTestId("status-mode")).toHaveText("Edit");
    await page.keyboard.type("ello");
    await page.keyboard.press("Enter");
    await waitForIdle(page);
    await expect(page.getByTestId("status-mode")).toHaveText("Ready");

    // Ensure the edit committed.
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("hello");

    // Ensure focus + selection are on the secondary pane before exercising clipboard shortcuts.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await secondary.focus();

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Copy A1 and paste into B1, all while focus remains in secondary pane.
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

    // Backspace should also clear the active cell (common on mac keyboards / laptop layouts).
    await page.keyboard.press(`${modifier}+V`);
    await waitForIdle(page);
    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"))).toBe("hello");

    await page.keyboard.press("Backspace");
    await waitForIdle(page);
    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C1"))).toBe("");
  });

  test("Shift+F2 opens the comments panel and focuses the new comment input from the secondary pane", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Focus/select A1 in secondary pane.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await secondary.focus();
    await expect.poll(() => page.evaluate(() => document.activeElement?.id)).toBe("grid-secondary");

    await page.keyboard.press("Shift+F2");

    const panel = page.getByTestId("comments-panel");
    await expect(panel).toBeVisible();

    const input = page.getByTestId("new-comment-input");
    await expect(input).toBeVisible();
    await expect(input).toBeFocused();
  });

  test("Shift+F2 does not open comments while secondary in-cell editing is active", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Focus/select A1 in secondary pane and begin editing.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("h");
    const editor = secondary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await expect(editor).toBeFocused();

    const panel = page.getByTestId("comments-panel");
    await expect(panel).not.toBeVisible();

    // Comments shortcuts should not steal focus / interrupt editing.
    await page.keyboard.press("Shift+F2");
    await expect(panel).not.toBeVisible();
    await expect(editor).toBeVisible();
    await expect(editor).toBeFocused();
  });

  test("Delete/Backspace do not clear sheet contents while secondary in-cell editing is active", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Focus/select A1 in secondary pane.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const valueBefore = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));

    // Enter edit mode and ensure the editor has focus.
    await page.keyboard.press("F2");
    const editor = secondary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await expect(editor).toBeFocused();

    // While editing, Delete/Backspace should apply to the editor text and should not clear the underlying cell.
    await page.keyboard.press("Delete");
    await expect(editor).toBeVisible();
    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe(valueBefore);

    await page.keyboard.press("Backspace");
    await expect(editor).toBeVisible();
    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe(valueBefore);

    // Cancel out of editing so later tests aren't affected.
    await page.keyboard.press("Escape");
    await waitForIdle(page);
    await expect(editor).toBeHidden();
    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe(valueBefore);
  });

  test("clipboard/comments/clear commands do not dispatch while secondary in-cell editing is active", async ({ page }) => {
    await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Focus/select A1 in secondary pane and open the in-cell editor.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await page.keyboard.press("F2");

    const editor = secondary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await expect(editor).toBeFocused();

    // Record commands executed while editing. (We should not dispatch any spreadsheet commands
    // from keybindings while a text editor is focused.)
    await page.evaluate(() => {
      const registry = (window as any).__formulaCommandRegistry;
      if (!registry) throw new Error("Missing window.__formulaCommandRegistry");

      (window as any).__executedCommandIds = [];
      (window as any).__disposeExecutedCommandListener = registry.onDidExecuteCommand((evt: any) => {
        (window as any).__executedCommandIds.push(evt.commandId);
      });
    });

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    await page.keyboard.press(`${modifier}+C`);
    await page.keyboard.press(`${modifier}+X`);
    await page.keyboard.press(`${modifier}+V`);
    await page.keyboard.press("Delete");
    await page.keyboard.press("Backspace");
    await page.keyboard.press("Shift+F2");

    // Give any async dispatch a chance to run if it was (incorrectly) triggered.
    await page.waitForTimeout(50);

    const executed = await page.evaluate(() => (window as any).__executedCommandIds ?? []);
    expect(executed).toEqual([]);

    await page.evaluate(() => {
      try {
        (window as any).__disposeExecutedCommandListener?.();
      } catch {
        // ignore
      }
      (window as any).__disposeExecutedCommandListener = null;
    });

    // Clean up: cancel editing so later tests aren't affected.
    await page.keyboard.press("Escape");
    await waitForIdle(page);
    await expect(editor).toBeHidden();
  });

  test("primary in-cell edits commit on blur when clicking another cell (shared grid)", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);
    await waitForGridCanvasesToBeSized(page, "#grid");

    const primary = page.locator("#grid");

    // Begin editing A1 in the primary pane, but do not press Enter.
    await primary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await page.keyboard.press("h");
    const editor = primary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("ello");
    await expect(editor).toHaveValue("hello");

    // Clicking another cell within the primary pane should commit and move selection.
    await primary.click({ position: { x: 48 + 100 + 12, y: 24 + 12 } }); // B1
    await waitForIdle(page);

    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("hello");
    await expect(editor).toBeHidden();
  });

  test("primary in-cell edits commit on blur without stealing focus when clicking the formula bar (shared grid)", async ({
    page,
  }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);
    await waitForGridCanvasesToBeSized(page, "#grid");

    const primary = page.locator("#grid");

    // Begin editing A1 in the primary pane, but do not press Enter.
    await primary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await page.keyboard.press("h");
    const editor = primary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("ello");
    await expect(editor).toHaveValue("hello");

    // Clicking the formula bar should commit the edit but keep the formula bar focused
    // (no focus ping-pong back to the grid).
    await page.getByTestId("formula-highlight").click();
    const input = page.getByTestId("formula-input");
    await expect(input).toBeVisible();
    await expect(input).toBeFocused();
    await waitForIdle(page);

    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("hello");
    await expect(editor).toBeHidden();
    await expect(input).toBeFocused();
  });

  test("primary in-cell edits commit on blur without stealing focus when clicking the secondary pane (shared grid)", async ({
    page,
  }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const primary = page.locator("#grid");
    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Begin editing A1 in the primary pane, but do not press Enter.
    await primary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await page.keyboard.press("h");
    const editor = primary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("ello");
    await expect(editor).toHaveValue("hello");

    // Clicking the secondary pane should commit the edit without stealing focus back to the primary grid.
    await secondary.click({ position: { x: 48 + 100 + 12, y: 24 + 12 } }); // B1
    await waitForIdle(page);

    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("hello");
    await expect(editor).toBeHidden();
    await expect.poll(() => page.evaluate(() => document.activeElement?.id)).toBe("grid-secondary");
  });

  test("secondary in-cell edits commit on blur when clicking the primary pane", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Begin editing A1 in the secondary pane, but do not press Enter.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await page.keyboard.press("h");
    const editor = secondary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("ello");
    await expect(editor).toHaveValue("hello");

    // Clicking the primary pane should commit the edit (Excel semantics).
    await page.locator("#grid").click({ position: { x: 48 + 100 + 12, y: 24 + 12 } }); // B1
    await waitForIdle(page);

    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("hello");
    await expect(editor).toBeHidden();
  });

  test("secondary in-cell edits commit when clicking another cell in the secondary pane", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Begin editing A1 in the secondary pane, but do not press Enter.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await page.keyboard.press("h");
    const editor = secondary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("ello");
    await expect(editor).toHaveValue("hello");

    // Clicking another cell within the secondary pane should commit and move selection.
    await secondary.click({ position: { x: 48 + 100 + 12, y: 24 + 12 } }); // B1
    await waitForIdle(page);

    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    await expect(page.getByTestId("formula-address")).toHaveValue("B1");
    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("hello");
    await expect(editor).toBeHidden();
    await expect.poll(() => page.evaluate(() => document.activeElement?.id)).toBe("grid-secondary");
  });

  test("selection is global across panes without cross-pane scrolling", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");

    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(4);
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
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(4);
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
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(4);
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
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(4);
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
    // Name box reflects the selection (Excel semantics), not the active/anchor cell.
    await expect(page.getByTestId("formula-address")).toHaveValue("B2:D4");
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
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(4);
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
    await expect(page.getByTestId("formula-address")).toHaveValue("2 ranges");
    await expect(page.locator("#grid").getByTestId("canvas-grid-a11y-active-cell")).toContainText("Cell C3");

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll());
    expect(Math.abs(scrollAfter.x - scrollBefore.x)).toBeLessThan(0.1);
    expect(Math.abs(scrollAfter.y - scrollBefore.y)).toBeLessThan(0.1);
  });

  test("primary multi-range selection sync does not scroll the secondary pane", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(4);
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
    await expect(page.getByTestId("formula-address")).toHaveValue("2 ranges");
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
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(4);
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
    await page.reload({ waitUntil: "domcontentloaded" });
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

      await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
      await expect(page.getByTestId("grid-secondary")).toBeVisible();
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
      await page.waitForFunction(() => Boolean((window as any).__formulaSecondaryGrid?.getCellRect), undefined, { timeout: 10_000 });
      await expect
        .poll(() => page.evaluate(() => (window as any).__formulaSecondaryGrid?.interactionMode ?? null))
        .toBe("rangeSelection");
      const [a1Rect, a2Rect] = await page.evaluate(() => {
        const grid = (window as any).__formulaSecondaryGrid;
        return [grid.getCellRect(1, 1), grid.getCellRect(2, 1)];
      });
      if (!a1Rect || !a2Rect) throw new Error("Missing secondary pane cell rects for A1/A2");
      await dragInLocator(page, secondary, rectCenter(a1Rect), rectCenter(a2Rect));

      await expect(input).toHaveValue("=SUM(A1:A2");
      await expect(secondaryStatus).toContainText("Selection A1:A2");
      await expect(input).toBeFocused();
      await expect
        .poll(() => page.evaluate(() => (window as any).__formulaSecondaryGrid?.renderer?.referenceHighlights?.length ?? 0))
        .toBeGreaterThan(0);

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
      await expect
        .poll(() => page.evaluate(() => (window as any).__formulaSecondaryGrid?.renderer?.referenceHighlights?.length ?? 0))
        .toBe(0);

      // Start editing again.
      await page.getByTestId("formula-highlight").click();
      await expect(input).toBeVisible();
      await input.fill("=SUM(");
      await expect
        .poll(() => page.evaluate(() => (window as any).__formulaSecondaryGrid?.interactionMode ?? null))
        .toBe("rangeSelection");

      // Drag-select again; focus should return to the formula bar so typing continues.
      await dragInLocator(page, secondary, rectCenter(a1Rect), rectCenter(a2Rect));

      await expect(input).toHaveValue("=SUM(A1:A2");
      await expect(input).toBeFocused();
      await expect
        .poll(() => page.evaluate(() => (window as any).__formulaSecondaryGrid?.renderer?.referenceHighlights?.length ?? 0))
        .toBeGreaterThan(0);

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

      // Formula-bar commit (Enter) advances the active cell like in-cell editing.
      await expect(secondaryStatus).toContainText("Selection C2");
      await expect
        .poll(() => page.evaluate(() => (window as any).__formulaSecondaryGrid?.renderer?.referenceHighlights?.length ?? 0))
        .toBe(0);
    });

    test(`secondary-pane range insertion on another sheet inserts a sheet-qualified reference (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Lazily create Sheet2 and seed A1:A2.
      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const doc = app.getDocument();
        doc.setCellValue("Sheet2", "A1", 7);
        doc.setCellValue("Sheet2", "A2", 8);
      });
      await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

      await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
      await expect(page.getByTestId("grid-secondary")).toBeVisible();
      await waitForGridCanvasesToBeSized(page, "#grid-secondary");

      // Start editing on Sheet1!C1.
      await page.click("#grid", { position: { x: 260, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("C1");

      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=SUM(");

      // Switch to Sheet2 while still editing and pick A1:A2 from the secondary pane.
      await page.getByTestId("sheet-tab-Sheet2").click();
      await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");

      const secondary = page.locator("#grid-secondary");
      await expect(secondary).toBeVisible();
      await page.waitForFunction(() => Boolean((window as any).__formulaSecondaryGrid?.getCellRect), undefined, { timeout: 10_000 });
      await expect
        .poll(() => page.evaluate(() => (window as any).__formulaSecondaryGrid?.interactionMode ?? null))
        .toBe("rangeSelection");
      const [a1Rect, a2Rect] = await page.evaluate(() => {
        const grid = (window as any).__formulaSecondaryGrid;
        return [grid.getCellRect(1, 1), grid.getCellRect(2, 1)];
      });
      if (!a1Rect || !a2Rect) throw new Error("Missing secondary pane cell rects for A1/A2");

      await dragInLocator(page, secondary, rectCenter(a1Rect), rectCenter(a2Rect));

      await expect(input).toHaveValue("=SUM(Sheet2!A1:A2");
    });

    test(`dragging a range in the primary pane updates secondary reference highlights (${mode})`, async ({ page }) => {
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

      // Split-view controls are owned by the ribbon UI. Scope to the ribbon root so the locator
      // stays stable even if other surfaces (e.g. legacy debug toolbars) add similarly named test ids.
      await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
      await expect(page.getByTestId("grid-secondary")).toBeVisible();
      await waitForGridCanvasesToBeSized(page, "#grid-secondary");
      await waitForGridCanvasesToBeSized(page, "#grid");

      // Select C1 in the primary pane.
      await page.click("#grid", { position: { x: 260, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("C1");

      // Start editing in the formula bar.
      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=SUM(");

      if (mode === "shared") {
        // Shared-grid range selection mode is driven by formula-bar overlays. Wait for the
        // interaction mode flip before attempting a drag-select, otherwise the gesture can be
        // interpreted as a normal selection change and no range reference is inserted.
        await expect
          .poll(() => page.evaluate(() => (window as any).__formulaApp?.sharedGrid?.interactionMode ?? null))
          .toBe("rangeSelection");
      }

      // Drag select A1:A2 in the *primary* pane.
      const primary = page.locator("#grid");
      const gridBox = await primary.boundingBox();
      if (!gridBox) throw new Error("Missing #grid bounding box");
      const { a1Rect, a2Rect } = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        return { a1Rect: app.getCellRectA1("A1"), a2Rect: app.getCellRectA1("A2") };
      });
      if (!a1Rect || !a2Rect) throw new Error("Missing primary pane cell rects for A1/A2");

      // `getCellRectA1` is expected to be viewport-relative, but some renderers return grid-relative
      // coordinates. Detect and normalize to locator-relative points so the drag remains stable even
      // if the layout shifts (e.g. formula bar resizing during edit/range selection).
      const rectsAreGridRelative = a1Rect.y < gridBox.y - 1;
      const normalizePoint = (point: { x: number; y: number }) =>
        rectsAreGridRelative ? point : { x: point.x - gridBox.x, y: point.y - gridBox.y };

      await dragInLocator(page, primary, normalizePoint(rectCenter(a1Rect)), normalizePoint(rectCenter(a2Rect)));

      await expect(input).toHaveValue("=SUM(A1:A2");
      await expect
        .poll(() => page.evaluate(() => (window as any).__formulaSecondaryGrid?.renderer?.referenceHighlights?.length ?? 0))
        .toBeGreaterThan(0);
    });

    test(`secondary-pane in-place edits apply to the active sheet after switching sheets (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      // Ensure split-view layout is deterministic (no persisted sheet/scroll/zoom surprises).
      await page.evaluate(() => localStorage.clear());
      await page.reload({ waitUntil: "domcontentloaded" });
      await waitForDesktopReady(page);
      await waitForIdle(page);

      // Seed distinct values so we can detect writes going to the wrong sheet.
      await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const doc = app.getDocument();
        doc.setCellValue("Sheet1", "A1", "sheet1");
        doc.setCellValue("Sheet2", "A1", "sheet2");
      });
      await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
      await waitForIdle(page);

      await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
      const secondary = page.locator("#grid-secondary");
      await expect(secondary).toBeVisible();
      await waitForGridCanvasesToBeSized(page, "#grid-secondary");

      // Switch to Sheet2 while split view is active.
      await page.getByTestId("sheet-tab-Sheet2").click();
      await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
      await waitForIdle(page);

      // Edit A1 in the secondary pane; it should mutate Sheet2, not Sheet1.
      await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
      await page.keyboard.press("e");
      const editor = secondary.locator("textarea.cell-editor");
      await expect(editor).toBeVisible();
      await page.keyboard.type("dited");
      await page.keyboard.press("Enter");
      await waitForIdle(page);

      const { sheet1, sheet2 } = await page.evaluate(() => {
        const app = (window as any).__formulaApp;
        const doc = app.getDocument();
        const s1 = doc.getCell("Sheet1", "A1");
        const s2 = doc.getCell("Sheet2", "A1");
        return { sheet1: s1?.value ?? null, sheet2: s2?.value ?? null };
      });

      expect(sheet1).toBe("sheet1");
      expect(sheet2).toBe("edited");
    });
  }

  test("secondary pane supports in-place editing without scrolling the primary pane", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    // Enable split view.
    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(4);

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

  test("secondary keyboard navigation updates global selection without scrolling primary", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Scroll the primary pane away so cross-pane scroll bugs are detectable.
    const primary = page.locator("#grid");
    await primary.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(100 * 100, 0);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().x)).toBeGreaterThan(0);
    await page.mouse.wheel(0, 200 * 24);
    await expect.poll(async () => await page.evaluate(() => (window as any).__formulaApp.getScroll().y)).toBeGreaterThan(0);

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll());

    // Focus secondary on A1 then use arrow keys to move selection.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("ArrowDown");
    await expect(page.getByTestId("active-cell")).toHaveText("A2");
    await expect(page.getByTestId("formula-address")).toHaveValue("A2");
    await expect(page.locator("#grid").getByTestId("canvas-grid-a11y-active-cell")).toContainText("Cell A2");

    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll());
    expect(Math.abs(scrollAfter.x - scrollBefore.x)).toBeLessThan(0.1);
    expect(Math.abs(scrollAfter.y - scrollBefore.y)).toBeLessThan(0.1);
  });

  test("F2 starts editing in the secondary pane", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    // Seed a deterministic initial value so the F2 edit semantics (caret-at-end)
    // are asserted without coupling to SpreadsheetApp's default seeded data.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();
      doc.setCellValue(sheetId, "A1", "seed");
    });
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("seed");

    await page.keyboard.press("F2");
    const editor = secondary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await expect(page.getByTestId("status-mode")).toHaveText("Edit");
    // F2 enters "edit existing value" mode (Excel semantics) rather than clearing the cell.
    // Ensure typing appends at the cursor position rather than replacing via a new edit session.
    await expect(editor).toHaveValue("seed");
    await page.keyboard.type("F2");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await expect(page.getByTestId("status-mode")).toHaveText("Ready");
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("seedF2");
  });

  test("clicking inside the secondary cell editor does not commit the edit", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    // Seed a known value so we can assert edits are not committed prematurely.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();
      doc.setCellValue(sheetId, "A1", "seed");
    });
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Begin editing and type a value, but do NOT press Enter/Tab.
    await page.keyboard.press("h");
    const editor = secondary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("ello");
    await expect(editor).toHaveValue("hello");

    // Clicking inside the textarea should keep it focused and not commit/close.
    await editor.click({ position: { x: 10, y: 10 } });
    await expect(editor).toBeFocused();
    await expect(editor).toHaveValue("hello");

    expect(await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("seed");

    // Commit explicitly.
    await page.keyboard.press("Enter");
    await waitForIdle(page);
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("hello");
  });

  test("clicking the primary pane commits an in-progress secondary-pane edit", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Start editing C2 in the secondary pane but do NOT press Enter/Tab.
    await secondary.click({ position: { x: 48 + 2 * 100 + 12, y: 24 + 1 * 24 + 12 } });
    await page.keyboard.press("h");
    const editor = secondary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("ello");

    // Click A1 in the primary pane; this should blur/commit the secondary editor.
    const rectA1 = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));
    if (!rectA1) throw new Error("Missing A1 rect");
    await page.locator("#grid").click({ position: { x: rectA1.x + rectA1.width / 2, y: rectA1.y + rectA1.height / 2 } });

    await expect(editor).not.toBeVisible();
    await waitForIdle(page);
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C2"))).toBe("hello");
  });

  test("split view cannot be disabled while secondary-pane editing is active (commit then disable)", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();

    // Start editing C2 in the secondary pane but do NOT press Enter/Tab.
    await secondary.click({ position: { x: 48 + 2 * 100 + 12, y: 24 + 1 * 24 + 12 } });
    await page.keyboard.press("h");
    const editor = secondary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("ello");

    const splitNone = page.getByTestId("ribbon-root").getByTestId("split-none");
    await expect(splitNone).toBeDisabled();

    // Commit the edit explicitly, then disable split view.
    await page.keyboard.press("Enter");
    await expect(editor).not.toBeVisible();
    await waitForIdle(page);
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C2"))).toBe("hello");

    await expect(splitNone).toBeEnabled();
    await splitNone.click();
    await expect(secondary).not.toBeVisible();
    await waitForIdle(page);
  });

  test("menu-save commits an in-progress secondary-pane edit", async ({ page }) => {
    // File -> Save should commit pending edits (including the split-view secondary editor).
    // In the Playwright harness we trigger the Tauri menu event directly.
    // Stub the minimal `__TAURI__` surface so the handler is registered in the browser-based
    // Playwright harness.
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;
      (window as any).__TAURI__ = {
        core: {
          invoke: async () => null,
        },
        event: {
          listen: async (name: string, handler: any) => {
            listeners[name] = handler;
            return () => {
              delete listeners[name];
            };
          },
          emit: async () => {},
        },
      };
    });

    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();

    // Start editing C2 in the secondary pane but do NOT press Enter/Tab.
    await secondary.click({ position: { x: 48 + 2 * 100 + 12, y: 24 + 1 * 24 + 12 } });
    await page.keyboard.press("h");
    const editor = secondary.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await page.keyboard.type("ello");

    // Trigger Save (which calls `commitAllPendingEditsForCommand()` in `main.ts`) while the
    // secondary editor is still active.
    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["menu-save"]), undefined, { timeout: 10_000 });
    await page.evaluate(() => {
      (window as any).__tauriListeners["menu-save"]({ payload: null });
    });
    await waitForIdle(page);

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("C2"))).toBe("hello");
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
    await gotoDesktop(page, "/?grid=shared");

    const secondaryGrid = page.getByTestId("grid-secondary");

    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("B1");
      return rect && typeof rect.x === "number" && rect.width > 0 && rect.height > 0;
    });

    // Enable split view.
    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
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

  test("fill handle drag works in the secondary pane", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    await page.waitForFunction(() => Boolean((window as any).__formulaSecondaryGrid), undefined, { timeout: 10_000 });

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      // Seed a simple numeric series so fill-mode "series" produces predictable results.
      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellValue(sheetId, "A2", 2);

      const grid = (window as any).__formulaSecondaryGrid;
      // Select A1:A2 (grid coordinates include headers at row/col 0).
      grid.setSelectionRanges(
        [
          {
            startRow: 1,
            endRow: 3,
            startCol: 1,
            endCol: 2,
          },
        ],
        { activeCell: { row: 2, col: 1 } },
      );
    });
    await waitForIdle(page);

    const secondaryBox = await page.locator("#grid-secondary").boundingBox();
    expect(secondaryBox).not.toBeNull();

    await page.waitForFunction(() => Boolean((window as any).__formulaSecondaryGrid?.renderer?.getFillHandleRect?.()));

    const handle = await page.evaluate(() => (window as any).__formulaSecondaryGrid.renderer.getFillHandleRect());
    expect(handle).not.toBeNull();

    const a3Rect = await page.evaluate(() => (window as any).__formulaSecondaryGrid.getCellRect(3, 1));
    expect(a3Rect).not.toBeNull();

    await dragFromTo(
      page,
      {
        x: secondaryBox!.x + handle!.x + handle!.width / 2,
        y: secondaryBox!.y + handle!.y + handle!.height / 2,
      },
      {
        x: secondaryBox!.x + a3Rect!.x + a3Rect!.width / 2,
        y: secondaryBox!.y + a3Rect!.y + a3Rect!.height / 2,
      },
    );
    await waitForIdle(page);

    const a3 = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A3"));
    expect(a3).toBe("3");
  });

  test("Escape cancels an in-progress fill handle drag in the secondary pane", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    await page.waitForFunction(() => Boolean((window as any).__formulaSecondaryGrid), undefined, { timeout: 10_000 });

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      const doc = app.getDocument();

      doc.setCellValue(sheetId, "A1", 1);
      doc.setCellValue(sheetId, "A2", 2);
      // Clear seeded values in A3/A4 so we can assert cancellation leaves them untouched.
      doc.setCellValue(sheetId, "A3", "");
      doc.setCellValue(sheetId, "A4", "");

      const grid = (window as any).__formulaSecondaryGrid;
      grid.setSelectionRanges(
        [
          {
            startRow: 1,
            endRow: 3,
            startCol: 1,
            endCol: 2,
          },
        ],
        { activeCell: { row: 2, col: 1 } },
      );
    });
    await waitForIdle(page);

    const [a3Before, a4Before] = await Promise.all([
      page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A3")),
      page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A4")),
    ]);

    const secondaryBox = await page.locator("#grid-secondary").boundingBox();
    expect(secondaryBox).not.toBeNull();

    await page.waitForFunction(() => Boolean((window as any).__formulaSecondaryGrid?.renderer?.getFillHandleRect?.()));

    const handle = await page.evaluate(() => (window as any).__formulaSecondaryGrid.renderer.getFillHandleRect());
    expect(handle).not.toBeNull();

    // Drag towards A4, then press Escape before releasing.
    const a4Rect = await page.evaluate(() => (window as any).__formulaSecondaryGrid.getCellRect(4, 1));
    expect(a4Rect).not.toBeNull();

    await page.mouse.move(secondaryBox!.x + handle!.x + handle!.width / 2, secondaryBox!.y + handle!.y + handle!.height / 2);
    await page.mouse.down();
    await page.mouse.move(secondaryBox!.x + a4Rect!.x + a4Rect!.width / 2, secondaryBox!.y + a4Rect!.y + a4Rect!.height / 2);

    await expect.poll(() => page.evaluate(() => document.activeElement?.id)).toBe("grid-secondary");
    await expect.poll(() => page.evaluate(() => (window as any).__formulaSecondaryGrid?.dragMode ?? null)).toBe("fillHandle");

    await page.keyboard.press("Escape");
    await expect.poll(() => page.evaluate(() => (window as any).__formulaSecondaryGrid?.dragMode ?? null)).toBe(null);

    await page.mouse.up();
    await waitForIdle(page);

    const [a3, a4] = await Promise.all([
      page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A3")),
      page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A4")),
    ]);
    expect(a3).toBe(a3Before);
    expect(a4).toBe(a4Before);
  });

  test("secondary pane opens the custom grid context menu on right click", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Focus/select A1, then right-click B2 to ensure the secondary pane handles the context menu event.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } }); // A1
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(
      ({ x, y }) => {
        const grid = document.getElementById("grid-secondary");
        if (!grid) throw new Error("Missing #grid-secondary");
        const rect = grid.getBoundingClientRect();
        grid.dispatchEvent(
          new MouseEvent("contextmenu", {
            bubbles: true,
            cancelable: true,
            button: 2,
            clientX: rect.left + x,
            clientY: rect.top + y,
          }),
        );
      },
      { x: 48 + 100 + 12, y: 24 + 24 + 12 },
    ); // B2
    await expect(page.getByTestId("active-cell")).toHaveText("B2");

    const menuOverlay = page.getByTestId("context-menu");
    await expect(menuOverlay).toBeVisible();

    await page.keyboard.press("Escape");
    await expect(menuOverlay).toBeHidden();
  });

  test("Shift+F10 opens the grid context menu anchored to the active split pane", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();

    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    // Focus/select A1 in the secondary pane so the split view marks it active.
    await secondary.click({ position: { x: 48 + 12, y: 24 + 12 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect
      .poll(() => page.evaluate(() => (window.__layoutController as any)?.layout?.splitView?.activePane ?? null))
      .toBe("secondary");

    await page.keyboard.press("Shift+F10");

    const menuOverlay = page.getByTestId("context-menu");
    const menu = menuOverlay.locator(".context-menu");
    await expect(menuOverlay).toBeVisible();

    const [secondaryBox, menuBox] = await Promise.all([secondary.boundingBox(), menu.boundingBox()]);
    if (!secondaryBox) throw new Error("Missing secondary grid bounding box");
    if (!menuBox) throw new Error("Missing context menu bounding box");

    // If the context menu were incorrectly anchored to the primary grid, it would open to the left
    // of the secondary pane in vertical split mode.
    expect(menuBox.x).toBeGreaterThan(secondaryBox.x - 1);

    await page.keyboard.press("Escape");
    await expect(menuOverlay).toBeHidden();
  });
}); 
