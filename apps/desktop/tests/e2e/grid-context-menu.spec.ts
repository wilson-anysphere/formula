import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  // Vite may occasionally trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-wait.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => Boolean((window.__formulaApp as any)?.whenIdle), null, { timeout: 60_000 });
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

async function waitForGridCanvasesToBeSized(page: import("@playwright/test").Page, rootSelector: string): Promise<void> {
  // Canvas sizing happens asynchronously (ResizeObserver + rAF). Ensure the renderer
  // has produced non-zero backing buffers before attempting hit-testing.
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

async function openGridContextMenuAt(
  page: import("@playwright/test").Page,
  selector: string,
  position: { x: number; y: number },
): Promise<void> {
  // Avoid flaky right-click handling in the desktop shell by dispatching a deterministic contextmenu event.
  await page.evaluate(
    ({ selector, x, y }) => {
      const root = document.querySelector(selector) as HTMLElement | null;
      if (!root) throw new Error(`Missing ${selector}`);
      const rect = root.getBoundingClientRect();
      root.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          // Match the right-click button used by browsers.
          button: 2,
          clientX: rect.left + x,
          clientY: rect.top + y,
        }),
      );
    },
    { selector, x: position.x, y: position.y },
  );
}

async function selectRange(
  page: import("@playwright/test").Page,
  range: { startRow: number; endRow: number; startCol: number; endCol: number },
): Promise<void> {
  await page.evaluate((r) => {
    (window.__formulaApp as any).selectRange({ range: r });
  }, range);
}

async function getActiveCell(page: import("@playwright/test").Page): Promise<{ row: number; col: number }> {
  return await page.evaluate(() => (window.__formulaApp as any).getActiveCell());
}

async function getActiveCellRect(
  page: import("@playwright/test").Page,
): Promise<{ x: number; y: number; width: number; height: number }> {
  const rect = await page.evaluate(() => (window.__formulaApp as any).getActiveCellRect());
  if (!rect) throw new Error("Active cell rect was null");
  return rect;
}

test.describe("Grid context menus", () => {
  test("right-clicking a row header opens a menu with Row Height…", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await expect(page.locator("#grid")).toBeVisible();
    await openGridContextMenuAt(page, "#grid", { x: 10, y: 40 });

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Row Height…" })).toBeVisible();
  });

  test("right-clicking a column header opens a menu with Column Width…", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await expect(page.locator("#grid")).toBeVisible();
    await openGridContextMenuAt(page, "#grid", { x: 100, y: 10 });

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Column Width…" })).toBeVisible();
  });

  test("right-clicking a cell shows a Paste Special… submenu with expected modes", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await expect(page.locator("#grid")).toBeVisible();
    // Click away from the row/column headers so we get the normal cell context menu.
    await openGridContextMenuAt(page, "#grid", { x: 80, y: 40 });

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const pasteSpecial = menu.getByRole("button", { name: "Paste Special…" });
    await expect(pasteSpecial).toBeVisible();

    // Mouse hover should open the submenu (Excel-like).
    await pasteSpecial.hover();
    const submenu = menu.locator(".context-menu__submenu");
    await expect(submenu).toBeVisible();
    await expect(submenu.getByRole("button", { name: /^Paste$/ })).toBeVisible();
    await expect(submenu.getByRole("button", { name: "Paste Values" })).toBeVisible();
    await expect(submenu.getByRole("button", { name: "Paste Formulas" })).toBeVisible();
    await expect(submenu.getByRole("button", { name: "Paste Formats" })).toBeVisible();

    // Keyboard navigation should open the submenu (ArrowRight).
    await pasteSpecial.focus();
    await page.keyboard.press("ArrowRight");
    await expect(submenu.getByRole("button", { name: /^Paste$/ })).toBeFocused();

    // ArrowLeft should close the submenu and restore focus to the parent item.
    await page.keyboard.press("ArrowLeft");
    await expect(submenu).toBeHidden();
    await expect(pasteSpecial).toBeFocused();
  });

  test("legacy mode: Hide/Unhide row + column via header context menu affects keyboard navigation", async ({ page }) => {
    await gotoDesktop(page, "/?grid=legacy");
    await waitForIdle(page);

    const grid = page.locator("#grid");
    await expect(grid).toBeVisible();
    const limits = await page.evaluate(() => (window.__formulaApp as any).getGridLimits());
    const gridBox = await grid.boundingBox();
    if (!gridBox) throw new Error("Missing #grid bounding box");

    // Hide row 2 (0-based row=1) using the row-header context menu.
    await selectRange(page, { startRow: 1, endRow: 1, startCol: 0, endCol: 0 }); // A2
    const a3Before = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("A3"));
    if (!a3Before) throw new Error("Missing A3 rect");
    const a2 = await getActiveCellRect(page);
    await openGridContextMenuAt(page, "#grid", {
      x: a2.x - gridBox.x - 5,
      y: a2.y - gridBox.y + a2.height / 2,
    });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Hide", exact: true }).click();
    await expect(menu).toBeHidden();

    // Hiding the active row should move selection to the next visible row (Excel-like).
    await expect.poll(() => getActiveCell(page)).toEqual({ row: 2, col: 0 });
    // And the row should visually collapse (A3 shifts up).
    await expect
      .poll(async () => {
        const rect = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("A3"));
        return rect?.y ?? Number.POSITIVE_INFINITY;
      })
      .toBeLessThan((a3Before as any).y - 10);

    // Verify ArrowDown skips the hidden row (A1 -> A3).
    await selectRange(page, { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }); // A1
    await page.evaluate(() => (window.__formulaApp as any).focus());
    await page.keyboard.press("ArrowDown");
    await expect.poll(() => getActiveCell(page)).toEqual({ row: 2, col: 0 });

    // Unhide the row by selecting a full-row band that spans it and using the row-header context menu.
    // (Right-clicking a row header will otherwise replace non-band selections with a single-row band.)
    await selectRange(page, { startRow: 0, endRow: 2, startCol: 0, endCol: limits.maxCols - 1 });
    const a1ForRowMenu = await getActiveCellRect(page);
    await openGridContextMenuAt(page, "#grid", {
      x: a1ForRowMenu.x - gridBox.x - 5,
      y: a1ForRowMenu.y - gridBox.y + a1ForRowMenu.height / 2,
    });
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Unhide", exact: true }).click();
    await expect(menu).toBeHidden();

    // Verify ArrowDown no longer skips (A1 -> A2).
    await selectRange(page, { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }); // A1
    await page.evaluate(() => (window.__formulaApp as any).focus());
    await page.keyboard.press("ArrowDown");
    await expect.poll(() => getActiveCell(page)).toEqual({ row: 1, col: 0 });

    // Hide column B (0-based col=1) using the column-header context menu.
    await selectRange(page, { startRow: 0, endRow: 0, startCol: 1, endCol: 1 }); // B1
    const c1Before = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("C1"));
    if (!c1Before) throw new Error("Missing C1 rect");
    const b1 = await getActiveCellRect(page);
    await openGridContextMenuAt(page, "#grid", {
      x: b1.x - gridBox.x + b1.width / 2,
      y: b1.y - gridBox.y - 5,
    });
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Hide", exact: true }).click();
    await expect(menu).toBeHidden();

    // Hiding the active column should move selection to the next visible column (Excel-like).
    await expect.poll(() => getActiveCell(page)).toEqual({ row: 0, col: 2 });
    // And the column should visually collapse (C shifts left).
    await expect
      .poll(async () => {
        const rect = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("C1"));
        return rect?.x ?? Number.POSITIVE_INFINITY;
      })
      .toBeLessThan((c1Before as any).x - 10);

    // Verify ArrowRight skips the hidden column (A1 -> C1).
    await selectRange(page, { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }); // A1
    await page.evaluate(() => (window.__formulaApp as any).focus());
    await page.keyboard.press("ArrowRight");
    await expect.poll(() => getActiveCell(page)).toEqual({ row: 0, col: 2 });

    // Unhide the column by selecting a full-column band that spans it and using the column-header context menu.
    // (Right-clicking a column header will otherwise replace non-band selections with a single-column band.)
    await selectRange(page, { startRow: 0, endRow: limits.maxRows - 1, startCol: 0, endCol: 2 });
    const a1ForColMenu = await getActiveCellRect(page);
    await openGridContextMenuAt(page, "#grid", {
      x: a1ForColMenu.x - gridBox.x + a1ForColMenu.width / 2,
      y: a1ForColMenu.y - gridBox.y - 5,
    });
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Unhide", exact: true }).click();
    await expect(menu).toBeHidden();

    // Verify ArrowRight no longer skips (A1 -> B1).
    await selectRange(page, { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }); // A1
    await page.evaluate(() => (window.__formulaApp as any).focus());
    await page.keyboard.press("ArrowRight");
    await expect.poll(() => getActiveCell(page)).toEqual({ row: 0, col: 1 });
  });

  test("shared mode: Hide/Unhide row + column via header context menu affects keyboard navigation", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);

    const grid = page.locator("#grid");
    await expect(grid).toBeVisible();
    const limits = await page.evaluate(() => (window.__formulaApp as any).getGridLimits());
    const gridBox = await grid.boundingBox();
    if (!gridBox) throw new Error("Missing #grid bounding box");

    // Hide row 2 (0-based row=1) using the row-header context menu.
    await selectRange(page, { startRow: 1, endRow: 1, startCol: 0, endCol: 0 }); // A2
    const a3Before = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("A3"));
    if (!a3Before) throw new Error("Missing A3 rect");
    const a2 = await getActiveCellRect(page);
    await openGridContextMenuAt(page, "#grid", {
      x: 10,
      y: a2.y - gridBox.y + a2.height / 2,
    });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Hide", exact: true }).click();
    await expect(menu).toBeHidden();

    // Hiding the active row should move selection to the next visible row (Excel-like).
    await expect.poll(() => getActiveCell(page)).toEqual({ row: 2, col: 0 });
    // And the row should visually collapse (A3 shifts up).
    await expect
      .poll(async () => {
        const rect = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("A3"));
        return rect?.y ?? Number.POSITIVE_INFINITY;
      })
      .toBeLessThan((a3Before as any).y - 10);

    // Verify ArrowDown skips the hidden row (A1 -> A3).
    await selectRange(page, { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }); // A1
    await page.evaluate(() => (window.__formulaApp as any).focus());
    await page.keyboard.press("ArrowDown");
    await expect.poll(() => getActiveCell(page)).toEqual({ row: 2, col: 0 });

    // Unhide the row by selecting a full-row band that spans it and using the row-header context menu.
    // (Right-clicking a row header will otherwise replace non-band selections with a single-row band.)
    await selectRange(page, { startRow: 0, endRow: 2, startCol: 0, endCol: limits.maxCols - 1 });
    const a1ForRowMenu = await getActiveCellRect(page);
    await openGridContextMenuAt(page, "#grid", {
      x: 10,
      y: a1ForRowMenu.y - gridBox.y + a1ForRowMenu.height / 2,
    });
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Unhide", exact: true }).click();
    await expect(menu).toBeHidden();

    // Verify ArrowDown no longer skips (A1 -> A2).
    await selectRange(page, { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }); // A1
    await page.evaluate(() => (window.__formulaApp as any).focus());
    await page.keyboard.press("ArrowDown");
    await expect.poll(() => getActiveCell(page)).toEqual({ row: 1, col: 0 });

    // Hide column B (0-based col=1) using the column-header context menu.
    await selectRange(page, { startRow: 0, endRow: 0, startCol: 1, endCol: 1 }); // B1
    const c1Before = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("C1"));
    if (!c1Before) throw new Error("Missing C1 rect");
    const b1 = await getActiveCellRect(page);
    await openGridContextMenuAt(page, "#grid", {
      x: b1.x - gridBox.x + b1.width / 2,
      y: 10,
    });
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Hide", exact: true }).click();
    await expect(menu).toBeHidden();

    // Hiding the active column should move selection to the next visible column (Excel-like).
    await expect.poll(() => getActiveCell(page)).toEqual({ row: 0, col: 2 });
    // And the column should visually collapse (C shifts left).
    await expect
      .poll(async () => {
        const rect = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("C1"));
        return rect?.x ?? Number.POSITIVE_INFINITY;
      })
      .toBeLessThan((c1Before as any).x - 10);

    // Verify ArrowRight skips the hidden column (A1 -> C1).
    await selectRange(page, { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }); // A1
    await page.evaluate(() => (window.__formulaApp as any).focus());
    await page.keyboard.press("ArrowRight");
    await expect.poll(() => getActiveCell(page)).toEqual({ row: 0, col: 2 });

    // Unhide the column by selecting a full-column band that spans it and using the column-header context menu.
    // (Right-clicking a column header will otherwise replace non-band selections with a single-column band.)
    await selectRange(page, { startRow: 0, endRow: limits.maxRows - 1, startCol: 0, endCol: 2 });
    const a1ForColMenu = await getActiveCellRect(page);
    await openGridContextMenuAt(page, "#grid", {
      x: a1ForColMenu.x - gridBox.x + a1ForColMenu.width / 2,
      y: 10,
    });
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Unhide", exact: true }).click();
    await expect(menu).toBeHidden();

    // Verify ArrowRight no longer skips (A1 -> B1).
    await selectRange(page, { startRow: 0, endRow: 0, startCol: 0, endCol: 0 }); // A1
    await page.evaluate(() => (window.__formulaApp as any).focus());
    await page.keyboard.press("ArrowRight");
    await expect.poll(() => getActiveCell(page)).toEqual({ row: 0, col: 1 });
  });

  test("Row Height… rejects Excel-scale selections in shared-grid mode", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);

    await expect(page.locator("#grid")).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid");

    // Use the corner context menu "Select All" to create an Excel-scale selection.
    await openGridContextMenuAt(page, "#grid", { x: 10, y: 10 });
    const cornerMenu = page.getByTestId("context-menu");
    await expect(cornerMenu).toBeVisible();
    await cornerMenu.getByRole("button", { name: "Select All" }).click();
    await expect(cornerMenu).toBeHidden();

    // Attempt to apply a row height to the full-sheet selection.
    await openGridContextMenuAt(page, "#grid", { x: 10, y: 40 });
    const rowMenu = page.getByTestId("context-menu");
    await expect(rowMenu).toBeVisible();
    await rowMenu.getByRole("button", { name: "Row Height…" }).click();

    const toast = page
      .getByTestId("toast")
      .filter({ hasText: "Selection too large to resize rows. Select fewer rows and try again." });
    await expect(toast).toBeVisible();
    await expect(toast).toHaveAttribute("data-type", "warning");

    await expect(page.getByTestId("input-box")).toHaveCount(0);
  });

  test("Column Width… rejects Excel-scale selections in shared-grid mode", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);

    await expect(page.locator("#grid")).toBeVisible();
    await waitForGridCanvasesToBeSized(page, "#grid");

    // Use the corner context menu "Select All" to create an Excel-scale selection.
    await openGridContextMenuAt(page, "#grid", { x: 10, y: 10 });
    const cornerMenu = page.getByTestId("context-menu");
    await expect(cornerMenu).toBeVisible();
    await cornerMenu.getByRole("button", { name: "Select All" }).click();
    await expect(cornerMenu).toBeHidden();

    // Attempt to apply a column width to the full-sheet selection.
    await openGridContextMenuAt(page, "#grid", { x: 100, y: 10 });
    const colMenu = page.getByTestId("context-menu");
    await expect(colMenu).toBeVisible();
    await colMenu.getByRole("button", { name: "Column Width…" }).click();

    const toast = page
      .getByTestId("toast")
      .filter({ hasText: "Selection too large to resize columns. Select fewer columns and try again." });
    await expect(toast).toBeVisible();
    await expect(toast).toHaveAttribute("data-type", "warning");

    await expect(page.getByTestId("input-box")).toHaveCount(0);
  });

  test("right-clicking a row header in split-view secondary pane opens a menu with Row Height…", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(4);
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    await openGridContextMenuAt(page, "#grid-secondary", { x: 10, y: 40 });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Row Height…" })).toBeVisible();
  });

  test("right-clicking a column header in split-view secondary pane opens a menu with Column Width…", async ({ page }) => {
    await gotoDesktop(page, "/?grid=shared");
    await waitForIdle(page);

    await page.getByTestId("ribbon-root").getByTestId("split-vertical").click();
    const secondary = page.locator("#grid-secondary");
    await expect(secondary).toBeVisible();
    await expect(secondary.locator("canvas")).toHaveCount(4);
    await waitForGridCanvasesToBeSized(page, "#grid-secondary");

    await openGridContextMenuAt(page, "#grid-secondary", { x: 100, y: 10 });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Column Width…" })).toBeVisible();
  });
});
