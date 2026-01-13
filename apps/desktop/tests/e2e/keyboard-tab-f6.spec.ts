import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function getActiveCell(page: import("@playwright/test").Page): Promise<{ row: number; col: number }> {
  return page.evaluate(() => (window.__formulaApp as any).getActiveCell());
}

async function dispatchF6(page: import("@playwright/test").Page, opts: { shiftKey?: boolean } = {}): Promise<void> {
  // Browsers can reserve F6 for built-in chrome focus cycling (address bar/toolbars),
  // which can prevent Playwright's `keyboard.press("F6")` from reaching the app.
  // Dispatching a synthetic `keydown` exercises our in-app keybinding pipeline
  // (KeybindingService -> CommandRegistry -> focus cycle) deterministically.
  await page.evaluate(({ shiftKey }) => {
    const target = (document.activeElement as HTMLElement | null) ?? document.getElementById("grid") ?? window;
    target.dispatchEvent(
      new KeyboardEvent("keydown", {
        key: "F6",
        code: "F6",
        shiftKey: Boolean(shiftKey),
        bubbles: true,
        cancelable: true,
      }),
    );
  }, opts);
}

test.describe("keyboard navigation: Tab grid traversal + F6 focus cycling", () => {
  const GRID_MODES = ["legacy", "shared"] as const;

  for (const mode of GRID_MODES) {
    test(`Tab/Shift+Tab moves selection and wraps (${mode})`, async ({ page }) => {
      const url = mode === "legacy" ? "/?grid=legacy&maxRows=3&maxCols=3" : "/?grid=shared";
      await gotoDesktop(page, url);

      const limits = await page.evaluate(() => {
        const app = window.__formulaApp as any;
        const limits = app.limits as { maxRows: number; maxCols: number };
        app.activateCell({ row: 0, col: limits.maxCols - 1 });
        app.focus();
        return limits;
      });

      await expect(page.locator("#grid")).toBeFocused();

      // End-of-row Tab wraps to the first column of the next row.
      await page.keyboard.press("Tab");
      await expect.poll(async () => await getActiveCell(page)).toEqual({ row: 1, col: 0 });
      await expect(page.locator("#grid")).toBeFocused();

      // Shift+Tab at the start of the row wraps to the last column of the previous row.
      await page.keyboard.press("Shift+Tab");
      await expect.poll(async () => await getActiveCell(page)).toEqual({ row: 0, col: limits.maxCols - 1 });
      await expect(page.locator("#grid")).toBeFocused();
    });
  }

  test("F6 / Shift+F6 cycles focus across ribbon, formula bar, sheet tabs, and grid", async ({ page }) => {
    await gotoDesktop(page, "/?grid=legacy");

    const ribbonRoot = page.getByTestId("ribbon-root");
    await expect(ribbonRoot).toBeVisible();

    await page.evaluate(() => (window.__formulaApp as any).focus());
    await expect(page.locator("#grid")).toBeFocused();

    const activeRibbonTab = ribbonRoot.locator('[role="tab"][aria-selected="true"]');
    const sheetTabsActiveTab = page.locator('#sheet-tabs button[role="tab"][aria-selected="true"]');
    const statusBarFirstFocusable = page.locator(".statusbar").getByTestId("open-version-history-panel");

    // The in-app focus cycle follows the vertical UI layout:
    // ribbon -> formula bar -> grid -> sheet tabs -> status bar -> ribbon (etc).
    // Forward cycle from the grid: grid -> sheet tabs -> status bar -> ribbon -> formula bar -> grid.
    await dispatchF6(page);
    await expect(sheetTabsActiveTab).toBeFocused();

    await dispatchF6(page);
    await expect(statusBarFirstFocusable).toBeFocused();

    await dispatchF6(page);
    await expect(activeRibbonTab).toBeFocused();

    await dispatchF6(page);
    await expect(page.getByTestId("formula-address")).toBeFocused();

    await dispatchF6(page);
    await expect(page.locator("#grid")).toBeFocused();

    // Reverse cycle from the grid: grid -> formula bar -> ribbon -> status bar -> sheet tabs -> grid.
    await dispatchF6(page, { shiftKey: true });
    await expect(page.getByTestId("formula-address")).toBeFocused();

    await dispatchF6(page, { shiftKey: true });
    await expect(activeRibbonTab).toBeFocused();

    await dispatchF6(page, { shiftKey: true });
    await expect(statusBarFirstFocusable).toBeFocused();

    await dispatchF6(page, { shiftKey: true });
    await expect(sheetTabsActiveTab).toBeFocused();

    await dispatchF6(page, { shiftKey: true });
    await expect(page.locator("#grid")).toBeFocused();
  });
});
