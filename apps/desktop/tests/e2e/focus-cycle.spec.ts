import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

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

async function dispatchCtrlOrCmdShiftP(page: import("@playwright/test").Page): Promise<void> {
  // `Ctrl/Cmd+Shift+P` can be intercepted by the browser/OS in some environments.
  // Dispatching a synthetic `keydown` keeps the test deterministic and still exercises
  // our in-app keybinding pipeline (KeybindingService -> CommandRegistry).
  await page.evaluate((isMac) => {
    const target = (document.activeElement as HTMLElement | null) ?? document.getElementById("grid") ?? window;
    target.dispatchEvent(
      new KeyboardEvent("keydown", {
        key: "P",
        code: "KeyP",
        shiftKey: true,
        metaKey: isMac,
        ctrlKey: !isMac,
        bubbles: true,
        cancelable: true,
      }),
    );
  }, process.platform === "darwin");
}

test.describe("focus cycling (Excel-style F6)", () => {
  test("F6 / Shift+F6 cycle focus between ribbon, formula bar, grid, sheet tabs, and status bar", async ({ page }) => {
    await gotoDesktop(page);

    // Start from the grid.
    await page.click("#grid", { position: { x: 80, y: 40 } });
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    // Forward cycle (matches apps/desktop/src/commands/workbenchFocusCycle.ts):
    // ribbon -> formula bar -> grid -> sheet tabs -> status bar -> ribbon
    //
    // Starting from the grid, that means: sheet tabs -> status bar -> ribbon -> formula bar -> grid.
    await dispatchF6(page);
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeFocused();

    await dispatchF6(page);
    await expect(page.getByTestId("zoom-control")).toBeFocused();

    await dispatchF6(page);
    await expect(page.getByTestId("ribbon-tab-home")).toBeFocused();

    await dispatchF6(page);
    await expect(page.getByTestId("formula-address")).toBeFocused();

    await dispatchF6(page);
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    // One more forward press returns to sheet tabs (wrap).
    await dispatchF6(page);
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeFocused();

    // Reverse cycle.
    await dispatchF6(page, { shiftKey: true });
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    await dispatchF6(page, { shiftKey: true });
    await expect(page.getByTestId("formula-address")).toBeFocused();

    await dispatchF6(page, { shiftKey: true });
    await expect(page.getByTestId("ribbon-tab-home")).toBeFocused();

    await dispatchF6(page, { shiftKey: true });
    await expect(page.getByTestId("zoom-control")).toBeFocused();

    await dispatchF6(page, { shiftKey: true });
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeFocused();
  });

  test("F6 does not cycle focus while the command palette (keybinding barrier) is open", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure the grid has focus before opening the palette.
    await page.click("#grid", { position: { x: 80, y: 40 } });
    await expect
      .poll(() => page.evaluate(() => (document.activeElement as HTMLElement | null)?.id))
      .toBe("grid");

    await dispatchCtrlOrCmdShiftP(page);
    await expect(page.getByTestId("command-palette")).toBeVisible();
    await expect(page.getByTestId("command-palette-input")).toBeFocused();

    // F6 should not escape the palette's focus trap / keybinding barrier.
    await dispatchF6(page);
    await expect(page.getByTestId("command-palette-input")).toBeFocused();

    await dispatchF6(page, { shiftKey: true });
    await expect(page.getByTestId("command-palette-input")).toBeFocused();
  });

  test("F6 does not cycle focus while a ribbon dropdown menu (keybinding barrier) is open", async ({ page }) => {
    await gotoDesktop(page);

    const viewTab = page.getByTestId("ribbon-tab-view");
    await expect(viewTab).toBeVisible();
    await viewTab.click();

    const themeSelector = page.getByTestId("theme-selector");
    await expect(themeSelector).toBeVisible();
    await themeSelector.click();

    // Menu open signal.
    await expect(page.getByTestId("theme-option-dark")).toBeVisible();

    await expect(themeSelector).toBeFocused();

    await dispatchF6(page);
    await expect(themeSelector).toBeFocused();

    await dispatchF6(page, { shiftKey: true });
    await expect(themeSelector).toBeFocused();
  });
});
