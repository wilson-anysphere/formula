import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("file shortcuts", () => {
  test("Ctrl/Cmd+S triggers the save workbook command (and prevents browser default)", async ({ page }) => {
    await gotoDesktop(page);

    // Clear any startup toasts so the assertion is deterministic.
    await page.evaluate(() => {
      document.getElementById("toast-root")?.replaceChildren();
    });

    // Ensure the grid has focus for key events.
    await page.evaluate(() => (window as any).__formulaApp.focus());

    // Dispatch a synthetic keydown event instead of relying on the browser shell's
    // handling of Ctrl/Cmd+S, which can be intercepted as a "Save page" shortcut.
    const defaultPrevented = await page.evaluate((isMac) => {
      const e = new KeyboardEvent("keydown", {
        key: "s",
        ctrlKey: !isMac,
        metaKey: isMac,
        bubbles: true,
        cancelable: true,
      });
      window.dispatchEvent(e);
      return e.defaultPrevented;
    }, process.platform === "darwin");

    expect(defaultPrevented).toBe(true);
    await expect(page.getByTestId("toast").first()).toHaveText(
      "Desktop-only: Saving workbooks is available in the desktop app.",
    );
  });
});
