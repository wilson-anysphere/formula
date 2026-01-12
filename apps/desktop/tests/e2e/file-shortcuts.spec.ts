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
    await page.evaluate(() => (window.__formulaApp as any).focus());

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

  test("Ctrl/Cmd+S does not trigger the save command while typing in an input", async ({ page }) => {
    await gotoDesktop(page);

    // Clear any startup toasts so the assertion is deterministic.
    await page.evaluate(() => {
      document.getElementById("toast-root")?.replaceChildren();
    });

    // Enter formula-bar edit mode (focus the textarea).
    await page.getByTestId("formula-highlight").click();
    await expect(page.getByTestId("formula-input")).toBeFocused();

    const defaultPrevented = await page.evaluate((isMac) => {
      const target = document.activeElement ?? window;
      const e = new KeyboardEvent("keydown", {
        key: "s",
        ctrlKey: !isMac,
        metaKey: isMac,
        bubbles: true,
        cancelable: true,
      });
      target.dispatchEvent(e);
      return e.defaultPrevented;
    }, process.platform === "darwin");

    // The file shortcut should be gated by focus.inTextInput == false.
    expect(defaultPrevented).toBe(false);
    await expect(page.getByTestId("toast")).toHaveCount(0);
  });
});
