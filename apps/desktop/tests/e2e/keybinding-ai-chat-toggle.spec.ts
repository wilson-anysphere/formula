import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("keybindings: AI Chat toggle", () => {
  test("Ctrl+Shift+A (Win/Linux) / Cmd+I (macOS) toggles AI Chat without triggering Select All", async ({ page }) => {
    await gotoDesktop(page);
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    // Select B2 first so we can assert the shortcut doesn't fall through to Ctrl/Cmd+A (Select All).
    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("B2");
      return rect && rect.width > 0 && rect.height > 0;
    });

    const b2 = (await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("B2"))) as {
      x: number;
      y: number;
      width: number;
      height: number;
    };

    const grid = page.locator("#grid");
    await grid.click({ position: { x: b2.x + b2.width / 2, y: b2.y + b2.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("B2");

    const shortcut =
      process.platform === "darwin"
        ? // macOS: Cmd+I toggles AI chat.
          { metaKey: true, shiftKey: false, key: "I", code: "KeyI" }
        : // Windows/Linux: Ctrl+Shift+A toggles AI chat.
          { ctrlKey: true, shiftKey: true, key: "A", code: "KeyA" };

    const dispatchShortcut = async () => {
      await page.evaluate((keys) => {
        const gridEl = document.getElementById("grid");
        if (!gridEl) throw new Error("Missing #grid");
        gridEl.dispatchEvent(
          new KeyboardEvent("keydown", { bubbles: true, cancelable: true, ...keys }),
        );
      }, shortcut);
    };

    await dispatchShortcut();
    await expect(page.getByTestId("dock-right").getByTestId("panel-aiChat")).toBeVisible();
    await expect(page.getByTestId("active-cell")).toHaveText("B2");

    await dispatchShortcut();
    await expect(page.getByTestId("panel-aiChat")).toHaveCount(0);
    await expect(page.getByTestId("active-cell")).toHaveText("B2");
  });
});
