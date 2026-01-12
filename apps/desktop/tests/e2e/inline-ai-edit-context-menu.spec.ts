import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

const PERMISSIONS_KEY = "formula.extensionHost.permissions";
const E2E_EVENTS_EXTENSION_ID = "formula.e2e-events";

test.describe("AI inline edit (context menu)", () => {
  test.setTimeout(120_000);

  test("opens from the grid context menu", async ({ page }) => {
    // Right-clicking the grid triggers deferred extension loading so extension-contributed
    // menu items can be surfaced. The built-in e2e extension activates on startup and writes
    // to extension storage, which can prompt for the `storage` permission on first run.
    //
    // Seed a grant so permission prompts don't intercept the context menu click in this test.
    await page.addInitScript(
      ({ key, extensionId }) => {
        try {
          const raw = localStorage.getItem(key);
          const parsed = raw ? JSON.parse(raw) : {};
          const store =
            parsed && typeof parsed === "object" && !Array.isArray(parsed) ? (parsed as Record<string, any>) : {};
          store[extensionId] = { ...(store[extensionId] ?? {}), storage: true };
          localStorage.setItem(key, JSON.stringify(store));
        } catch {
          // Best-effort: if localStorage is unavailable, the app will still show a prompt and
          // the test may fail (which is preferable to silently passing with no prompt UI).
        }
      },
      { key: PERMISSIONS_KEY, extensionId: E2E_EVENTS_EXTENSION_ID },
    );

    await gotoDesktop(page);

    const grid = page.locator("#grid");
    const a1Rect = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("A1"));

    // Select A1 before opening the context menu.
    await grid.click({ position: { x: a1Rect.x + a1Rect.width / 2, y: a1Rect.y + a1Rect.height / 2 } });

    // Open the grid context menu and run "Inline AI Edit…".
    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(
      ({ x, y }) => {
        const grid = document.getElementById("grid");
        if (!grid) throw new Error("Missing #grid container");
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
      { x: a1Rect.x + a1Rect.width / 2, y: a1Rect.y + a1Rect.height / 2 },
    );

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const expectedShortcut = process.platform === "darwin" ? "⌘K" : "Ctrl+K";

    const item = menu.getByRole("button", { name: /Inline AI Edit/ });
    await expect(item).toBeEnabled();
    await expect(item.locator('span[aria-hidden="true"]')).toHaveText(expectedShortcut);
    await item.click();

    await expect(page.getByTestId("inline-edit-overlay")).toBeVisible();
  });
});
