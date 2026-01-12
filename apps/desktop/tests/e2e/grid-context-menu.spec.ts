import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  // Vite may occasionally trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-wait.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 60_000 });
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

test.describe("Grid context menus", () => {
  test("right-clicking a row header opens a menu with Row Height…", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await expect(page.locator("#grid")).toBeVisible();
    await page.click("#grid", { button: "right", position: { x: 10, y: 40 } });

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Row Height…" })).toBeVisible();
  });

  test("right-clicking a column header opens a menu with Column Width…", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await expect(page.locator("#grid")).toBeVisible();
    await page.click("#grid", { button: "right", position: { x: 100, y: 10 } });

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Column Width…" })).toBeVisible();
  });
});
