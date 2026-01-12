import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

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

test.describe("status bar mode indicator", () => {
  const modes = ["legacy", "shared"] as const;

  for (const mode of modes) {
    test(`shows Ready/Edit when entering cell edit mode (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Avoid grid-mode-specific click coordinates by driving selection via the app API.
      await page.evaluate(() => {
        const app = window.__formulaApp as any;
        app.focus();
        app.activateCell({ row: 0, col: 0 }); // A1
      });

      const statusMode = page.getByTestId("status-mode");

      await expect(statusMode).toHaveText("Ready");

      await page.keyboard.press("F2");
      await expect(statusMode).toHaveText("Edit");

      await page.keyboard.press("Escape");
      await expect(statusMode).toHaveText("Ready");
    });
  }
});
