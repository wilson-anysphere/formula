import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("collab: beforeunload unsaved-changes prompt", () => {
  test("does not show a beforeunload confirm dialog when a collab session is active", async ({ page }) => {
    await gotoDesktop(page);

    // Create a user gesture + local edit so browsers are allowed to show a beforeunload prompt.
    await page.click("#grid", { position: { x: 5, y: 5 } });
    await page.keyboard.press("h");
    await page.keyboard.type("ello");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(true);

    // Simulate collab mode by attaching the API expected by main.ts. In real collab mode
    // this is provided by the collaboration bootstrap layer.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getCollabSession = () => ({ id: "e2e-collab-session" });
    });

    let beforeUnloadDialogs = 0;
    page.on("dialog", async (dialog) => {
      if (dialog.type() === "beforeunload") beforeUnloadDialogs += 1;
      await dialog.accept();
    });

    await page.reload();
    await waitForDesktopReady(page);

    expect(beforeUnloadDialogs).toBe(0);
  });
});

