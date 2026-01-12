import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("non-collab: beforeunload unsaved-changes prompt", () => {
  test("shows a beforeunload confirm dialog when the document is dirty", async ({ page }) => {
    await gotoDesktop(page);

    // Create a user gesture + local edit so browsers are allowed to show a beforeunload prompt.
    await page.click("#grid", { position: { x: 80, y: 40 } });
    await page.keyboard.press("h");
    await page.keyboard.type("ello");
    await page.keyboard.press("Enter");
    await page.evaluate(() => (window as any).__formulaApp.whenIdle());

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(true);

    let beforeUnloadDialogs = 0;
    page.on("dialog", async (dialog) => {
      if (dialog.type() === "beforeunload") beforeUnloadDialogs += 1;
      await dialog.accept();
    });

    await page.reload();
    await waitForDesktopReady(page);

    expect(beforeUnloadDialogs).toBeGreaterThan(0);
  });
});

