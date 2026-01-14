import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";
import { installCollabSessionStub } from "./collabSessionStub";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window.__formulaApp as any).whenIdle());
}

test.describe("collab: beforeunload unsaved-changes prompt", () => {
  test("does not show a beforeunload confirm dialog when a collab session is active", async ({ page }) => {
    await gotoDesktop(page);

    // Create a user gesture + local edit so browsers are allowed to show a beforeunload prompt.
    // Click inside A1 (avoid the shared-grid corner header/select-all region).
    await page.click("#grid", { position: { x: 80, y: 40 } });
    await page.keyboard.press("h");
    await page.keyboard.type("ello");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getDocument().isDirty)).toBe(true);


    await installCollabSessionStub(page);

    let beforeUnloadDialogs = 0;
    page.on("dialog", (dialog) => {
      void (async () => {
        if (dialog.type() === "beforeunload") beforeUnloadDialogs += 1;
        await dialog.accept();
      })().catch(() => {
        // Best-effort: don't surface unhandled rejections from Playwright event handlers.
      });
    });

    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    expect(beforeUnloadDialogs).toBe(0);
  });
});
