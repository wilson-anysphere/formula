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
    await page.evaluate(() => (window.__formulaApp as any).whenIdle());

    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getDocument().isDirty)).toBe(true);

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

    expect(beforeUnloadDialogs).toBeGreaterThan(0);
  });

  test("shows a beforeunload confirm dialog when the document is dirty from a sheet rename", async ({ page }) => {
    await gotoDesktop(page);

    const tab = page.getByTestId("sheet-tab-Sheet1");
    await expect(tab).toBeVisible();

    await tab.dblclick();
    const input = tab.locator("input.sheet-tab__input");
    await expect(input).toBeVisible();
    await expect(input).toBeFocused();

    await input.fill("RenamedSheet1");
    await input.press("Enter");

    await expect(tab.locator(".sheet-tab__name")).toHaveText("RenamedSheet1");
    await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getDocument().isDirty)).toBe(true);

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

    expect(beforeUnloadDialogs).toBeGreaterThan(0);
  });
});
