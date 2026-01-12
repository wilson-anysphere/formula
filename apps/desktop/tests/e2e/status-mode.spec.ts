import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window.__formulaApp as any)?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window.__formulaApp as any).whenIdle());
}

test.describe("status bar mode indicator", () => {
  test("shows Ready by default and toggles to Edit during F2 editing", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await expect(page.getByTestId("status-mode")).toHaveText("Ready");

    // Focus the grid so keyboard shortcuts target the sheet.
    // Click inside A1 (avoid the shared-grid corner header/select-all region).
    await page.click("#grid", { position: { x: 80, y: 40 } });

    await page.keyboard.press("F2");
    await expect(page.getByTestId("status-mode")).toHaveText("Edit");

    await page.keyboard.press("Escape");
    await expect(page.getByTestId("status-mode")).toHaveText("Ready");
  });
});
