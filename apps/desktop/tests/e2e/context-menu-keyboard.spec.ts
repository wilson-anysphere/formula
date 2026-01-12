import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("grid context menu keyboard invocation", () => {
  test("Shift+F10 opens the context menu without changing selection", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    // Ensure the grid is focused and has an active cell.
    await page.click("#grid", { position: { x: 5, y: 5 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("Shift+F10");
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("Escape");
    await expect(menu).toBeHidden();

    // Grid should still receive arrow navigation without requiring an extra click.
    await page.keyboard.press("ArrowRight");
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
  });
});
