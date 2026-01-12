import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("AI inline edit (context menu)", () => {
  test.setTimeout(120_000);

  test("opens from the grid context menu", async ({ page }) => {
    await gotoDesktop(page);

    const grid = page.locator("#grid");
    const a1Rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));

    // Select A1 before opening the context menu.
    await grid.click({ position: { x: a1Rect.x + a1Rect.width / 2, y: a1Rect.y + a1Rect.height / 2 } });

    // Open the grid context menu and run "Inline AI Edit…".
    await grid.click({
      button: "right",
      position: { x: a1Rect.x + a1Rect.width / 2, y: a1Rect.y + a1Rect.height / 2 },
    });

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const item = menu.getByRole("button", { name: /Inline AI Edit/ });
    await expect(item).toBeEnabled();
    const expectedShortcut = process.platform === "darwin" ? "⌘K" : "Ctrl+K";
    await expect(item.locator('span[aria-hidden="true"]')).toHaveText(expectedShortcut);
    await item.click();

    await expect(page.getByTestId("inline-edit-overlay")).toBeVisible();
  });
});
