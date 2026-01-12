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
    // Click inside A1 (avoid the shared-grid corner header/select-all region).
    await page.click("#grid", { position: { x: 80, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("Shift+F10");
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    // Built-in items should display the familiar shortcut hints (display-only).
    const expectedCopyShortcut = process.platform === "darwin" ? "⌘C" : "Ctrl+C";
    const copy = menu.getByRole("button", { name: "Copy" });
    await expect(copy.locator('span[aria-hidden="true"]')).toHaveText(expectedCopyShortcut);

    await page.keyboard.press("Escape");
    await expect(menu).toBeHidden();

    // Grid should still receive arrow navigation without requiring an extra click.
    await page.keyboard.press("ArrowRight");
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
  });

  test("Shift+F10 opens the row header menu when a whole row is selected", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.click("#grid", { position: { x: 80, y: 40 } });

    const selectionBefore = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const limits = app?.limits ?? { maxRows: 1_048_576, maxCols: 16_384 };
      app.selectRange({ range: { startRow: 0, endRow: 0, startCol: 0, endCol: limits.maxCols - 1 } });
      return app.getSelectionRanges();
    });

    await page.keyboard.press("Shift+F10");
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Row Height…" })).toBeVisible();

    const selectionAfter = await page.evaluate(() => (window as any).__formulaApp.getSelectionRanges());
    expect(selectionAfter).toEqual(selectionBefore);
  });

  test("Shift+F10 opens the column header menu when a whole column is selected", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.click("#grid", { position: { x: 80, y: 40 } });

    const selectionBefore = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const limits = app?.limits ?? { maxRows: 1_048_576, maxCols: 16_384 };
      app.selectRange({ range: { startRow: 0, endRow: limits.maxRows - 1, startCol: 0, endCol: 0 } });
      return app.getSelectionRanges();
    });

    await page.keyboard.press("Shift+F10");
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(menu.getByRole("button", { name: "Column Width…" })).toBeVisible();

    const selectionAfter = await page.evaluate(() => (window as any).__formulaApp.getSelectionRanges());
    expect(selectionAfter).toEqual(selectionBefore);
  });
});
