import { expect, test } from "@playwright/test";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("grid scrolling + virtualization", () => {
  test("wheel scroll down reaches far rows and clicking selects correct cell", async ({ page }) => {
    await page.goto("/");

    // Seed A200 (0-based row 199) with a sentinel string.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.getDocument().setCellValue(sheetId, { row: 199, col: 0 }, "Bottom");
      app.refresh();
    });
    await waitForIdle(page);

    // Wheel-scroll so that row 200 is near the top of the viewport.
    // (cellHeight is currently fixed at 24px in SpreadsheetApp.)
    const grid = page.locator("#grid");
    await grid.hover({ position: { x: 60, y: 40 } });
    await page.mouse.wheel(0, 199 * 24);

    await expect.poll(async () => {
      return await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    }).toBeGreaterThan(0);

    // Click within A200.
    await grid.click({ position: { x: 60, y: 24 + 12 } });

    await expect(page.getByTestId("active-cell")).toHaveText("A200");
    await expect(page.getByTestId("active-value")).toHaveText("Bottom");
  });

  test("ArrowDown navigation auto-scrolls to keep the active cell visible", async ({ page }) => {
    await page.goto("/");
    const grid = page.locator("#grid");

    // Focus A1 (account for headers).
    await grid.click({ position: { x: 60, y: 40 } });

    const scrollBefore = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scrollBefore).toBe(0);

    for (let i = 0; i < 200; i += 1) {
      await page.keyboard.press("ArrowDown");
    }

    await expect(page.getByTestId("active-cell")).toHaveText("A201");
    const scrollAfter = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scrollAfter).toBeGreaterThan(scrollBefore);
  });

  test("name box Go To scrolls and updates selection", async ({ page }) => {
    await page.goto("/");

    const address = page.getByTestId("formula-address");
    await address.click();
    await address.fill("A500");
    await address.press("Enter");

    await expect(page.getByTestId("active-cell")).toHaveText("A500");
    const scroll = await page.evaluate(() => (window as any).__formulaApp.getScroll().y);
    expect(scroll).toBeGreaterThan(0);
  });
});

