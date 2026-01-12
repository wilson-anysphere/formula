import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("Clear Contents context menu", () => {
  test("clears a single cell via the cell context menu", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, { row: 0, col: 0 }, "Hello");
    });
    await waitForIdle(page);

    const grid = page.locator("#grid");
    const a1Rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));
    await grid.click({ position: { x: a1Rect.x + a1Rect.width / 2, y: a1Rect.y + a1Rect.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await grid.click({
      button: "right",
      position: { x: a1Rect.x + a1Rect.width / 2, y: a1Rect.y + a1Rect.height / 2 },
    });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const item = menu.getByRole("button", { name: "Clear Contents" });
    await expect(item).toBeEnabled();
    await item.click();

    await waitForIdle(page);
    const value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(value).toBe("");
  });

  test("clears a single cell via the command palette", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, { row: 0, col: 0 }, "Hello");
    });
    await waitForIdle(page);

    const grid = page.locator("#grid");
    const a1Rect = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));
    await grid.click({ position: { x: a1Rect.x + a1Rect.width / 2, y: a1Rect.y + a1Rect.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);

    const input = page.getByTestId("command-palette-input");
    await expect(input).toBeVisible();
    await input.fill("Clear Contents");
    await page.keyboard.press("Enter");

    await waitForIdle(page);
    const value = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(value).toBe("");
  });
});
