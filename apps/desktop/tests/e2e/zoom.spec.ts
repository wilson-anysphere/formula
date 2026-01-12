import { expect, test } from "@playwright/test";
import { gotoDesktop } from "./helpers";

test("shared grid zoom updates geometry + status bar", async ({ page }) => {
  await gotoDesktop(page, "/?grid=shared");

  await page.evaluate(() => (window as any).__formulaApp.whenIdle());

  await expect(page.getByTestId("status-zoom")).toHaveText("100%");

  const before = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));
  expect(before).not.toBeNull();

  await page.selectOption('[data-testid="zoom-control"]', "150");

  await page.waitForFunction(() => Math.abs((window as any).__formulaApp.getZoom() - 1.5) < 1e-6);

  const after = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));
  expect(after).not.toBeNull();
  expect(after!.width).toBeGreaterThan(before!.width);

  await expect(page.getByTestId("status-zoom")).toHaveText("150%");
});
