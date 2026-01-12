import { expect, test } from "@playwright/test";

test("gridApi.setZoom scales cell geometry", async ({ page }) => {
  await page.goto("/?e2e=1");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  await page.waitForFunction(() => {
    const api = (window as any).__gridApi;
    return api && typeof api.getCellRect === "function" && typeof api.setZoom === "function";
  });

  await page.evaluate(() => {
    (window as any).__gridApi.scrollTo(0, 0);
    (window as any).__gridApi.setZoom(1);
  });

  const before = await page.evaluate(() => (window as any).__gridApi.getCellRect(1, 1));
  expect(before).not.toBeNull();

  const after = await page.evaluate(() => {
    (window as any).__gridApi.setZoom(2);
    return (window as any).__gridApi.getCellRect(1, 1);
  });
  expect(after).not.toBeNull();

  expect(after!.width).toBeGreaterThan(before!.width * 1.9);
  expect(after!.width).toBeLessThan(before!.width * 2.1);
  expect(after!.height).toBeGreaterThan(before!.height * 1.9);
  expect(after!.height).toBeLessThan(before!.height * 2.1);
});
