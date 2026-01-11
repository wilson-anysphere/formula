import { expect, test } from "@playwright/test";

test("gridApi.setZoom scales cell geometry", async ({ page }) => {
  await page.goto("/?e2e=1");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  await page.waitForFunction(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).__gridApi as any;
    return api && typeof api.setZoom === "function" && typeof api.getCellRect === "function";
  });

  const before = await page.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (window as any).__gridApi.getCellRect(1, 1);
  });
  expect(before).not.toBeNull();

  const after = await page.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (window as any).__gridApi.setZoom(2);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (window as any).__gridApi.getCellRect(1, 1);
  });
  expect(after).not.toBeNull();

  expect(after!.width).toBeCloseTo(before!.width * 2, 5);
  expect(after!.height).toBeCloseTo(before!.height * 2, 5);
});
