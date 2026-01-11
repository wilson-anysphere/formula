import { expect, test } from "@playwright/test";

test("clicking a column header selects the full column in the data region", async ({ page }) => {
  await page.goto("/?e2e=1");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  await page.waitForFunction(() => (window as any).__gridApi?.getSelectionRange, null, { timeout: 30_000 });

  const selectionCanvas = page.getByTestId("canvas-grid-selection");

  // Defaults from `VirtualScrollManager`: col width = 100, row height = 21.
  // Column A header is at row 0, col 1, so click within that header cell.
  await selectionCanvas.click({ position: { x: 150, y: 10 } });

  const range = await page.evaluate(() => (window as any).__gridApi.getSelectionRange());
  expect(range).toEqual({ startRow: 1, endRow: 1_000_001, startCol: 1, endCol: 2 });
});
