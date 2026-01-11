import { expect, test } from "@playwright/test";

test("supports keyboard navigation and PageDown", async ({ page }) => {
  await page.goto("/");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const selectionCanvas = page.getByTestId("canvas-grid-selection");

  // Click A1 (row 1, col 1) to focus the grid.
  await selectionCanvas.click({ position: { x: 150, y: 31 } });
  await expect(page.getByTestId("active-address")).toHaveText("A1");

  await page.keyboard.press("ArrowRight");
  await expect(page.getByTestId("active-address")).toHaveText("B1");

  await page.keyboard.press("ArrowDown");
  await expect(page.getByTestId("active-address")).toHaveText("B2");

  await page.keyboard.press("PageDown");

  const after = await page.getByTestId("active-address").textContent();
  expect(after).toBeTruthy();
  const match = after!.match(/^B(\d+)$/);
  expect(match).not.toBeNull();
  expect(Number(match![1])).toBeGreaterThan(2);
});

