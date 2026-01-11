import { expect, test } from "@playwright/test";

test("editing A1 updates dependent formula cells via engine recalc", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const grid = page.getByTestId("canvas-grid-selection");

  // Click A1 (row 1, col 1), type 10, commit.
  await grid.click({ position: { x: 150, y: 31 } });
  await expect(page.getByTestId("active-address")).toHaveText("A1");

  const input = page.getByTestId("formula-input");
  await expect(input).toBeEnabled();
  await input.fill("10");
  await input.press("Enter");

  await expect(page.getByTestId("formula-bar-value")).toHaveText("10");

  // Click B1 (row 1, col 2) and verify recalculated value.
  await grid.click({ position: { x: 250, y: 31 } });
  await expect(page.getByTestId("active-address")).toHaveText("B1");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("12");
});
