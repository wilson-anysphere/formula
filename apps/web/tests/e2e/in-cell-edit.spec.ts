import { expect, test } from "@playwright/test";

test("type-to-edit commits the cell and updates dependent formulas", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const selectionCanvas = page.getByTestId("canvas-grid-selection");

  // Click A1 (row 1, col 1) to focus the grid.
  await selectionCanvas.click({ position: { x: 150, y: 31 } });
  await expect(page.getByTestId("active-address")).toHaveText("A1");

  // Start editing by typing. This should open the in-cell editor overlay.
  await page.keyboard.type("10");
  const editor = page.getByTestId("cell-editor");
  await expect(editor).toBeVisible();

  // Commit with Tab, which should move the selection to the right (B1).
  await editor.press("Tab");
  await expect(editor).toHaveCount(0);

  await expect(page.getByTestId("active-address")).toHaveText("B1");
  await expect(page.getByTestId("formula-bar-value")).toHaveText("12");
});
