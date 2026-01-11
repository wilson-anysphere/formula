import { expect, test } from "@playwright/test";

test("grid container is focusable and announces selection via live region", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const gridContainer = page.getByTestId("canvas-grid");
  await expect(gridContainer).toHaveAttribute("tabindex", "0");
  await expect(gridContainer).toHaveAccessibleName("Spreadsheet grid");

  await gridContainer.focus();
  await expect(gridContainer).toBeFocused();

  const grid = page.getByTestId("canvas-grid-selection");
  await grid.click({ position: { x: 150, y: 31 } }); // A1

  const status = page.getByTestId("canvas-grid-a11y-status");
  await expect(status).toContainText("Active cell A1");
  await expect(status).toContainText("value 1");

  // Keyboard navigation should move the active cell and update the live region.
  await page.keyboard.press("ArrowRight"); // B1
  await expect(status).toContainText("Active cell B1");
  await expect(status).toContainText("value 3");
});
