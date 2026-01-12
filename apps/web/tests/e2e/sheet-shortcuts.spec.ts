import { expect, test } from "@playwright/test";

test("Ctrl/Cmd+PageUp/PageDown switches sheets with wrap-around", async ({ page }) => {
  await page.goto("/?e2e=1");
  await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

  const primary = process.platform === "darwin" ? "Meta" : "Control";
  const sheetSelect = page.getByRole("combobox").first();
  const selectionCanvas = page.getByTestId("canvas-grid-selection");
  const formulaInput = page.getByTestId("formula-input");

  // Click A1 (row 1, col 1) to focus the grid.
  await selectionCanvas.click({ position: { x: 150, y: 31 } });
  await expect(page.getByTestId("active-address")).toHaveText("A1");
  await expect(formulaInput).toHaveValue("1");
  await expect(sheetSelect).toHaveValue("Sheet1");

  await page.keyboard.press(`${primary}+PageDown`);
  await expect(sheetSelect).toHaveValue("Sheet2");
  await expect(formulaInput).toHaveValue("Hello from Sheet2");
  await expect(page.getByTestId("canvas-grid")).toBeFocused();

  // Wrap around.
  await page.keyboard.press(`${primary}+PageDown`);
  await expect(sheetSelect).toHaveValue("Sheet1");
  await expect(formulaInput).toHaveValue("1");

  await page.keyboard.press(`${primary}+PageUp`);
  await expect(sheetSelect).toHaveValue("Sheet2");
  await expect(formulaInput).toHaveValue("Hello from Sheet2");
});
