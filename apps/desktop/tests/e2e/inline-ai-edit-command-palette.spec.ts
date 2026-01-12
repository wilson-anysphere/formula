import { expect, test } from "@playwright/test";
import { gotoDesktop } from "./helpers";

test.describe("AI inline edit command", () => {
  test("opens from the command palette", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure keyboard shortcuts are dispatched to the grid (not the browser UI).
    await page.locator("#grid").click();

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);

    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Inline AI Edit");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("inline-edit-overlay")).toBeVisible();
  });
});
