import { expect, test } from "@playwright/test";

test.describe("command palette go to", () => {
  test("typing a cell reference navigates immediately", async ({ page }) => {
    await page.goto("/");

    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("B3");
      return rect && typeof rect.x === "number" && rect.width > 0 && rect.height > 0;
    });

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("B3");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("active-cell")).toHaveText("B3");
  });
});

