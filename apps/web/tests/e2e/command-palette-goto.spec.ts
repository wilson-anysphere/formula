import { expect, test } from "@playwright/test";

test.describe("command palette go to", () => {
  test("typing a cell reference navigates immediately", async ({ page }) => {
    await page.goto("/?e2e=1");
    await expect(page.getByTestId("engine-status")).toContainText("ready", { timeout: 30_000 });

    const primary = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${primary}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.keyboard.type("B3");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("active-address")).toHaveText("B3");
  });
});

