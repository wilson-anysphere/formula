import { expect, test } from "@playwright/test";

test("loads wasm engine and evaluates a simple formula", async ({ page }) => {
  await page.goto("/");

  await expect(page.getByTestId("engine-status")).toContainText("ready");
  await expect(page.getByTestId("engine-status")).toContainText("B1=3");
});
