import { expect, test } from "@playwright/test";

test.describe("sheet tabs", () => {
  test("switching sheets updates the visible cell values", async ({ page }) => {
    await page.goto("/");
    await page.waitForFunction(() => (window as any).__formulaApp != null);

    // Ensure A1 is active before switching sheets so the status bar reflects A1 values.
    await page.click("#grid", { position: { x: 5, y: 5 } });
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();

    // Lazily create Sheet2 by writing a value into it.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    await page.getByTestId("sheet-tab-Sheet2").click();
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("Hello from Sheet2");

    // Switching back restores the original Sheet1 value.
    await page.getByTestId("sheet-tab-Sheet1").click();
    await expect(page.getByTestId("active-value")).toHaveText("Seed");
  });
});
