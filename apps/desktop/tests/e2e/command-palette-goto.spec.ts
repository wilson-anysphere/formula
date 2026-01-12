import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("command palette go to", () => {
  test("typing a cell reference navigates immediately", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("B3");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("active-cell")).toHaveText("B3");
  });

  test("typing a range reference selects the range", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("B3:D4");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("selection-range")).toHaveText("B3:D4");
    await expect(page.getByTestId("active-cell")).toHaveText("B3");
  });

  test("typing a sheet-qualified reference switches sheets", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "C3", "Hello from Sheet2 C3");
    });

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("Sheet2!C3");
    await page.keyboard.press("Enter");

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId())).toBe("Sheet2");
    await expect(page.getByTestId("active-cell")).toHaveText("C3");
    await expect(page.getByTestId("active-value")).toHaveText("Hello from Sheet2 C3");
  });
});
