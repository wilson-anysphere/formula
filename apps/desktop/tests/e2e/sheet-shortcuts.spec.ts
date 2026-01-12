import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("sheet navigation shortcuts", () => {
  test("Ctrl+PageDown / Ctrl+PageUp switches the active sheet (wraps)", async ({ page }) => {
    await gotoDesktop(page);

    // Ensure the grid has focus and the status bar reflects A1 values.
    await page.click("#grid", { position: { x: 5, y: 5 } });
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();

    // Lazily create Sheet2 by writing a value into it.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    // Next sheet.
    await page.keyboard.press(`${modifier}+PageDown`);
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("Hello from Sheet2");

    // Previous sheet.
    await page.keyboard.press(`${modifier}+PageUp`);
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("Seed");

    // Wrap-around at the start.
    await page.keyboard.press(`${modifier}+PageUp`);
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveAttribute("data-active", "true");

    // Wrap-around at the end.
    await page.keyboard.press(`${modifier}+PageDown`);
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-active", "true");
  });
});

