import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("formula bar - fx function picker", () => {
  test("inserts a selected function template into the formula bar with the cursor inside parentheses", async ({
    page,
  }) => {
    await gotoDesktop(page, "/");

    // 1) Focus the formula bar.
    await page.getByTestId("formula-highlight").click();
    const formulaInput = page.getByTestId("formula-input");
    await expect(formulaInput).toBeVisible();

    // 2) Click fx to open the function picker.
    await page.getByTestId("formula-fx-button").click();

    // 3) Search/select a known function.
    const pickerInput = page.getByTestId("formula-function-picker-input");
    await expect(pickerInput).toBeVisible();
    await pickerInput.fill("sum");
    const sumOption = page.getByTestId("formula-function-picker-item-SUM");
    await expect(sumOption).toBeVisible();
    await sumOption.click();

    // 4) Assert formula bar text becomes `=SUM()` with cursor inside parens.
    await expect(formulaInput).toHaveValue("=SUM()");
    await expect(formulaInput).toBeFocused();

    await page.keyboard.type("1");
    await expect(formulaInput).toHaveValue("=SUM(1)");
  });
});
