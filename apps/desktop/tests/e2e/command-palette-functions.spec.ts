import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("command palette - function insertion", () => {
  test("searches Excel functions and inserts a template into the formula bar", async ({ page }) => {
    await gotoDesktop(page, "/");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("sum");
    await expect(page.getByTestId("command-palette-list")).toContainText("SUM");
    await page.keyboard.press("Enter");

    const input = page.getByTestId("formula-input");
    await expect(input).toBeVisible();
    await expect(input).toHaveValue("=SUM()");
    await expect(input).toBeFocused();

    const selection = await page.evaluate(() => {
      const el = document.querySelector<HTMLTextAreaElement>('[data-testid="formula-input"]');
      return { start: el?.selectionStart ?? null, end: el?.selectionEnd ?? null };
    });

    expect(selection).toEqual({ start: 5, end: 5 });
  });
});
