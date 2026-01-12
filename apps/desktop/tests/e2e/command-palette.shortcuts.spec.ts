import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("command palette shortcut hints", () => {
  test("renders the platform shortcut hint for Copy", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    const expectedCopyShortcut = process.platform === "darwin" ? "⌘C" : "Ctrl+C";

    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("copy");

    const item = page.locator("li.command-palette__item", { hasText: "Copy" }).first();
    await expect(item).toBeVisible();
    await expect(item.locator(".command-palette__shortcut")).toHaveText(expectedCopyShortcut);
  });

  test("renders the platform shortcut hint for Replace", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    const expectedReplaceShortcut = process.platform === "darwin" ? "⌥⌘F" : "Ctrl+H";

    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("replace");

    const item = page.locator("li.command-palette__item", { hasText: "Replace" }).first();
    await expect(item).toBeVisible();
    await expect(item.locator(".command-palette__shortcut")).toHaveText(expectedReplaceShortcut);
  });
});
