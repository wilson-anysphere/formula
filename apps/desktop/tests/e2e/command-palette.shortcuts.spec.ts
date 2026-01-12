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

  test("renders the shortcut hint for Edit Cell (F2)", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";

    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("edit cell");

    const item = page
      .locator("li.command-palette__item", { hasText: "Edit Cell" })
      .filter({ hasText: "Edit the active cell" })
      .first();
    await expect(item).toBeVisible();
    await expect(item.locator(".command-palette__shortcut")).toHaveText("F2");
  });

  test("supports '/' shortcut search mode", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    const expectedCopyShortcut = process.platform === "darwin" ? "⌘C" : "Ctrl+C";

    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    const input = page.getByTestId("command-palette-input");

    // Enter shortcut mode with a leading '/' (after trimming).
    await input.fill(" /");
    await expect(page.locator(".command-palette__hint")).toBeVisible();

    // Shortcut mode should only render rows with visible shortcut pills.
    await expect(page.locator(".command-palette__item-right[hidden]")).toHaveCount(0);

    // Shortcut mode should still return matching commands with shortcuts.
    await input.fill("/ copy");
    const copy = page.locator("li.command-palette__item", { hasText: "Copy" }).first();
    await expect(copy).toBeVisible();
    await expect(copy.locator(".command-palette__shortcut")).toHaveText(expectedCopyShortcut);
  });

  test("renders the platform shortcut hint for Replace", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    const expectedReplaceShortcut = process.platform === "darwin" ? "⌥⌘F" : "Ctrl+H";

    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("replace");

    // "replace" also matches the REPLACE() function result. Ensure we're asserting on the
    // *command* (Replace…) row, not the spreadsheet function suggestion.
    const item = page
      .locator("li.command-palette__item", { hasText: "Replace…" })
      .filter({ hasText: "Show the Replace dialog" })
      .first();
    await expect(item).toBeVisible();
    await expect(item.locator(".command-palette__shortcut")).toHaveText(expectedReplaceShortcut);
  });
});
