import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("command palette", () => {
  test("shows keybinding hint for Comments and toggles the comments panel", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("Comments");

    const expectedShortcut = process.platform === "darwin" ? "⇧⌘M" : "Ctrl+Shift+M";
    const item = page
      .getByTestId("command-palette-list")
      .locator(".command-palette__item")
      .filter({ hasText: "Toggle Comments Panel" })
      .first();
    await expect(item).toBeVisible();
    await expect(item.locator(".command-palette__shortcut")).toHaveText(expectedShortcut);

    await page.keyboard.press("Enter");

    await expect(page.getByTestId("comments-panel")).toBeVisible();
  });

  test("shows keybinding hint for Add Comment and focuses the new comment input", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("Add Comment");

    const expectedShortcut = process.platform === "darwin" ? "⇧F2" : "Shift+F2";
    const item = page
      .getByTestId("command-palette-list")
      .locator(".command-palette__item")
      .filter({ hasText: "Add Comment" })
      .first();
    await expect(item).toBeVisible();
    await expect(item.locator(".command-palette__shortcut")).toHaveText(expectedShortcut);

    await page.keyboard.press("Enter");

    await expect(page.getByTestId("comments-panel")).toBeVisible();
    await expect(page.getByTestId("new-comment-input")).toBeFocused();
  });
});
