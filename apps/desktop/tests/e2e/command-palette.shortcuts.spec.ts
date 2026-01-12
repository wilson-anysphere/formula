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

    // Shortcut mode should only render rows with non-empty shortcut pills.
    const shortcuts = await page.locator("li.command-palette__item").evaluateAll((els) =>
      els.map((el) => (el.querySelector(".command-palette__shortcut")?.textContent ?? "").trim()),
    );
    expect(shortcuts.length).toBeGreaterThan(0);
    expect(shortcuts.every((text) => text.length > 0)).toBe(true);

    // Shortcut mode should still return matching commands with shortcuts.
    await input.fill("/ copy");
    const copy = page.locator("li.command-palette__item", { hasText: "Copy" }).first();
    await expect(copy).toBeVisible();
    await expect(copy.locator(".command-palette__shortcut")).toHaveText(expectedCopyShortcut);

    // When a command has multiple keybindings, shortcut search should display the binding that matched the query.
    // Copy has both the primary chord and an explicit Ctrl+Cmd/Meta fallback chord for remote desktop setups.
    const expectedFallbackCopyShortcut = process.platform === "darwin" ? "⌃⌘C" : "Ctrl+Meta+C";
    await input.fill("/ ctrl+cmd+c");
    const fallbackCopy = page.locator("li.command-palette__item", { hasText: "Copy" }).first();
    await expect(fallbackCopy).toBeVisible();
    await expect(fallbackCopy.locator(".command-palette__shortcut")).toHaveText(expectedFallbackCopyShortcut);

    // Toggle Comments Panel has both the primary chord and an explicit Ctrl+Cmd/Meta fallback chord for remote desktop setups.
    const expectedFallbackCommentsShortcut = process.platform === "darwin" ? "⌃⇧⌘M" : "Ctrl+Shift+Meta+M";
    await input.fill("/ ctrl+cmd+shift+m");
    const comments = page.locator("li.command-palette__item", { hasText: "Toggle Comments Panel" }).first();
    await expect(comments).toBeVisible();
    await expect(comments.locator(".command-palette__shortcut")).toHaveText(expectedFallbackCommentsShortcut);

    // Add Comment (Shift+F2).
    const expectedAddCommentShortcut = process.platform === "darwin" ? "⇧F2" : "Shift+F2";
    await input.fill("/ shift+f2");
    const addComment = page.locator("li.command-palette__item", { hasText: "Add Comment" }).first();
    await expect(addComment).toBeVisible();
    await expect(addComment.locator(".command-palette__shortcut")).toHaveText(expectedAddCommentShortcut);

    // Function keys should be searchable as tokens (e.g. `/ f2`).
    await input.fill("/ f2");
    const editCell = page
      .locator("li.command-palette__item", { hasText: "Edit Cell" })
      .filter({ hasText: "Edit the active cell" })
      .first();
    await expect(editCell).toBeVisible();
    await expect(editCell.locator(".command-palette__shortcut")).toHaveText("F2");
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

  test("renders the platform shortcut hint for Save", async ({ page }) => {
    await gotoDesktop(page);

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    const expectedSaveShortcut = process.platform === "darwin" ? "⌘S" : "Ctrl+S";

    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    // Use the command id to avoid ambiguity (other commands can match "save").
    await page.getByTestId("command-palette-input").fill("workbench.saveWorkbook");

    const item = page
      .locator("li.command-palette__item")
      .filter({ has: page.locator(".command-palette__item-label", { hasText: /^Save$/ }) })
      .first();
    await expect(item).toBeVisible();
    await expect(item.locator(".command-palette__shortcut")).toHaveText(expectedSaveShortcut);
  });
});
