import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("Edit Cell command", () => {
  test("command palette runs Edit Cell and focuses the inline editor", async ({ page }) => {
    await gotoDesktop(page);

    // Select A1.
    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("A1");
      return rect && typeof rect.x === "number" && rect.width > 0 && rect.height > 0;
    });
    const a1 = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      return app.getCellRectA1("A1");
    });
    await page.click("#grid", { position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("Edit Cell");
    await page.keyboard.press("Enter");

    const editor = page.getByTestId("cell-editor");
    await expect(editor).toBeVisible();
    await expect(editor).toBeFocused();
  });

  test("shortcut search mode (/ f2) runs Edit Cell and focuses the inline editor", async ({ page }) => {
    await gotoDesktop(page);

    // Select A1.
    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("A1");
      return rect && typeof rect.x === "number" && rect.width > 0 && rect.height > 0;
    });
    const a1 = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      return app.getCellRectA1("A1");
    });
    await page.click("#grid", { position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("/ f2");

    const editCell = page
      .locator("li.command-palette__item", { hasText: "Edit Cell" })
      .filter({ hasText: "Edit the active cell" })
      .first();
    await expect(editCell).toBeVisible();
    await editCell.click();

    const editor = page.getByTestId("cell-editor");
    await expect(editor).toBeVisible();
    await expect(editor).toBeFocused();
  });

  test("F2 keybinding runs Edit Cell and focuses the inline editor", async ({ page }) => {
    await gotoDesktop(page);

    // Select A1.
    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("A1");
      return rect && typeof rect.x === "number" && rect.width > 0 && rect.height > 0;
    });
    const a1 = await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      return app.getCellRectA1("A1");
    });
    await page.click("#grid", { position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("F2");

    const editor = page.getByTestId("cell-editor");
    await expect(editor).toBeVisible();
    await expect(editor).toBeFocused();
  });
});
