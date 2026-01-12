import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("command palette (Show Formulas)", () => {
  test("shows keybinding hint and runs view.toggleShowFormulas", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellFormula(sheetId, "A1", "=1+1", { label: "Set formula" });
      app.refresh();
      await app.whenIdle();
    });

    const before = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      return await app.getCellDisplayTextForRenderA1("A1");
    });
    expect(before).toBe("2");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("Show Formulas");

    const expectedShortcut = process.platform === "darwin" ? "⌘`" : "Ctrl+`";
    const item = page
      .getByTestId("command-palette-list")
      .locator(".command-palette__item")
      .filter({ hasText: "Show Formulas" })
      .first();
    await expect(item).toBeVisible();
    await expect(item.locator(".command-palette__shortcut")).toHaveText(expectedShortcut);

    await item.click();
    await expect(page.getByTestId("command-palette")).toBeHidden();

    const after = await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      return await app.getCellDisplayTextForRenderA1("A1");
    });
    expect(after).toBe("=1+1");
  });
});

