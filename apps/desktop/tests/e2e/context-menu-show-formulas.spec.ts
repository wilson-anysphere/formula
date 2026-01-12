import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("grid context menu (Show Formulas)", () => {
  test("shows shortcut hint and toggles to formula display", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellFormula(sheetId, "A1", "=1+1", { label: "Set formula" });
      app.refresh();
    });

    const before = await page.evaluate(() => (window as any).__formulaApp.getCellDisplayTextForRenderA1("A1"));
    expect(before).toBe("2");

    // Open context menu at A1.
    await page.locator("#grid").click({ button: "right", position: { x: 53, y: 29 } });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const showFormulas = menu.getByRole("button", { name: "Show Formulas" });
    const expectedShortcut = process.platform === "darwin" ? "⌘`" : "Ctrl+`";
    await expect(showFormulas.locator('span[aria-hidden="true"]')).toHaveText(expectedShortcut);

    await showFormulas.click();
    await waitForIdle(page);

    const after = await page.evaluate(() => (window as any).__formulaApp.getCellDisplayTextForRenderA1("A1"));
    expect(after).toBe("=1+1");
  });
});

