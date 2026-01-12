import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window.__formulaApp as any)?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window.__formulaApp as any).whenIdle());
}

test.describe("grid context menu (Show Formulas)", () => {
  test("shows shortcut hint and toggles to formula display", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    await page.evaluate(() => {
      const app = window.__formulaApp as any;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellFormula(sheetId, "A1", "=1+1", { label: "Set formula" });
      app.refresh();
    });

    const before = await page.evaluate(() => (window.__formulaApp as any).getCellDisplayTextForRenderA1("A1"));
    expect(before).toBe("2");

    // Open context menu at A1.
    await page.waitForFunction(() => {
      const app = window.__formulaApp as any;
      const rect = app?.getCellRectA1?.("A1");
      return rect && rect.width > 0 && rect.height > 0;
    });
    const a1 = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("A1"));
    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(
      ({ x, y }) => {
        const grid = document.getElementById("grid");
        if (!grid) throw new Error("Missing #grid");
        const rect = grid.getBoundingClientRect();
        grid.dispatchEvent(
          new MouseEvent("contextmenu", {
            bubbles: true,
            cancelable: true,
            button: 2,
            clientX: rect.left + x,
            clientY: rect.top + y,
          }),
        );
      },
      { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 },
    );
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const showFormulas = menu.getByRole("button", { name: "Show Formulas" });
    const expectedShortcut = process.platform === "darwin" ? "⌘`" : "Ctrl+`";
    await expect(showFormulas.locator('span[aria-hidden="true"]')).toHaveText(expectedShortcut);

    await showFormulas.click();
    await waitForIdle(page);

    const after = await page.evaluate(() => (window.__formulaApp as any).getCellDisplayTextForRenderA1("A1"));
    expect(after).toBe("=1+1");
  });
});
