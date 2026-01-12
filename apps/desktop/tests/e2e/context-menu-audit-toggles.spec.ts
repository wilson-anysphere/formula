import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window.__formulaApp as any)?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window.__formulaApp as any).whenIdle());
}

test.describe("grid context menu (Audit toggles)", () => {
  test("shows shortcut hints and toggles auditing mode", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    const expectedPrecedentsShortcut = process.platform === "darwin" ? "⌘[" : "Ctrl+[";
    const expectedDependentsShortcut = process.platform === "darwin" ? "⌘]" : "Ctrl+]";

    await page.waitForFunction(() => {
      const app = window.__formulaApp as any;
      const rect = app?.getCellRectA1?.("A1");
      return rect && rect.width > 0 && rect.height > 0;
    });
    const a1 = await page.evaluate(() => (window.__formulaApp as any).getCellRectA1("A1"));

    // Toggle precedents.
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

    const precedents = menu.getByRole("button", { name: "Toggle Trace Precedents" });
    await expect(precedents.locator('span[aria-hidden="true"]')).toHaveText(expectedPrecedentsShortcut);
    await precedents.click();
    await waitForIdle(page);

    const afterPrecedents = await page.evaluate(() => (window.__formulaApp as any).getAuditingHighlights());
    expect(afterPrecedents.mode).toBe("precedents");

    // Toggle dependents (should become BOTH).
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
    await expect(menu).toBeVisible();

    const dependents = menu.getByRole("button", { name: "Toggle Trace Dependents" });
    await expect(dependents.locator('span[aria-hidden="true"]')).toHaveText(expectedDependentsShortcut);
    await dependents.click();
    await waitForIdle(page);

    const afterDependents = await page.evaluate(() => (window.__formulaApp as any).getAuditingHighlights());
    expect(afterDependents.mode).toBe("both");
  });
});
