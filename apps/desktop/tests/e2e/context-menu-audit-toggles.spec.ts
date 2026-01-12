import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("grid context menu (Audit toggles)", () => {
  test("shows shortcut hints and toggles auditing mode", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    const expectedPrecedentsShortcut = process.platform === "darwin" ? "⌘[" : "Ctrl+[";
    const expectedDependentsShortcut = process.platform === "darwin" ? "⌘]" : "Ctrl+]";

    await page.waitForFunction(() => {
      const app = (window as any).__formulaApp;
      const rect = app?.getCellRectA1?.("A1");
      return rect && rect.width > 0 && rect.height > 0;
    });
    const a1 = await page.evaluate(() => (window as any).__formulaApp.getCellRectA1("A1"));

    // Toggle precedents.
    await page.locator("#grid").click({ button: "right", position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const precedents = menu.getByRole("button", { name: "Toggle Trace Precedents" });
    await expect(precedents.locator('span[aria-hidden="true"]')).toHaveText(expectedPrecedentsShortcut);
    await precedents.click();
    await waitForIdle(page);

    const afterPrecedents = await page.evaluate(() => (window as any).__formulaApp.getAuditingHighlights());
    expect(afterPrecedents.mode).toBe("precedents");

    // Toggle dependents (should become BOTH).
    await page.locator("#grid").click({ button: "right", position: { x: a1.x + a1.width / 2, y: a1.y + a1.height / 2 } });
    await expect(menu).toBeVisible();

    const dependents = menu.getByRole("button", { name: "Toggle Trace Dependents" });
    await expect(dependents.locator('span[aria-hidden="true"]')).toHaveText(expectedDependentsShortcut);
    await dependents.click();
    await waitForIdle(page);

    const afterDependents = await page.evaluate(() => (window as any).__formulaApp.getAuditingHighlights());
    expect(afterDependents.mode).toBe("both");
  });
});
