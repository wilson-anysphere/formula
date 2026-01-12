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

    // Toggle precedents.
    await page.locator("#grid").click({ button: "right", position: { x: 53, y: 29 } });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const precedents = menu.getByRole("button", { name: "Toggle Trace Precedents" });
    await expect(precedents.locator('span[aria-hidden="true"]')).toHaveText(expectedPrecedentsShortcut);
    await precedents.click();
    await waitForIdle(page);

    const afterPrecedents = await page.evaluate(() => (window as any).__formulaApp.getAuditingHighlights());
    expect(afterPrecedents.mode).toBe("precedents");

    // Toggle dependents (should become BOTH).
    await page.locator("#grid").click({ button: "right", position: { x: 53, y: 29 } });
    await expect(menu).toBeVisible();

    const dependents = menu.getByRole("button", { name: "Toggle Trace Dependents" });
    await expect(dependents.locator('span[aria-hidden="true"]')).toHaveText(expectedDependentsShortcut);
    await dependents.click();
    await waitForIdle(page);

    const afterDependents = await page.evaluate(() => (window as any).__formulaApp.getAuditingHighlights());
    expect(afterDependents.mode).toBe("both");
  });
});

