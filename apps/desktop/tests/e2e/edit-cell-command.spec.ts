import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("Edit Cell command", () => {
  test("command palette runs Edit Cell and focuses the inline editor", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    // Select A1.
    await page.click("#grid", { position: { x: 5, y: 5 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette")).toBeVisible();

    await page.getByTestId("command-palette-input").fill("Edit Cell");
    await page.keyboard.press("Enter");

    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await expect(editor).toBeFocused();
  });
});

