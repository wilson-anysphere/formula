import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("in-cell editor F4 toggles absolute/relative references", () => {
  const modes = ["legacy", "shared"] as const;

  for (const mode of modes) {
    test(`pressing F4 toggles =A1 to =$A$1 while editing a cell (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Select B1 (avoid self-reference in A1).
      await page.click("#grid", { position: { x: 160, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("B1");

      // Start in-cell editing.
      await page.keyboard.press("F2");
      const editor = page.locator("textarea.cell-editor");
      await expect(editor).toBeVisible();
      await editor.fill("=A1");

      // Place caret inside A1.
      await editor.focus();
      await page.keyboard.press("ArrowLeft");

      await page.keyboard.press("F4");
      await expect(editor).toHaveValue("=$A$1");
      await expect(editor).toHaveJSProperty("selectionStart", 1);
      await expect(editor).toHaveJSProperty("selectionEnd", 5);
    });
  }
});

