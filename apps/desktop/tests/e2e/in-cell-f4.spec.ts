import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("in-cell editor F4 toggles absolute/relative references", () => {
  const modes = ["legacy", "shared"] as const;

  for (const mode of modes) {
    test(`pressing F4 cycles absolute/relative modes while editing a cell (${mode})`, async ({ page }) => {
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

      await page.keyboard.press("F4");
      await expect(editor).toHaveValue("=A$1");
      await expect(editor).toHaveJSProperty("selectionStart", 1);
      await expect(editor).toHaveJSProperty("selectionEnd", 4);

      await page.keyboard.press("F4");
      await expect(editor).toHaveValue("=$A1");
      await expect(editor).toHaveJSProperty("selectionStart", 1);
      await expect(editor).toHaveJSProperty("selectionEnd", 4);

      await page.keyboard.press("F4");
      await expect(editor).toHaveValue("=A1");
      await expect(editor).toHaveJSProperty("selectionStart", 1);
      await expect(editor).toHaveJSProperty("selectionEnd", 3);
    });

    test(`F4 expands full-token selections to cover the toggled token (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Select B1 (avoid self-reference in A1).
      await page.click("#grid", { position: { x: 160, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("B1");

      await page.keyboard.press("F2");
      const editor = page.locator("textarea.cell-editor");
      await expect(editor).toBeVisible();
      await editor.fill("=A1");

      // Select full reference token.
      await editor.evaluate((el) => {
        const textarea = el as HTMLTextAreaElement;
        textarea.focus();
        textarea.setSelectionRange(1, 3);
      });

      await page.keyboard.press("F4");
      await expect(editor).toHaveValue("=$A$1");
      await expect(editor).toHaveJSProperty("selectionStart", 1);
      await expect(editor).toHaveJSProperty("selectionEnd", 5);
    });

    test(`F4 is a no-op when the selection is not contained within a reference token (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Select B1 (avoid self-reference in A1).
      await page.click("#grid", { position: { x: 160, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("B1");

      await page.keyboard.press("F2");
      const editor = page.locator("textarea.cell-editor");
      await expect(editor).toBeVisible();
      await editor.fill("=A1+B1");

      // Selection spans A1 and the "+" operator (not fully contained within a reference).
      await editor.evaluate((el) => {
        const textarea = el as HTMLTextAreaElement;
        textarea.focus();
        textarea.setSelectionRange(1, 4);
      });

      await page.keyboard.press("F4");
      await expect(editor).toHaveValue("=A1+B1");
    });

    test(`F4 toggles sheet-qualified range references (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Select B1 (avoid self-reference in A1).
      await page.click("#grid", { position: { x: 160, y: 40 } });
      await expect(page.getByTestId("active-cell")).toHaveText("B1");

      await page.keyboard.press("F2");
      const editor = page.locator("textarea.cell-editor");
      await expect(editor).toBeVisible();
      await editor.fill("='My Sheet'!A1:B2");
      await editor.focus();
      await page.keyboard.press("ArrowLeft"); // inside B2

      await page.keyboard.press("F4");
      await expect(editor).toHaveValue("='My Sheet'!$A$1:$B$2");
      const value = await editor.inputValue();
      await expect(editor).toHaveJSProperty("selectionStart", 1);
      await expect(editor).toHaveJSProperty("selectionEnd", value.length);
    });
  }
});
