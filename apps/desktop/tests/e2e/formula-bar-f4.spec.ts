import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window.__formulaApp as any)?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window.__formulaApp as any).whenIdle());
}

test.describe("formula bar F4 toggles absolute/relative references", () => {
  const modes = ["legacy", "shared"] as const;

  for (const mode of modes) {
    test(`F4 toggles references inside function calls (SUM(A1)) (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Start editing in the formula bar.
      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=SUM(A1)");

      // Place caret inside the A1 token.
      await input.focus();
      await page.keyboard.press("ArrowLeft"); // just before the closing paren

      await page.keyboard.press("F4");
      await expect(input).toHaveValue("=SUM($A$1)");

      await page.keyboard.press("F4");
      await expect(input).toHaveValue("=SUM(A$1)");
    });

    test(`pressing F4 cycles absolute/relative modes for a reference token (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      // Start editing in the formula bar.
      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=A1");

      // Place caret inside the A1 token.
      await input.focus();
      await page.keyboard.press("ArrowLeft");

      await page.keyboard.press("F4");
      await expect(input).toHaveValue("=$A$1");
      expect(
        await input.evaluate((el) => ({
          start: (el as HTMLTextAreaElement).selectionStart,
          end: (el as HTMLTextAreaElement).selectionEnd,
        }))
      ).toEqual({ start: 1, end: 5 });

      await page.keyboard.press("F4");
      await expect(input).toHaveValue("=A$1");
      expect(
        await input.evaluate((el) => ({
          start: (el as HTMLTextAreaElement).selectionStart,
          end: (el as HTMLTextAreaElement).selectionEnd,
        }))
      ).toEqual({ start: 1, end: 4 });

      await page.keyboard.press("F4");
      await expect(input).toHaveValue("=$A1");
      expect(
        await input.evaluate((el) => ({
          start: (el as HTMLTextAreaElement).selectionStart,
          end: (el as HTMLTextAreaElement).selectionEnd,
        }))
      ).toEqual({ start: 1, end: 4 });

      await page.keyboard.press("F4");
      await expect(input).toHaveValue("=A1");
      expect(
        await input.evaluate((el) => ({
          start: (el as HTMLTextAreaElement).selectionStart,
          end: (el as HTMLTextAreaElement).selectionEnd,
        }))
      ).toEqual({ start: 1, end: 3 });
    });

    test(`F4 expands full-token selections to cover the toggled token (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=A1+B1");

      // Select the full first reference token (A1).
      await input.evaluate((el) => {
        const textarea = el as HTMLTextAreaElement;
        textarea.focus();
        textarea.setSelectionRange(1, 3);
      });

      await page.keyboard.press("F4");
      await expect(input).toHaveValue("=$A$1+B1");
      expect(
        await input.evaluate((el) => ({
          start: (el as HTMLTextAreaElement).selectionStart,
          end: (el as HTMLTextAreaElement).selectionEnd,
        }))
      ).toEqual({ start: 1, end: 5 });
    });

    test(`F4 is a no-op when the selection is not contained within a reference token (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();
      await input.fill("=A1+B1");

      // Select across A1 and the "+" operator (not fully contained within a reference).
      await input.evaluate((el) => {
        const textarea = el as HTMLTextAreaElement;
        textarea.focus();
        textarea.setSelectionRange(1, 4);
      });

      await page.keyboard.press("F4");
      await expect(input).toHaveValue("=A1+B1");
    });

    test(`preserves sheet qualifiers + toggles both endpoints of a range (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);
      await waitForIdle(page);

      await page.getByTestId("formula-highlight").click();
      const input = page.getByTestId("formula-input");
      await expect(input).toBeVisible();

      await input.fill("='My Sheet'!A1:B2");
      await input.focus();
      await page.keyboard.press("ArrowLeft"); // inside B2

      await page.keyboard.press("F4");
      await expect(input).toHaveValue("='My Sheet'!$A$1:$B$2");
      const value = await input.inputValue();
      expect(
        await input.evaluate((el) => ({
          start: (el as HTMLTextAreaElement).selectionStart,
          end: (el as HTMLTextAreaElement).selectionEnd,
        }))
      ).toEqual({ start: 1, end: value.length });
    });
  }
});
