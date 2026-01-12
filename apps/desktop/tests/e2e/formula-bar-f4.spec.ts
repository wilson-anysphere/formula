import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 10_000 });
  await page.evaluate(() => (window as any).__formulaApp.whenIdle());
}

test.describe("formula bar F4 toggles absolute/relative references", () => {
  const modes = ["legacy", "shared"] as const;

  for (const mode of modes) {
    test(`pressing F4 toggles =A1 to =$A$1 (${mode})`, async ({ page }) => {
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
    });
  }
});

