import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("formula reference colors", () => {
  test("colors each reference in the formula bar and renders matching grid overlays", async ({ page }) => {
    await gotoDesktop(page);

    // Select C1 (avoid overlapping the referenced cells).
    await page.click("#grid", { position: { x: 260, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("C1");

    await page.getByTestId("formula-highlight").click();
    const input = page.getByTestId("formula-input");
    await expect(input).toBeVisible();

    await input.fill("=A1+B1");

    const refs = page.locator('[data-testid="formula-highlight"] span[data-kind="reference"]');
    await expect(refs).toHaveCount(2);

    const colors = await refs.evaluateAll((els) => els.map((el) => getComputedStyle(el).color));
    expect(colors[0]).not.toBe(colors[1]);

    // Excel UX: clicking inside a reference selects that reference span.
    await input.evaluate((el) => {
      const inputEl = el as HTMLInputElement;
      inputEl.focus();
      inputEl.setSelectionRange(2, 2); // inside "A1"
      inputEl.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });
    await expect(input).toHaveJSProperty("selectionStart", 1);
    await expect(input).toHaveJSProperty("selectionEnd", 3);

    // Clicking again on the same reference toggles back to a caret for manual edits.
    await input.evaluate((el) => {
      const inputEl = el as HTMLInputElement;
      inputEl.focus();
      inputEl.setSelectionRange(2, 2);
      inputEl.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    });
    await expect(input).toHaveJSProperty("selectionStart", 2);
    await expect(input).toHaveJSProperty("selectionEnd", 2);

    await page.waitForFunction(() => (window.__formulaApp as any).getReferenceHighlightCount() === 2);
  });
});
