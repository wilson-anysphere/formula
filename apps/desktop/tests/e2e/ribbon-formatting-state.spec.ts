import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  // Vite may occasionally trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-wait.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => Boolean((window.__formulaApp as any)?.whenIdle), null, { timeout: 10_000 });
      await page.evaluate(() => (window.__formulaApp as any).whenIdle());
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      throw err;
    }
  }
}

test.describe("ribbon formatting state", () => {
  test("Bold/Align/Number format reflect current selection formatting", async ({ page }) => {
    await gotoDesktop(page);
    await waitForIdle(page);

    // Select A1. (Avoid the top-left corner header, which selects all.)
    await page.click("#grid", { position: { x: 60, y: 40 } });
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const ribbon = page.getByTestId("ribbon-root");
    await expect(ribbon).toBeVisible();

    // The desktop shell currently opens on the View tab; switch to Home so
    // formatting controls are visible.
    const homeTab = ribbon.getByRole("tab", { name: "Home" });
    await homeTab.click();
    await expect(homeTab).toHaveAttribute("aria-selected", "true");

    const bold = ribbon.locator('button[data-command-id="format.toggleBold"]');
    await expect(bold).toBeVisible();
    await expect(bold).toHaveAttribute("aria-pressed", "false");

    // Toggle Bold on for A1.
    await bold.click();
    await expect(bold).toHaveAttribute("aria-pressed", "true");

    // Move to B1 and ensure Bold shows unpressed.
    await page.keyboard.press("ArrowRight");
    await expect(page.getByTestId("active-cell")).toHaveText("B1");
    await expect(bold).toHaveAttribute("aria-pressed", "false");

    // Back to A1 and ensure Bold shows pressed again.
    await page.keyboard.press("ArrowLeft");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(bold).toHaveAttribute("aria-pressed", "true");

    // Alignment buttons (not toggles) should still show a pressed style.
    const alignLeft = ribbon.locator('button[data-command-id="format.alignLeft"]');
    const alignCenter = ribbon.locator('button[data-command-id="format.alignCenter"]');
    await expect(alignLeft).toHaveClass(/is-pressed/);
    await expect(alignCenter).not.toHaveClass(/is-pressed/);

    await alignCenter.click();
    await expect(alignCenter).toHaveClass(/is-pressed/);
    await expect(alignLeft).not.toHaveClass(/is-pressed/);

    // Number format label should update after applying a preset.
    const numberFormatDropdown = ribbon.locator('button[data-command-id="home.number.numberFormat"]');
    const numberFormatLabel = numberFormatDropdown.locator(".ribbon-button__label");
    await expect(numberFormatLabel).toHaveText("General");

    const percent = ribbon.locator('button[data-command-id="format.numberFormat.percent"]');
    await percent.scrollIntoViewIfNeeded();
    await percent.click();

    await expect(numberFormatLabel).toHaveText("Percent");
  });
});
