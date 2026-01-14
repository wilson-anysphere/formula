import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function getA1NumberFormat(page: Page): Promise<string | null> {
  return await page.evaluate(() => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const app: any = (window as any).__formulaApp;
    if (!app) return null;
    const sheetId = app.getCurrentSheetId();
    const doc = app.getDocument();
    return doc.getCellFormat(sheetId, "A1").numberFormat ?? null;
  });
}

test.describe("Ribbon: Home → Number → More → Custom…", () => {
  test.setTimeout(120_000);

  const GRID_MODES = ["shared", "legacy"] as const;

  for (const mode of GRID_MODES) {
    test(`prompts for a custom number format and applies it (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`, { appReadyTimeoutMs: 120_000 });

      // Seed A1 with a custom number format so the prompt can pre-fill it.
      await page.evaluate(() => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const app: any = (window as any).__formulaApp;
        if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();
        doc.setCellValue(sheetId, "A1", 1234.56);
        doc.setRangeFormat(sheetId, "A1", { numberFormat: "0.00" }, { label: "Seed number format" });
        app.selectRange({ sheetId, range: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 } });
        app.focus();
      });
      await page.evaluate(() => (window as any).__formulaApp.whenIdle());

      const ribbon = page.getByTestId("ribbon-root");
      await ribbon.getByRole("tab", { name: "Home" }).click();

      // Open the Home → Number → More dropdown.
      await ribbon.locator('[data-command-id="home.number.moreFormats"]').click();

      // Click the Custom… menu item.
      await page.locator('[role="menuitem"][data-command-id="home.number.moreFormats.custom"]').click();

      // Expect an input prompt (showInputBox).
      const dialog = page.getByTestId("input-box");
      await expect(dialog).toBeVisible();
      await expect(dialog.locator(".dialog__title")).toContainText("Excel custom number format code");

      // It should be pre-filled with the current selection's (A1) effective custom code.
      const field = page.getByTestId("input-box-field");
      await expect(field).toHaveValue("0.00");

      // Apply a new custom number format.
      await field.fill("#,##0");
      await page.getByTestId("input-box-ok").click();
      await expect(dialog).toHaveCount(0);

      await expect.poll(() => getA1NumberFormat(page), { timeout: 5_000 }).toBe("#,##0");

      // Verify invalid number format codes show a toast and do not apply.
      await page.evaluate(() => {
        document.getElementById("toast-root")?.replaceChildren();
      });
      await ribbon.locator('[data-command-id="home.number.moreFormats"]').click();
      await page.locator('[role="menuitem"][data-command-id="home.number.moreFormats.custom"]').click();
      await expect(dialog).toBeVisible();
      await expect(field).toHaveValue("#,##0");

      await field.fill('"0.00'); // unbalanced quotes -> invalid
      await page.getByTestId("input-box-ok").click();
      await expect(dialog).toHaveCount(0);
      await expect(page.getByTestId("toast").first()).toHaveText("Invalid number format code.");
      await expect.poll(() => getA1NumberFormat(page), { timeout: 5_000 }).toBe("#,##0");

      // Re-open the prompt to verify it pre-fills with the newly-applied format and that "General"
      // clears the custom number format.
      await page.evaluate(() => {
        document.getElementById("toast-root")?.replaceChildren();
      });
      await ribbon.locator('[data-command-id="home.number.moreFormats"]').click();
      await page.locator('[role="menuitem"][data-command-id="home.number.moreFormats.custom"]').click();
      await expect(dialog).toBeVisible();
      await expect(field).toHaveValue("#,##0");

      await field.fill("General");
      await page.getByTestId("input-box-ok").click();
      await expect(dialog).toHaveCount(0);

      await expect.poll(() => getA1NumberFormat(page), { timeout: 5_000 }).toBeNull();
    });
  }
});
