import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

function captureAppErrors(page: import("@playwright/test").Page): string[] {
  const errors: string[] = [];

  page.on("console", (msg) => {
    if (msg.type() !== "error") return;
    errors.push(msg.text());
  });

  page.on("pageerror", (err) => {
    errors.push(err.message ?? String(err));
  });

  return errors;
}

async function waitForIdle(page: import("@playwright/test").Page, capturedErrors: string[]): Promise<void> {
  // Vite may occasionally trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-wait.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => Boolean((window as any).__formulaApp?.whenIdle), null, { timeout: 30_000 });
      await page.evaluate(() => (window as any).__formulaApp.whenIdle());
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      const errorText = capturedErrors.length > 0 ? capturedErrors.join("\n") : "(no console errors captured)";
      throw new Error(`Spreadsheet app failed to initialize.\n\nCaptured browser errors:\n${errorText}\n\nOriginal error:\n${String(err)}`);
    }
  }
}

test("shared grid uses Excel-scale limits and can activate bottom-right cell", async ({ page }) => {
  const capturedErrors = captureAppErrors(page);
  await gotoDesktop(page, "/?grid=shared");
  await waitForIdle(page, capturedErrors);

  const limits = await page.evaluate(() => (window as any).__formulaApp.limits as { maxRows: number; maxCols: number });
  expect(limits.maxRows).toBe(1_048_576);
  expect(limits.maxCols).toBe(16_384);

  await page.evaluate(() => {
    const app = (window as any).__formulaApp as any;
    const { maxRows, maxCols } = app.limits as { maxRows: number; maxCols: number };
    app.focus();
    app.activateCell({ row: maxRows - 1, col: maxCols - 1 });
  });

  await expect(page.getByTestId("active-cell")).toHaveText("XFD1048576");

  const scroll = await page.evaluate(() => (window as any).__formulaApp.getScroll());
  expect(scroll.x).toBeGreaterThan(0);
  expect(scroll.y).toBeGreaterThan(0);

  // Shared-grid mode should not build the legacy visibility caches (which are O(maxRows/maxCols)).
  const cacheSizes = await page.evaluate(() => {
    const app = (window as any).__formulaApp as any;
    const rows = Array.isArray(app.rowIndexByVisual) ? app.rowIndexByVisual.length : null;
    const cols = Array.isArray(app.colIndexByVisual) ? app.colIndexByVisual.length : null;
    return { rows, cols };
  });

  if (cacheSizes.rows != null) expect(cacheSizes.rows).toBeLessThan(10_000);
  if (cacheSizes.cols != null) expect(cacheSizes.cols).toBeLessThan(10_000);
});

