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

test.describe("status bar selection stats", () => {
  const GRID_MODES = ["legacy", "shared"] as const;

  for (const mode of GRID_MODES) {
    test(`Count uses non-empty cells (Sum/Avg numeric only) (${mode})`, async ({ page }) => {
      await gotoDesktop(page, `/?grid=${mode}`);

      await page.evaluate(() => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const app: any = window.__formulaApp as any;
        if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");

        const sheetId = app.getCurrentSheetId();
        const doc = app.getDocument();

        doc.setCellValue(sheetId, "A1", 1);
        doc.setCellValue(sheetId, "A2", "hello");

        app.selectRange({
          sheetId,
          range: { startRow: 0, endRow: 1, startCol: 0, endCol: 0 },
        });
      });
      await waitForIdle(page);

      await expect(page.getByTestId("selection-sum")).toHaveText("Sum: 1");
      await expect(page.getByTestId("selection-avg")).toHaveText("Avg: 1");
      await expect(page.getByTestId("selection-count")).toHaveText("Count: 2");
    });
  }
});
