import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("Extensions UI integration", () => {
  test("runs sampleHello.openPanel and renders the panel webview", async ({ page }) => {
    await gotoDesktop(page);

    await page.getByTestId("open-extensions-panel").click();
    await expect(page.getByTestId("panel-extensions")).toBeVisible();

    await page.getByTestId("run-command-sampleHello.openPanel").click();

    await expect(page.getByTestId("panel-sampleHello.panel")).toBeVisible();

    const frame = page.frameLocator('iframe[data-testid="extension-webview-sampleHello.panel"]');
    await expect(frame.locator("h1")).toHaveText("Sample Hello Panel");
  });

  test("runs sampleHello.sumSelection via the Extensions panel and shows a toast", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();

      doc.setCellValue(sheetId, { row: 0, col: 0 }, 1);
      doc.setCellValue(sheetId, { row: 0, col: 1 }, 2);
      doc.setCellValue(sheetId, { row: 1, col: 0 }, 3);
      doc.setCellValue(sheetId, { row: 1, col: 1 }, 4);

      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await page.getByTestId("open-extensions-panel").click();
    await page.getByTestId("run-command-sampleHello.sumSelection").click();

    await expect(page.getByTestId("toast-root")).toContainText("Sum: 10");
  });

  test("persists an extension panel in the layout and re-activates it after reload", async ({ page }) => {
    await gotoDesktop(page);

    await page.getByTestId("open-extensions-panel").click();
    await page.getByTestId("run-command-sampleHello.openPanel").click();

    await expect(page.getByTestId("panel-sampleHello.panel")).toBeVisible();
    const frame = page.frameLocator('iframe[data-testid="extension-webview-sampleHello.panel"]');
    await expect(frame.locator("h1")).toHaveText("Sample Hello Panel");

    await page.reload();
    await waitForDesktopReady(page);

    await expect(page.getByTestId("panel-sampleHello.panel")).toBeVisible();
    const frameAfter = page.frameLocator('iframe[data-testid="extension-webview-sampleHello.panel"]');
    await expect(frameAfter.locator("h1")).toHaveText("Sample Hello Panel");
  });
});
