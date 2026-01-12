import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

async function grantSampleHelloPermissions(page: Page): Promise<void> {
  await page.evaluate(() => {
    const extensionId = "formula.sample-hello";
    const key = "formula.extensionHost.permissions";
    const existing = (() => {
      try {
        const raw = localStorage.getItem(key);
        return raw ? JSON.parse(raw) : {};
      } catch {
        return {};
      }
    })();

    existing[extensionId] = {
      ...(existing[extensionId] ?? {}),
      "ui.commands": true,
      "ui.panels": true,
      "cells.read": true,
      "cells.write": true,
    };

    localStorage.setItem(key, JSON.stringify(existing));
  });
}

test.describe("Extensions UI integration", () => {
  test("runs sampleHello.openPanel and renders the panel webview", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.getByTestId("open-extensions-panel").click();
    await expect(page.getByTestId("panel-extensions")).toBeVisible();

    await page.getByTestId("run-command-sampleHello.openPanel").click();

    await expect(page.getByTestId("panel-sampleHello.panel")).toBeVisible();

    const frame = page.frameLocator('iframe[data-testid="extension-webview-sampleHello.panel"]');
    await expect(frame.locator("h1")).toHaveText("Sample Hello Panel");
    await expect(
      frame.locator('meta[http-equiv="Content-Security-Policy"]'),
      "webview should inject a restrictive CSP meta tag",
    ).toHaveCount(1);
  });

  test("runs sampleHello.sumSelection via the Extensions panel and shows a toast", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

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
    await grantSampleHelloPermissions(page);

    await page.getByTestId("open-extensions-panel").click();
    await page.getByTestId("run-command-sampleHello.openPanel").click();

    await expect(page.getByTestId("panel-sampleHello.panel")).toBeVisible();
    const frame = page.frameLocator('iframe[data-testid="extension-webview-sampleHello.panel"]');
    await expect(frame.locator("h1")).toHaveText("Sample Hello Panel");

    await page.reload();
    await waitForDesktopReady(page);
    await grantSampleHelloPermissions(page);

    await expect(page.getByTestId("panel-sampleHello.panel")).toBeVisible();
    const frameAfter = page.frameLocator('iframe[data-testid="extension-webview-sampleHello.panel"]');
    await expect(frameAfter.locator("h1")).toHaveText("Sample Hello Panel");
  });

  test("executes a contributed keybinding when its when-clause matches", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

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
    await expect(page.getByTestId("run-command-sampleHello.sumSelection")).toBeVisible();

    await page.keyboard.press("Control+Shift+Y");
    await expect(page.getByTestId("toast-root")).toContainText("Sum: 10");
  });

  test("does not execute a keybinding when its when-clause fails", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, { row: 0, col: 0 }, 5);
    });

    await page.getByTestId("open-extensions-panel").click();
    await expect(page.getByTestId("run-command-sampleHello.sumSelection")).toBeVisible();

    const before = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const cell = doc.getCell(sheetId, { row: 2, col: 0 }) as any;
      return cell?.value ?? null;
    });

    // Default selection is a single cell, so `hasSelection` should be false and the keybinding should be ignored.
    await page.keyboard.press("Control+Shift+Y");

    // Give the UI a brief moment in case a command mistakenly fires.
    await page.waitForTimeout(250);

    const after = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      const cell = doc.getCell(sheetId, { row: 2, col: 0 }) as any;
      return cell?.value ?? null;
    });

    expect(after).toEqual(before);
  });

  test("executes a contributed context menu item when its when-clause matches", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPermissions(page);

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
    await expect(page.getByTestId("run-command-sampleHello.sumSelection")).toBeVisible();

    // Right-click inside the selection so the selection remains intact and `hasSelection` stays true.
    await page.locator("#grid").click({ button: "right", position: { x: 100, y: 40 } });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    const item = menu.getByRole("button", { name: "Sample Hello: Sum Selection" });
    await expect(item).toBeEnabled();
    await item.click();

    await expect(page.getByTestId("toast-root")).toContainText("Sum: 10");
  });

  test("right-clicking outside a multi-cell selection moves the active cell before showing the menu", async ({ page }) => {
    await gotoDesktop(page);

    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const sheetId = app.getCurrentSheetId();
      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      });
    });

    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("selection-range")).toHaveText("A1:B2");

    // Ensure the extensions host is running so the contributed context menu renders.
    await page.getByTestId("open-extensions-panel").click();
    await expect(page.getByTestId("panel-extensions")).toBeVisible();
    await expect(page.getByTestId("run-command-sampleHello.openPanel")).toBeVisible();

    const d4 = await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      return app.getCellRectA1("D4");
    });
    if (!d4) throw new Error("Missing D4 rect");

    // Right-click a cell outside the current selection. Excel/Sheets move the active
    // cell to the clicked cell before showing the menu so commands apply to it.
    await page.locator("#grid").click({
      button: "right",
      position: { x: d4.x + d4.width / 2, y: d4.y + d4.height / 2 },
    });

    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();
    await expect(page.getByTestId("active-cell")).toHaveText("D4");
  });
});
