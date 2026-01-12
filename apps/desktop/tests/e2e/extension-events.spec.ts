import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop, openExtensionsPanel } from "./helpers";

const EXTENSION_ID = "formula.e2e-events";
const STORAGE_KEY = `formula.extensionHost.storage.${EXTENSION_ID}`;

async function grantSampleHelloPanelPermissions(page: Page): Promise<void> {
  await page.evaluate(() => {
    const key = "formula.extensionHost.permissions";
    const extensionId = "formula.sample-hello";
    const e2eExtensionId = "formula.e2e-events";
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

    // The built-in e2e extension activates on startup and writes event traces into extension storage.
    // Pre-grant its `storage` permission so permission prompts don't block this suite.
    existing[e2eExtensionId] = {
      ...(existing[e2eExtensionId] ?? {}),
      storage: true,
    };
    localStorage.setItem(key, JSON.stringify(existing));
  });
}

test.describe("formula.events desktop wiring", () => {
  test("emits workbook/selection/cell/sheet events into the extension host", async ({ page }) => {
    await gotoDesktop(page);
    await grantSampleHelloPanelPermissions(page);

    // Ensure the extension host is loaded (deferred until Extensions panel is opened).
    await openExtensionsPanel(page);
    await expect(page.getByTestId("panel-extensions")).toBeVisible();
    await expect(page.getByTestId("run-command-sampleHello.sumSelection")).toBeVisible({ timeout: 30_000 });

    // Ensure the e2e extension has activated and initialized its storage.
    await page.waitForFunction((storageKey) => {
      const raw = localStorage.getItem(String(storageKey));
      if (!raw) return false;
      try {
        const parsed = JSON.parse(raw);
        return parsed && typeof parsed === "object" && Object.prototype.hasOwnProperty.call(parsed, "workbookOpened");
      } catch {
        return false;
      }
    }, STORAGE_KEY);

    // Opening an extension-contributed panel should emit formula.events.onViewActivated.
    await expect(page.getByTestId("open-panel-e2eEvents.panel")).toBeVisible({ timeout: 30_000 });
    await page.getByTestId("open-panel-e2eEvents.panel").click();
    await page.waitForFunction((storageKey) => {
      const raw = localStorage.getItem(String(storageKey));
      if (!raw) return false;
      try {
        const parsed = JSON.parse(raw);
        return parsed?.viewActivated?.viewId === "e2eEvents.panel";
      } catch {
        return false;
      }
    }, STORAGE_KEY);

    // Simulate a workbook open using the extension host's stub `openWorkbook()` API. This does
    // not exercise native file IO, but it should still emit `formula.events.onWorkbookOpened`
    // to all running extensions.
    await page.evaluate(() => {
      const host = (window as any).__formulaExtensionHost;
      if (!host) throw new Error("Missing window.__formulaExtensionHost");
      host.openWorkbook("/tmp/fake.xlsx");
    });

    // Workbook open should reach formula.events.onWorkbookOpened.
    await page.waitForFunction((storageKey) => {
      const raw = localStorage.getItem(String(storageKey));
      if (!raw) return false;
      try {
        const parsed = JSON.parse(raw);
        return parsed?.workbookOpened?.workbook?.path === "/tmp/fake.xlsx";
      } catch {
        return false;
      }
    }, STORAGE_KEY);

    // Selection changes should reach formula.events.onSelectionChanged.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.selectRange({
        sheetId,
        range: { startRow: 0, startCol: 0, endRow: 0, endCol: 1 },
      });
    });

    await page.waitForFunction((storageKey) => {
      const raw = localStorage.getItem(String(storageKey));
      if (!raw) return false;
      try {
        const parsed = JSON.parse(raw);
        return parsed?.selectionChanged?.selection?.address === "A1:B1";
      } catch {
        return false;
      }
    }, STORAGE_KEY);

    // Cell edits should reach formula.events.onCellChanged.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.getDocument().setCellValue(sheetId, { row: 1, col: 1 }, 123);
    });

    await page.waitForFunction((storageKey) => {
      const raw = localStorage.getItem(String(storageKey));
      if (!raw) return false;
      try {
        const parsed = JSON.parse(raw);
        return parsed?.cellChanged?.row === 1 && parsed?.cellChanged?.col === 1 && parsed?.cellChanged?.value === 123;
      } catch {
        return false;
      }
    }, STORAGE_KEY);

    // Sheet switching should reach formula.events.onSheetActivated.
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await page.getByTestId("sheet-tab-Sheet2").click();

    await page.waitForFunction((storageKey) => {
      const raw = localStorage.getItem(String(storageKey));
      if (!raw) return false;
      try {
        const parsed = JSON.parse(raw);
        return parsed?.sheetActivated?.sheet?.id === "Sheet2";
      } catch {
        return false;
      }
    }, STORAGE_KEY);

    // Saving should emit formula.events.onBeforeSave.
    await page.evaluate(() => {
      const host = (window as any).__formulaExtensionHost;
      if (!host) throw new Error("Missing window.__formulaExtensionHost");
      host.saveWorkbook();
    });

    await page.waitForFunction((storageKey) => {
      const raw = localStorage.getItem(String(storageKey));
      if (!raw) return false;
      try {
        const parsed = JSON.parse(raw);
        return parsed?.beforeSave?.workbook?.path === "/tmp/fake.xlsx";
      } catch {
        return false;
      }
    }, STORAGE_KEY);
  });
});
