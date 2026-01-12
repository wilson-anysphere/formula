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
    await page.addInitScript(() => {
      // Avoid permission modal flakiness in this suite; other e2e tests cover the
      // explicit permission prompt UI.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (globalThis as any).__formulaPermissionPrompt = async () => true;

      // Desktop falls back to `window.confirm` / `window.alert` when the Tauri dialog plugin is
      // unavailable. Override them so open/save flows don't hang or get cancelled under Playwright.
      (window as any).confirm = () => true;
      (window as any).alert = () => {};

      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokeCalls = [];

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            (window as any).__tauriInvokeCalls.push({ cmd, args });
            switch (cmd) {
              case "open_workbook":
                return {
                  path: args?.path ?? null,
                  origin_path: args?.path ?? null,
                  sheets: [{ id: "Sheet1", name: "Sheet1" }],
                };

              case "get_sheet_used_range":
                return { start_row: 0, end_row: 0, start_col: 0, end_col: 0 };

              case "get_range": {
                const startRow = Number(args?.start_row ?? 0);
                const endRow = Number(args?.end_row ?? startRow);
                const startCol = Number(args?.start_col ?? 0);
                const endCol = Number(args?.end_col ?? startCol);
                const rows = Math.max(0, endRow - startRow + 1);
                const cols = Math.max(0, endCol - startCol + 1);

                const values = Array.from({ length: rows }, (_v, r) =>
                  Array.from({ length: cols }, (_w, c) => {
                    const row = startRow + r;
                    const col = startCol + c;
                    if (row === 0 && col === 0) {
                      return { value: "Hello", formula: null, display_value: "Hello" };
                    }
                    return { value: null, formula: null, display_value: "" };
                  })
                );

                return { values, start_row: startRow, start_col: startCol };
              }

              case "stat_file":
                return { mtime_ms: 0, size_bytes: 0 };

              case "power_query_state_get":
                return "";

              case "get_macro_security_status":
                return { has_macros: false, trust: "trusted_always" };

              case "set_macro_ui_context":
                return null;

              case "fire_workbook_open":
                return { ok: true, output: [], updates: [] };

              case "fire_selection_change":
              case "fire_worksheet_change":
                return { ok: true, updates: [] };

              case "set_cell":
              case "set_range":
              case "save_workbook":
                return null;

              default:
                throw new Error(`Unexpected invoke: ${cmd} ${JSON.stringify(args)}`);
            }
          },
        },
        event: {
          listen: async (name: string, handler: any) => {
            listeners[name] = handler;
            return () => {
              delete listeners[name];
            };
          },
          emit: async () => {},
        },
        window: {
          getCurrentWebviewWindow: () => ({
            hide: async () => {
              (window as any).__tauriHidden = true;
            },
            close: async () => {
              (window as any).__tauriClosed = true;
            },
          }),
        },
      };
    });
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

    // Open a workbook via the Tauri file-dropped hook.
    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await expect
      .poll(() => page.evaluate(() => (window as any).__tauriInvokeCalls?.some?.((c: any) => c?.cmd === "open_workbook")), {
        timeout: 30_000,
      })
      .toBe(true);

    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1")), { timeout: 30_000 })
      .toBe("Hello");
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
        return (
          parsed?.selectionChanged?.selection?.address === "A1:B1" &&
          parsed?.selectionChanged?.sheetId === "Sheet1"
        );
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
