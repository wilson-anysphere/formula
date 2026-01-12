import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop, openExtensionsPanel } from "./helpers";

test.describe("Extensions workbook API integration (desktop)", () => {
  test.use({ viewport: { width: 1280, height: 900 } });

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
        "workbook.manage": true,
      };

      localStorage.setItem(key, JSON.stringify(existing));
    });
  }

  test("extension can open a workbook, observe workbookOpened, and saveAs with beforeSave", async ({ page }) => {
    // On cold starts the desktop app may spend significant time doing Vite dependency optimization
    // + bootstrapping engine/extension host. Give this integration test extra runway so we don't
    // flake on slower CI hosts.
    test.setTimeout(120_000);

    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;

      let activePath: string | null = null;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            switch (cmd) {
              case "open_workbook": {
                activePath = typeof args?.path === "string" ? args.path : null;
                return {
                  path: activePath,
                  origin_path: activePath,
                  sheets: [{ id: "Sheet1", name: "Sheet1" }],
                };
              }

              case "new_workbook":
                activePath = null;
                return {
                  path: null,
                  origin_path: null,
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
                  }),
                );

                return { values, start_row: startRow, start_col: startCol };
              }

              case "save_workbook": {
                if (typeof args?.path === "string" && args.path.trim() !== "") {
                  activePath = args.path;
                }
                return null;
              }

              case "mark_saved":
              case "set_macro_ui_context":
              case "set_cell":
              case "set_range":
                return null;

              case "fire_workbook_open":
                return { ok: true, output: [], updates: [] };

              case "get_macro_security_status":
                return { has_macros: false, trust: "trusted_always" };

              case "get_workbook_theme_palette":
                return null;

              case "list_defined_names":
              case "list_tables":
                return [];

              case "stat_file":
                return { mtimeMs: 0, sizeBytes: 0 };

              default:
                // Best-effort: ignore unknown commands so unrelated UI features don't
                // break this test when new backend calls are introduced.
                return null;
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
    await grantSampleHelloPermissions(page);

    await openExtensionsPanel(page);
    await expect(page.getByTestId("panel-extensions")).toBeVisible();

    const openDemo = page.getByTestId("run-command-sampleHello.workbookOpenDemo");
    await expect(openDemo).toBeVisible({ timeout: 30_000 });
    await openDemo.dispatchEvent("click");
    await expect(page.getByTestId("input-box")).toBeVisible();
    await page.getByTestId("input-box-field").fill("/tmp/ext-open.xlsx");
    await page.getByTestId("input-box-ok").click();

    await expect(page.getByTestId("toast-root")).toContainText("Workbook opened: ext-open.xlsx (/tmp/ext-open.xlsx)");
    await expect(page.getByTestId("toast-root")).toContainText("event: /tmp/ext-open.xlsx");

    await expect(page.getByTestId("active-value")).toHaveText("Hello");

    const saveAsDemo = page.getByTestId("run-command-sampleHello.workbookSaveAsDemo");
    await expect(saveAsDemo).toBeVisible({ timeout: 30_000 });
    await saveAsDemo.dispatchEvent("click");
    await expect(page.getByTestId("input-box")).toBeVisible();
    await page.getByTestId("input-box-field").fill("/tmp/ext-save.xlsx");
    await page.getByTestId("input-box-ok").click();

    await expect(page.getByTestId("toast-root")).toContainText("Workbook beforeSave: /tmp/ext-save.xlsx");
    await expect(page.getByTestId("toast-root")).toContainText("active: ext-save.xlsx /tmp/ext-save.xlsx");
  });
});
