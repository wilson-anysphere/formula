import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("extension workbook lifecycle (tauri)", () => {
  test("workbook.createWorkbook/saveAs delegate to the desktop tauri flows", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            switch (cmd) {
              case "new_workbook":
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
                const values = Array.from({ length: rows }, () =>
                  Array.from({ length: cols }, () => ({ value: null, formula: null, display_value: "" })),
                );
                return { values, start_row: startRow, start_col: startCol };
              }

              case "list_defined_names":
                return [];
              case "list_tables":
                return [];
              case "get_workbook_theme_palette":
                return null;

              case "get_macro_security_status":
                return { has_macros: false, trust: "trusted_always" };
              case "set_macro_ui_context":
                return null;
              case "fire_workbook_open":
                return { ok: true, output: [], updates: [] };

              case "set_cell":
              case "set_range":
              case "save_workbook":
              case "mark_saved":
                return null;

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

      // The desktop extension host now uses a real permission prompt UI. Pre-grant
      // the permissions needed by this ad-hoc test extension so the worker can
      // activate without blocking on an interactive modal.
      try {
        const extensionId = "formula-test.wb-test";
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
      } catch {
        // ignore
      }
    });

    await gotoDesktop(page);
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager));

    const result = await page.evaluate(async () => {
      const mgr: any = (window as any).__formulaExtensionHostManager;
      if (!mgr) throw new Error("Missing window.__formulaExtensionHostManager");
      const host = mgr.host;
      if (!host) throw new Error("Missing extension host");

      const commandId = "wbTest.createAndSaveAs";
      const extensionId = "formula-test.wb-test";
      const manifest = {
        name: "wb-test",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Create workbook + Save As" }] },
        permissions: ["ui.commands", "workbook.manage"],
      };

      // Avoid module imports so the worker can keep strict import sandboxing enabled in Vite.
      // BrowserExtensionHost loads the extension API runtime into the worker already.
      const code = `
        export async function activate(context) {
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          const beforeSave = [];
          const opened = [];
          context.subscriptions.push(formula.events.onBeforeSave((e) => beforeSave.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(formula.events.onWorkbookOpened((e) => opened.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
            await formula.workbook.createWorkbook();
            await formula.workbook.saveAs(${JSON.stringify("/tmp/ext-save.xlsx")});
            const wb = await formula.workbook.getActiveWorkbook();
            return {
              workbook: { name: wb.name, path: wb.path, sheets: wb.sheets, activeSheet: wb.activeSheet },
              events: { beforeSave, opened }
            };
          }));
        }
        export default { activate };
      `;

      const blob = new Blob([code], { type: "text/javascript" });
      const mainUrl = URL.createObjectURL(blob);

      await host.loadExtension({
        extensionId,
        extensionPath: "memory://wb-test/",
        manifest,
        mainUrl,
      });

      try {
        return await host.executeCommand(commandId);
      } finally {
        URL.revokeObjectURL(mainUrl);
        await host.unloadExtension(extensionId).catch(() => {});
      }
    });

    expect(result.workbook.path).toBe("/tmp/ext-save.xlsx");
    expect(result.workbook.name).toBe("ext-save.xlsx");
    expect(Array.isArray(result.workbook.sheets)).toBe(true);
    expect(result.workbook.sheets.length).toBeGreaterThan(0);
    expect(result.workbook.activeSheet).toEqual({ id: "Sheet1", name: "Sheet1" });

    expect(result.events.beforeSave).toEqual(["/tmp/ext-save.xlsx"]);
    // createWorkbook should emit exactly one workbookOpened for the synthetic new workbook.
    expect(result.events.opened.length).toBe(1);

    const invokes = await page.evaluate(() => (window as any).__tauriInvokes);
    expect(invokes.some((entry: any) => entry?.cmd === "new_workbook")).toBe(true);
    expect(
      invokes.some((entry: any) => entry?.cmd === "save_workbook" && entry?.args?.path === "/tmp/ext-save.xlsx"),
    ).toBe(true);
  });
});
