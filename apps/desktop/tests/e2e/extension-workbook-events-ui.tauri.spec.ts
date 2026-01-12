import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("extension workbook events (tauri UI flows)", () => {
  test("menu open + save-as trigger workbookOpened and beforeSave events for extensions", async ({ page }) => {
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
                  }),
                );
                return { values, start_row: startRow, start_col: startCol };
              }

              case "save_workbook":
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
                return { mtime_ms: 0, size_bytes: 0 };

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
        dialog: {
          open: async () => "/tmp/ui-open.xlsx",
          save: async () => "/tmp/ui-save.xlsx",
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

      // Pre-grant permissions for the ad-hoc test extension.
      try {
        const extensionId = "formula-test.wb-ui-events";
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
        };

        localStorage.setItem(key, JSON.stringify(existing));
      } catch {
        // ignore
      }
    });

    await gotoDesktop(page);
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager));
    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["menu-open"]));
    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["menu-save-as"]));

    // Load an in-memory extension that listens for workbook events.
    await page.evaluate(async () => {
      const mgr: any = (window as any).__formulaExtensionHostManager;
      if (!mgr) throw new Error("Missing window.__formulaExtensionHostManager");
      const host = mgr.host;
      if (!host) throw new Error("Missing extension host");

      const extensionId = "formula-test.wb-ui-events";
      const commandId = "wbUiEvents.get";
      const manifest = {
        name: "wb-ui-events",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Get workbook event log" }] },
        permissions: ["ui.commands"],
      };

      const code = `
        export async function activate(context) {
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          const opened = [];
          const beforeSave = [];
          context.subscriptions.push(formula.events.onWorkbookOpened((e) => opened.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(formula.events.onBeforeSave((e) => beforeSave.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
            const wb = await formula.workbook.getActiveWorkbook();
            return {
              opened,
              beforeSave,
              workbook: { name: wb?.name ?? null, path: wb?.path ?? null }
            };
          }));
        }
        export default { activate };
      `;

      const blob = new Blob([code], { type: "text/javascript" });
      const mainUrl = URL.createObjectURL(blob);

      await host.loadExtension({
        extensionId,
        extensionPath: "memory://wb-ui-events/",
        manifest,
        mainUrl,
      });
    });

    // Activate the extension so its event handlers are registered before triggering the UI flows.
    const initial = await page.evaluate(async () => {
      const mgr: any = (window as any).__formulaExtensionHostManager;
      const host = mgr.host;
      return host.executeCommand("wbUiEvents.get");
    });
    expect(initial).toEqual({ opened: [], beforeSave: [], workbook: { name: "Workbook", path: null } });

    // Trigger menu open -> open_workbook.
    await page.evaluate(() => {
      (window as any).__tauriListeners["menu-open"]({ payload: null });
    });
    await page.waitForFunction(async () => (await (window.__formulaApp as any).getCellValueA1("A1")) === "Hello");

    // Trigger save-as -> save_workbook.
    await page.evaluate(() => {
      (window as any).__tauriListeners["menu-save-as"]({ payload: null });
    });
    await page.waitForFunction(
      () => (window as any).__tauriInvokes?.some((entry: any) => entry?.cmd === "save_workbook" && entry?.args?.path) ?? false,
    );

    const events = await page.evaluate(async () => {
      const mgr: any = (window as any).__formulaExtensionHostManager;
      const host = mgr.host;
      return host.executeCommand("wbUiEvents.get");
    });

    expect(events.opened).toContain("/tmp/ui-open.xlsx");
    expect(events.beforeSave).toContain("/tmp/ui-save.xlsx");
    expect(events.workbook).toEqual({ name: "ui-save.xlsx", path: "/tmp/ui-save.xlsx" });
  });
});
