import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";
import { installCollabSessionStub } from "./collabSessionStub";

function installTauriStubForTests() {
  const listeners: Record<string, any> = {};
  (window as any).__tauriListeners = listeners;
  (window as any).__tauriInvokeCalls = [];

  const pushCall = (cmd: string, args: any) => {
    (window as any).__tauriInvokeCalls.push({ cmd, args });
  };

  (window as any).__TAURI__ = {
    core: {
      invoke: async (cmd: string, args: any) => {
        pushCall(cmd, args);
        switch (cmd) {
          case "open_workbook":
            return {
              path: args?.path ?? null,
              origin_path: args?.path ?? null,
              sheets: [{ id: "Sheet1", name: "Sheet1" }],
            };

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

          case "stat_file":
            return { mtime_ms: 0, size_bytes: 0 };

          case "get_macro_security_status":
            return { has_macros: false, trust: "trusted_always" };

          case "fire_workbook_open":
          case "fire_workbook_before_close":
            return { ok: true, updates: [] };

          case "get_workbook_theme_palette":
          case "list_defined_names":
          case "list_tables":
          case "set_macro_ui_context":
          case "mark_saved":
          case "set_cell":
          case "set_range":
          case "save_workbook":
          case "set_tray_status":
          case "quit_app":
          case "restart_app":
            return null;

          default:
            // Important: do *not* allow sheet operations like `delete_sheet` here. In collab mode,
            // sheet metadata is persisted to Yjs and should not hit the local Tauri workbook backend.
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
}

async function grantExtensionPermissions(page: Page, extensionId: string, permissions: string[]): Promise<void> {
  await page.addInitScript(
    ({ extensionId, permissions }) => {
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
        ...Object.fromEntries(permissions.map((perm) => [perm, true])),
      };

      localStorage.setItem(key, JSON.stringify(existing));
    },
    { extensionId, permissions },
  );
}

test.describe("collab: extension sheet deletion (tauri)", () => {
  test("formula.sheets.deleteSheet does not call the Tauri workbook backend in collab mode", async ({ page }) => {
    await grantExtensionPermissions(page, "formula-test.sheet-delete-collab-ext", ["ui.commands", "sheets.manage"]);
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);

    await installCollabSessionStub(page);
    // Seed a second sheet so the extension delete call does not trip the "cannot delete the last
    // sheet" guard. Also force a `syncSheetUi()` pass so the desktop sheet metadata store sees it.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = (window as any).__formulaApp;
      const session = app?.getCollabSession?.();
      if (!session) throw new Error("Missing collab session");

      const existing = session.sheets?.toArray?.() ?? [];
      const hasSheet2 = Array.isArray(existing) && existing.some((s: any) => String(s?.id ?? s?.get?.("id") ?? "") === "Sheet2");
      if (!hasSheet2) {
        session.transactLocal(() => {
          session.sheets.insert(1, [{ id: "Sheet2", name: "Sheet2" }]);
        });
      }

      try {
        // Ensure the DocumentController sheet exists (it creates sheets lazily).
        app.getDocument().getCell("Sheet2", { row: 0, col: 0 });
      } catch {
        // ignore
      }

      // `activateSheet()` is wrapped in main.ts to call `syncSheetUi()`, which rebuilds the
      // workbook sheet store from `session.sheets`.
      app.activateSheet(app.getCurrentSheetId());
    });

    await page.evaluate(async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const manager: any = (window as any).__formulaExtensionHostManager;
      if (!manager) throw new Error("Missing window.__formulaExtensionHostManager (desktop e2e harness)");

      if (!manager.ready) {
        await manager.loadBuiltInExtensions();
      }

      const commandId = "sheetExt.deleteSheetCollab";
      const manifest = {
        name: "sheet-delete-collab-ext",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Delete sheet (collab)" }] },
        permissions: ["ui.commands", "sheets.manage"],
      };

      const code = `
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          if (!formula) throw new Error("Missing formula extension API runtime");
          export async function activate(context) {
            context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
            await formula.sheets.deleteSheet("Sheet2");
            return true;
          }));
        }
        export default { activate };
      `;

      const blob = new Blob([code], { type: "text/javascript" });
      const mainUrl = URL.createObjectURL(blob);
      const extensionId = `${manifest.publisher}.${manifest.name}`;

      try {
        await manager.host.loadExtension({
          extensionId,
          extensionPath: "memory://sheet-delete-collab-ext/",
          manifest,
          mainUrl,
        });
        await manager.host.executeCommand(commandId);
      } finally {
        try {
          await manager.host.unloadExtension(extensionId);
        } catch {
          // ignore cleanup failures
        }
        URL.revokeObjectURL(mainUrl);
      }
    });

    await expect
      .poll(() => page.evaluate(() => (window as any).__tauriInvokeCalls?.some((c: any) => c.cmd === "delete_sheet") ?? false))
      .toBe(false);
  });
});
