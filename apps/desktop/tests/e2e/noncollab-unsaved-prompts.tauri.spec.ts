import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

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

async function makeDocumentDirty(page: import("@playwright/test").Page): Promise<void> {
  await page.evaluate(() => {
    const app = (window as any).__formulaApp;
    const sheetId = app.getCurrentSheetId();
    app.getDocument().setCellValue(sheetId, "A1", "dirty");
  });
  await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(true);
}

test.describe("non-collab: desktop unsaved-change confirmations", () => {
  test("prompts when opening another workbook and respects cancel", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);

    await makeDocumentDirty(page);

    let confirmDialogs = 0;
    page.on("dialog", async (dialog) => {
      if (dialog.type() === "confirm") {
        confirmDialogs += 1;
        await dialog.dismiss();
        return;
      }
      await dialog.accept();
    });

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    // Confirm was shown, and open was aborted.
    await expect.poll(() => confirmDialogs).toBe(1);
    // Give the queued open-workbook task a moment to run (if it would) after the dialog resolves.
    await page.waitForTimeout(50);

    const opened = await page.evaluate(
      () => ((window as any).__tauriInvokeCalls ?? []).some((c: any) => c.cmd === "open_workbook") ?? false,
    );
    expect(opened).toBe(false);

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("dirty");
  });

  test("prompts when quitting and respects cancel", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);

    await makeDocumentDirty(page);

    let confirmDialogs = 0;
    page.on("dialog", async (dialog) => {
      if (dialog.type() === "confirm") {
        confirmDialogs += 1;
        await dialog.dismiss();
        return;
      }
      await dialog.accept();
    });

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["tray-quit"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["tray-quit"]();
    });

    // Confirm was shown, and quit was aborted.
    await expect.poll(() => confirmDialogs).toBe(1);
    await page.waitForTimeout(50);

    const quitCalled = await page.evaluate(
      () => ((window as any).__tauriInvokeCalls ?? []).some((c: any) => c.cmd === "quit_app") ?? false,
    );
    expect(quitCalled).toBe(false);
  });
});
