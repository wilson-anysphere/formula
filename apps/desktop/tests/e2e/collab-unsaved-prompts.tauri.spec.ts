import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";
import { installCollabSessionStub } from "./collabSessionStub";

function installTauriStubForTests() {
  const listeners: Record<string, any> = {};
  (window as any).__tauriListeners = listeners;
  (window as any).__tauriInvokeCalls = [];
  (window as any).__tauriDialogConfirmCalls = [];
  (window as any).__tauriDialogAlertCalls = [];

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
    dialog: {
      confirm: async (message: string) => {
        (window as any).__tauriDialogConfirmCalls.push(message);
        return true;
      },
      message: async (message: string) => {
        (window as any).__tauriDialogAlertCalls.push(message);
      },
      alert: async (message: string) => {
        (window as any).__tauriDialogAlertCalls.push(message);
      },
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
    const app = window.__formulaApp as any;
    const sheetId = app.getCurrentSheetId();
    app.getDocument().setCellValue(sheetId, "A1", "dirty");
  });
  await expect.poll(() => page.evaluate(() => (window.__formulaApp as any).getDocument().isDirty)).toBe(true);
}

test.describe("collab: desktop unsaved-change confirmations", () => {
  test("does not prompt when opening another workbook in collab mode", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);

    await installCollabSessionStub(page);
    await makeDocumentDirty(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await page.waitForFunction(async () => (await (window.__formulaApp as any).getCellValueA1("A1")) === "Hello");

    expect(await page.evaluate(() => (window as any).__tauriDialogConfirmCalls.length)).toBe(0);
  });

  test("does not prompt when creating a new workbook in collab mode", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);

    await installCollabSessionStub(page);
    await makeDocumentDirty(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["tray-new"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["tray-new"]();
    });

    await expect
      .poll(
        () =>
          page.evaluate(() => (window as any).__tauriInvokeCalls?.some((c: any) => c.cmd === "new_workbook") ?? false),
        { timeout: 10_000 },
      )
      .toBe(true);

    expect(await page.evaluate(() => (window as any).__tauriDialogConfirmCalls.length)).toBe(0);
  });

  test("does not prompt when closing/quitting in collab mode", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);

    await installCollabSessionStub(page);
    await makeDocumentDirty(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["menu-close-window"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["menu-close-window"]();
    });
    await expect.poll(() => page.evaluate(() => Boolean((window as any).__tauriHidden))).toBe(true);

    // Trigger quit via the same path as tray/menu/keyboard shortcuts.
    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["tray-quit"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["tray-quit"]();
    });
    await expect
      .poll(
        () =>
          page.evaluate(() => (window as any).__tauriInvokeCalls?.some((c: any) => c.cmd === "quit_app") ?? false),
        { timeout: 10_000 },
      )
      .toBe(true);

    expect(await page.evaluate(() => (window as any).__tauriDialogConfirmCalls.length)).toBe(0);
  });
});
