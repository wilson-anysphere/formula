import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

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
        return false;
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
    const app = (window as any).__formulaApp;
    const sheetId = app.getCurrentSheetId();
    app.getDocument().setCellValue(sheetId, "A1", "dirty");
  });
  await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(true);
}

async function startInProgressEdit(page: import("@playwright/test").Page, value: string): Promise<void> {
  // Avoid relying on pointer clicks: canvases in the shared-grid test harness can
  // intermittently intercept pointer events. Focusing the grid is sufficient to
  // drive keyboard-based editing (F2).
  await page.locator("#grid").focus();
  await expect(page.getByTestId("active-cell")).toHaveText("A1");

  await page.keyboard.press("F2");
  const editor = page.locator("textarea.cell-editor");
  await expect(editor).toBeVisible();
  await editor.fill(value);
  await expect(editor).toHaveValue(value);
}

test.describe("non-collab: desktop unsaved-change confirmations", () => {
  test("prompts when opening another workbook and respects cancel", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);

    await makeDocumentDirty(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(true);
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    // Confirm was shown, and open was aborted.
    await expect.poll(() => page.evaluate(() => (window as any).__tauriDialogConfirmCalls.length)).toBe(1);
    // Give the queued open-workbook task a moment to run (if it would) after the dialog resolves.
    await page.waitForTimeout(50);

    const opened = await page.evaluate(
      () => ((window as any).__tauriInvokeCalls ?? []).some((c: any) => c.cmd === "open_workbook") ?? false,
    );
    expect(opened).toBe(false);

    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("dirty");
  });

  test("commits an in-progress edit before prompting to open another workbook", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);

    // Start editing without committing; the DocumentController should still be clean.
    await startInProgressEdit(page, "pending");
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(false);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    // We should prompt to discard changes; without the commit step this would have been clean
    // and the open would have proceeded without prompting.
    await expect.poll(() => page.evaluate(() => (window as any).__tauriDialogConfirmCalls.length)).toBe(1);
    await page.waitForTimeout(50);

    // Open was aborted because the user cancelled.
    const opened = await page.evaluate(
      () => ((window as any).__tauriInvokeCalls ?? []).some((c: any) => c.cmd === "open_workbook") ?? false,
    );
    expect(opened).toBe(false);

    // The in-progress edit should have been committed (and therefore made the doc dirty).
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeHidden();
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"))).toBe("pending");
    await expect.poll(() => page.evaluate(() => (window as any).__formulaApp.getDocument().isDirty)).toBe(true);
  });

  test("prompts when quitting and respects cancel", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);

    await makeDocumentDirty(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["tray-quit"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["tray-quit"]();
    });

    // Confirm was shown, and quit was aborted.
    await expect.poll(() => page.evaluate(() => (window as any).__tauriDialogConfirmCalls.length)).toBe(1);
    await page.waitForTimeout(50);

    const quitCalled = await page.evaluate(
      () => ((window as any).__tauriInvokeCalls ?? []).some((c: any) => c.cmd === "quit_app") ?? false,
    );
    expect(quitCalled).toBe(false);
  });
});
