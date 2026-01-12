import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("sheet tab context menu (tauri persistence)", () => {
  test("invokes Tauri persistence commands for visibility + tab color", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            // The desktop shell expects a minimal workbook backend during boot. Provide lightweight
            // responses for the commands that `main.ts` issues during initialization, and accept
            // the metadata persistence commands under test.
            switch (cmd) {
              case "open_workbook":
                return {
                  path: args?.path ?? null,
                  origin_path: args?.path ?? null,
                  sheets: [
                    { id: "Sheet1", name: "Sheet1" },
                    { id: "Sheet2", name: "Sheet2" },
                  ],
                };
              case "new_workbook":
                return {
                  path: null,
                  origin_path: null,
                  sheets: [
                    { id: "Sheet1", name: "Sheet1" },
                    { id: "Sheet2", name: "Sheet2" },
                  ],
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

              case "set_sheet_visibility":
              case "set_sheet_tab_color":
              case "set_cell":
              case "set_range":
              case "get_workbook_theme_palette":
              case "list_defined_names":
              case "list_tables":
              case "set_macro_ui_context":
              case "mark_saved":
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
          open: async () => null,
          save: async () => null,
        },
        window: {
          getCurrentWebviewWindow: () => ({
            hide: async () => {},
            close: async () => {},
          }),
        },
      };
    });

    await gotoDesktop(page);

    // Lazily create Sheet2 by writing a value into it. (The default desktop e2e workbook starts
    // with only Sheet1 unless we explicitly open/create a workbook via the file menu.)
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      app.getDocument().setCellValue("Sheet2", "A1", "Hello from Sheet2");
    });

    const sheet2Tab = page.getByTestId("sheet-tab-Sheet2");
    await expect(sheet2Tab).toBeVisible();

    // Hide Sheet2.
    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(() => {
      const tab = document.querySelector('[data-testid="sheet-tab-Sheet2"]') as HTMLElement | null;
      if (!tab) throw new Error("Missing Sheet2 tab");
      const rect = tab.getBoundingClientRect();
      tab.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: rect.left + 10,
          clientY: rect.top + 10,
        }),
      );
    });
    const menu = page.getByTestId("sheet-tab-context-menu");
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Hide", exact: true }).click();
    await expect(sheet2Tab).toHaveCount(0);

    // Unhide Sheet2.
    await page.evaluate(() => {
      const tab = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
      if (!tab) throw new Error("Missing Sheet1 tab");
      const rect = tab.getBoundingClientRect();
      tab.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: rect.left + 10,
          clientY: rect.top + 10,
        }),
      );
    });
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Unhide…" }).click();
    await menu.getByRole("button", { name: "Sheet2" }).click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Set a tab color on Sheet2.
    const sheet2TabVisible = page.getByTestId("sheet-tab-Sheet2");
    await page.evaluate(() => {
      const tab = document.querySelector('[data-testid="sheet-tab-Sheet2"]') as HTMLElement | null;
      if (!tab) throw new Error("Missing Sheet2 tab");
      const rect = tab.getBoundingClientRect();
      tab.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: rect.left + 10,
          clientY: rect.top + 10,
        }),
      );
    });
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Tab Color", exact: true }).click();
    await menu.getByRole("button", { name: "Red", exact: true }).click();
    await expect(sheet2TabVisible).toHaveAttribute("data-tab-color", "#ff0000");

    // Clear the tab color.
    await page.evaluate(() => {
      const tab = document.querySelector('[data-testid="sheet-tab-Sheet2"]') as HTMLElement | null;
      if (!tab) throw new Error("Missing Sheet2 tab");
      const rect = tab.getBoundingClientRect();
      tab.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          button: 2,
          clientX: rect.left + 10,
          clientY: rect.top + 10,
        }),
      );
    });
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Tab Color", exact: true }).click();
    await menu.getByRole("button", { name: "No Color" }).click();
    await expect(sheet2TabVisible).not.toHaveAttribute("data-tab-color");

    const invokes = await page.evaluate(() => (window as any).__tauriInvokes as Array<{ cmd: string; args: any }>);
    const relevant = invokes.filter((entry) => entry.cmd === "set_sheet_visibility" || entry.cmd === "set_sheet_tab_color");

    expect(relevant).toEqual([
      { cmd: "set_sheet_visibility", args: { sheet_id: "Sheet2", visibility: "hidden" } },
      { cmd: "set_sheet_visibility", args: { sheet_id: "Sheet2", visibility: "visible" } },
      { cmd: "set_sheet_tab_color", args: { sheet_id: "Sheet2", tab_color: { rgb: "FFFF0000" } } },
      { cmd: "set_sheet_tab_color", args: { sheet_id: "Sheet2", tab_color: null } },
    ]);
  });
});
