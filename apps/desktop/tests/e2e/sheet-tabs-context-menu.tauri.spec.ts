import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop, openSheetTabContextMenu } from "./helpers";

async function openWorkbookFromFileDrop(page: Page, path: string = "/tmp/fake.xlsx"): Promise<void> {
  // The desktop shell only starts workbookSync (and therefore sheet-metadata persistence)
  // once a workbook is opened via the Tauri backend. Simulate a native file drop so the
  // app runs `open_workbook` and installs the sync bridge.
  await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
  await page.evaluate((workbookPath) => {
    (window as any).__tauriListeners["file-dropped"]({ payload: [workbookPath] });
  }, path);
  // Wait for the workbook snapshot to be applied (our stub returns "Hello" in A1).
  await page.waitForFunction(async () => (await (window as any).__formulaApp.getCellValueA1("A1")) === "Hello");
}

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
                    { id: "Sheet1", name: "Sheet1", visibility: "visible" },
                    { id: "Sheet2", name: "Sheet2", visibility: "visible" },
                  ],
                };
              case "new_workbook":
                return {
                  path: null,
                  origin_path: null,
                  sheets: [
                    { id: "Sheet1", name: "Sheet1", visibility: "visible" },
                    { id: "Sheet2", name: "Sheet2", visibility: "visible" },
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
    await openWorkbookFromFileDrop(page);

    const sheet2Tab = page.getByTestId("sheet-tab-Sheet2");
    await expect(sheet2Tab).toBeVisible();

    // Hide Sheet2.
    const menu = await openSheetTabContextMenu(page, "Sheet2");
    await menu.getByRole("button", { name: "Hide", exact: true }).click();
    await expect(sheet2Tab).toHaveCount(0);

    // Unhide Sheet2.
    await openSheetTabContextMenu(page, "Sheet1");
    await menu.getByRole("button", { name: "Unhide…", exact: true }).click();
    await menu.getByRole("button", { name: "Sheet2" }).click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Set a tab color on Sheet2.
    const sheet2TabVisible = page.getByTestId("sheet-tab-Sheet2");
    await openSheetTabContextMenu(page, "Sheet2");
    await menu.getByRole("button", { name: "Tab Color", exact: true }).click();
    await menu.getByRole("button", { name: "Red", exact: true }).click();
    await expect(sheet2TabVisible).toHaveAttribute("data-tab-color", "#ff0000");

    // Clear the tab color.
    await openSheetTabContextMenu(page, "Sheet2");
    await menu.getByRole("button", { name: "Tab Color", exact: true }).click();
    await menu.getByRole("button", { name: "No Color", exact: true }).click();
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
