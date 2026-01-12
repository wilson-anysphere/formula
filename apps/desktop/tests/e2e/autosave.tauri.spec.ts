import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("AutoSave (tauri)", () => {
  test("automatically saves the workbook after a committed edit (debounced)", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;
      (window as any).__tauriDialogOpenCalls = 0;

      // Avoid modal prompts blocking the test.
      window.confirm = () => true;

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
                // Best-effort: ignore unrelated invocations so new backend calls don't break the test.
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
          open: async () => {
            (window as any).__tauriDialogOpenCalls += 1;
            return "/tmp/fake.xlsx";
          },
          save: async () => "/tmp/fake-save.xlsx",
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

    // Open a workbook with a real path so AutoSave can save without prompting Save As.
    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["menu-open"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["menu-open"]({ payload: null });
    });

    await page.waitForFunction(() => (window as any).__tauriDialogOpenCalls === 1);
    await page.waitForFunction(async () => (await (window.__formulaApp as any).getCellValueA1("A1")) === "Hello");

    // Enable AutoSave from the File backstage.
    const ribbon = page.getByTestId("ribbon-root");
    await ribbon.getByRole("tab", { name: "File" }).click();
    const autoSave = ribbon.getByTestId("file-auto-save");
    await expect(autoSave).toBeVisible();
    await autoSave.click();

    const initialSaveCount = await page.evaluate(() => {
      const invokes = (window as any).__tauriInvokes as Array<{ cmd: string; args: any }> | undefined;
      return (invokes ?? []).filter((entry) => entry.cmd === "save_workbook").length;
    });

    // Make a committed edit (A1 -> "AutoSaved").
    await page.locator("#grid").focus();
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("F2");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await editor.fill("AutoSaved");
    await page.keyboard.press("Enter");
    await expect(editor).toBeHidden();

    // AutoSave should invoke save_workbook automatically within the debounce timeout.
    await page.waitForFunction(
      (previousCount) => {
        const invokes = (window as any).__tauriInvokes as Array<{ cmd: string; args: any }> | undefined;
        const count = (invokes ?? []).filter((entry) => entry.cmd === "save_workbook").length;
        return count > Number(previousCount);
      },
      initialSaveCount,
      { timeout: 12_000 },
    );
  });
});
