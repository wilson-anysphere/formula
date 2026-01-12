import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("ribbon File save (tauri)", () => {
  test("Save commits an in-progress cell edit before saving", async ({ page }) => {
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

    // Open a workbook (so save is enabled and routed through the workbook sync bridge).
    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["menu-open"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["menu-open"]({ payload: null });
    });

    await page.waitForFunction(() => (window as any).__tauriDialogOpenCalls === 1);
    await page.waitForFunction(async () => (await (window.__formulaApp as any).getCellValueA1("A1")) === "Hello");

    // Start editing A1 but do not commit (no Enter).
    // In the Tauri-stubbed environment, pointer events may be contended by async
    // workbook-open flows; focusing the grid is sufficient to drive keyboard edits.
    await page.locator("#grid").focus();
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    await page.keyboard.press("F2");
    const editor = page.locator("textarea.cell-editor");
    await expect(editor).toBeVisible();
    await editor.fill("Unsaved");
    await expect(editor).toHaveValue("Unsaved");

    // File -> Save should commit the editor value first, then flush it via set_cell, then save_workbook.
    const ribbon = page.getByTestId("ribbon-root");
    await ribbon.getByRole("tab", { name: "File" }).click();
    const save = ribbon.getByTestId("file-save");
    await expect(save).toBeVisible();
    await save.click();

    await page.waitForFunction(() => {
      const invokes = (window as any).__tauriInvokes as Array<{ cmd: string; args: any }> | undefined;
      return Boolean(invokes?.some((entry) => entry.cmd === "save_workbook"));
    });

    // The editor should be closed and selection should not have moved.
    await expect(editor).toBeHidden();
    await expect(page.getByTestId("active-cell")).toHaveText("A1");

    const value = await page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A1"));
    expect(value).toBe("Unsaved");

    const invokes = await page.evaluate(() => (window as any).__tauriInvokes as Array<{ cmd: string; args: any }>);
    const editIndex = invokes.findIndex((entry) => {
      if (entry.cmd === "set_cell") {
        return entry.args?.row === 0 && entry.args?.col === 0 && entry.args?.value === "Unsaved";
      }
      if (entry.cmd === "set_range") {
        const values = entry.args?.values;
        const first = Array.isArray(values) && Array.isArray(values[0]) ? values[0][0] : null;
        return (
          entry.args?.start_row === 0 &&
          entry.args?.end_row === 0 &&
          entry.args?.start_col === 0 &&
          entry.args?.end_col === 0 &&
          first?.value === "Unsaved"
        );
      }
      return false;
    });
    expect(editIndex).toBeGreaterThanOrEqual(0);

    const saveIndex = invokes.findIndex((entry) => entry.cmd === "save_workbook");
    expect(saveIndex).toBeGreaterThan(editIndex);
  });
});
