import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("tauri open-file integration", () => {
  test("open-file event opens a workbook and populates the document", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const emitted: Array<{ event: string; payload: any }> = [];
      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
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

              case "set_macro_ui_context":
                return null;

              case "get_macro_security_status":
                return { has_macros: false, trust: "trusted_always" };

              case "fire_workbook_open":
                return { ok: true, output: [], updates: [] };

              case "set_cell":
              case "set_range":
              case "save_workbook":
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
          emit: async (event: string, payload?: any) => {
            emitted.push({ event, payload });
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
    });

    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["open-file"]));
    await page.waitForFunction(() =>
      Boolean((window as any).__tauriEmittedEvents?.some((entry: any) => entry?.event === "open-file-ready")),
    );

    await page.evaluate(() => {
      (window as any).__tauriListeners["open-file"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await page.waitForFunction(async () => (await (window as any).__formulaApp.getCellValueA1("A1")) === "Hello");

    await expect(page.getByTestId("sheet-switcher")).toHaveValue("Sheet1");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("Hello");

    const openFileReadyEmitted = await page.evaluate(() =>
      (window as any).__tauriEmittedEvents?.some((entry: any) => entry?.event === "open-file-ready"),
    );
    expect(openFileReadyEmitted).toBe(true);
  });
});

