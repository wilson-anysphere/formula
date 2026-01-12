import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("tauri native menu integration", () => {
  test("menu-open opens a workbook; menu-quit triggers the close flow", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;
      (window as any).__tauriDialogOpenCalls = 0;
      (window as any).__tauriClipboardWriteText = null;

      // Avoid accidental modal prompts blocking the test.
      window.confirm = () => true;
      window.alert = () => {};

      // Force menu copy/paste to exercise the fallback code paths instead of relying on
      // browser/OS clipboard behavior (which is often disabled in headless WebViews).
      const originalExecCommand = document.execCommand?.bind(document) ?? null;
      document.execCommand = ((commandId: string, showUI?: boolean, value?: string) => {
        if (commandId === "paste" || commandId === "copy" || commandId === "cut") {
          return false;
        }
        return originalExecCommand ? originalExecCommand(commandId, showUI, value) : false;
      }) as any;

      // Ensure the menu fallback logic uses the Tauri clipboard IPC (deterministic) rather than
      // Chromium's clipboard implementation (which can vary across CI/headless environments).
      //
      // Avoid redefining `window.navigator` / `navigator.clipboard` entirely because those are
      // non-configurable in real browsers. Instead, patch readText/writeText when possible.
      const stubReadText = async () => {
        throw new Error("clipboard read blocked");
      };
      const stubWriteText = async () => {
        throw new Error("clipboard write blocked");
      };

      const clipboard = (navigator as any)?.clipboard;
      if (clipboard && typeof clipboard === "object") {
        try {
          Object.defineProperty(clipboard, "readText", { value: stubReadText, configurable: true });
        } catch {
          try {
            clipboard.readText = stubReadText;
          } catch {
            // ignore
          }
        }
        try {
          Object.defineProperty(clipboard, "writeText", { value: stubWriteText, configurable: true });
        } catch {
          try {
            clipboard.writeText = stubWriteText;
          } catch {
            // ignore
          }
        }
      }

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

              case "fire_workbook_open":
                return { ok: true, output: [], updates: [] };

              case "clipboard_read":
                return { text: "PASTED" };

              case "clipboard_write":
                (window as any).__tauriClipboardWriteText = args?.payload?.text ?? null;
                return null;

              case "get_macro_security_status":
                // Avoid macro trust prompts (not relevant to this test); treat the workbook as macro-free.
                return { has_macros: false, trust: "blocked" };

              case "set_cell":
              case "set_range":
              case "save_workbook":
                return null;

              case "quit_app":
                (window as any).__tauriClosed = true;
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
          open: async () => {
            (window as any).__tauriDialogOpenCalls += 1;
            return "/tmp/fake.xlsx";
          },
          save: async () => null,
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

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["menu-open"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["menu-open"]({ payload: null });
    });

    await page.waitForFunction(() => (window as any).__tauriDialogOpenCalls === 1);
    // `menu-open` triggers an async workbook load; wait for the value to be materialized
    // in the app state (not just for the menu handler to be invoked).
    await expect
      .poll(() => page.evaluate(() => (window.__formulaApp as any).getCellValueA1("A1")), { timeout: 20_000 })
      .toBe("Hello");
    await expect(page.getByTestId("active-value")).toHaveText("Hello");

    // Menu Edit items should work while editing text (formula bar).
    await page.getByTestId("formula-highlight").click();
    const formulaInput = page.getByTestId("formula-input");
    await expect(formulaInput).toBeVisible();
    await formulaInput.fill("CopyMe");
    await page.evaluate(() => {
      const input = document.querySelector<HTMLTextAreaElement>('[data-testid="formula-input"]');
      input?.setSelectionRange(0, input.value.length);
    });

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["menu-copy"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["menu-copy"]({ payload: null });
    });
    await page.waitForFunction(() => (window as any).__tauriClipboardWriteText === "CopyMe");

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["menu-paste"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["menu-paste"]({ payload: null });
    });
    await expect(page.getByTestId("formula-input")).toHaveValue(/PASTED/);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["menu-quit"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["menu-quit"]({ payload: null });
    });

    await page.waitForFunction(() => Boolean((window as any).__tauriClosed));
  });
});
