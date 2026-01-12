import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("tauri open-file integration", () => {
  test("open-file event opens a workbook and populates the document", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const emitted: Array<{ event: string; payload: any }> = [];
      const callOrder: Array<{ kind: "listen" | "listen-registered" | "emit"; name: string; seq: number }> = [];
      const invokes: Array<{ cmd: string; args: any }> = [];
      const dialogs: Array<{ kind: "confirm" | "message" | "alert"; message: string; options?: any }> = [];
      let seq = 0;
      let activeWorkbookPath: string | null = null;
      const a1ValueByPath: Record<string, string> = {
        "/tmp/fake.xlsx": "Hello",
        "/tmp/first.xlsx": "First",
        "/tmp/second.xlsx": "Second",
      };
      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;
      (window as any).__tauriCallOrder = callOrder;
      (window as any).__tauriInvokes = invokes;
      (window as any).__tauriDialogCalls = dialogs;

      // eslint-disable-next-line no-alert
      (window as any).confirm = () => true;
      // eslint-disable-next-line no-alert
      (window as any).alert = () => {};

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            switch (cmd) {
              case "open_workbook":
                activeWorkbookPath = typeof args?.path === "string" ? args.path : null;
                return {
                  path: args?.path ?? null,
                  origin_path: args?.path ?? null,
                  sheets: [{ id: "Sheet1", name: "Sheet1" }],
                };

              case "stat_file":
                // Used to compute a workbook signature for caching. Keep stable/deterministic.
                return { mtimeMs: 0, sizeBytes: 0 };

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
                      const value =
                        (activeWorkbookPath && a1ValueByPath[activeWorkbookPath]) ??
                        a1ValueByPath["/tmp/fake.xlsx"] ??
                        "Hello";
                      return { value, formula: null, display_value: value };
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
              case "fire_workbook_before_close":
                return { ok: true, output: [], updates: [] };

              case "set_tray_status":
              case "power_query_state_get":
              case "power_query_state_set":
              case "power_query_refresh_state_get":
              case "power_query_refresh_state_set":
              case "report_startup_webview_loaded":
              case "report_startup_tti":
              case "mark_saved":
              case "get_workbook_theme_palette":
              case "list_defined_names":
              case "list_tables":
                return null;

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
            // Tauri's `listen` is async and only resolves once the backend confirms the
            // handler registration. Simulate that so the test can catch regressions
            // where we emit `open-file-ready` before the `open-file` listener is ready.
            callOrder.push({ kind: "listen", name, seq: ++seq });
            await Promise.resolve();
            listeners[name] = handler;
            callOrder.push({ kind: "listen-registered", name, seq: ++seq });
            return () => {
              delete listeners[name];
            };
          },
          emit: async (event: string, payload?: any) => {
            callOrder.push({ kind: "emit", name: event, seq: ++seq });
            emitted.push({ event, payload });
          },
        },
        dialog: {
          // Opening a workbook may prompt to discard unsaved changes. In e2e we always accept so
          // open-file events deterministically proceed.
          confirm: async (message: string, options?: any) => {
            dialogs.push({ kind: "confirm", message, options });
            return true;
          },
          message: async (message: string, options?: any) => {
            dialogs.push({ kind: "message", message, options });
          },
          alert: async (message: string, options?: any) => {
            dialogs.push({ kind: "alert", message, options });
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

    const ordering = await page.evaluate(() => {
      const calls = (window as any).__tauriCallOrder as Array<{ kind: string; name: string; seq: number }> | undefined;
      if (!Array.isArray(calls)) return null;
      const openFileRegistered = calls.find((c) => c.kind === "listen-registered" && c.name === "open-file")?.seq ?? null;
      const openFileReadyEmitted = calls.find((c) => c.kind === "emit" && c.name === "open-file-ready")?.seq ?? null;
      return { openFileRegistered, openFileReadyEmitted };
    });
    expect(ordering).not.toBeNull();
    expect(ordering!.openFileRegistered).not.toBeNull();
    expect(ordering!.openFileReadyEmitted).not.toBeNull();
    expect(ordering!.openFileReadyEmitted!).toBeGreaterThan(ordering!.openFileRegistered!);

    await page.evaluate(() => {
      (window as any).__tauriListeners["open-file"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await page.waitForFunction(
      () => {
        const invokes = (window as any).__tauriInvokes as Array<{ cmd: string; args: any }> | undefined;
        const dialogs = (window as any).__tauriDialogCalls as Array<{ kind: string }> | undefined;
        if (Array.isArray(dialogs) && dialogs.some((d) => d?.kind === "message" || d?.kind === "alert")) return true;
        if (!Array.isArray(invokes)) return false;
        return invokes.some((i) => i?.cmd === "open_workbook" && i?.args?.path === "/tmp/fake.xlsx");
      },
      undefined,
      { timeout: 20_000 },
    );

    const openWorkbookOutcome = await page.evaluate(() => {
      const invokes = (window as any).__tauriInvokes as Array<{ cmd: string; args: any }> | undefined;
      const dialogs = (window as any).__tauriDialogCalls as Array<{ kind: string; message: string }> | undefined;
      const openWorkbookCalled =
        Array.isArray(invokes) && invokes.some((i) => i?.cmd === "open_workbook" && i?.args?.path === "/tmp/fake.xlsx");
      const errorDialog = Array.isArray(dialogs)
        ? dialogs.find((d) => d?.kind === "message" || d?.kind === "alert")?.message ?? null
        : null;
      return { openWorkbookCalled, errorDialog };
    });
    if (!openWorkbookOutcome.openWorkbookCalled) {
      throw new Error(openWorkbookOutcome.errorDialog ?? "open_workbook was never invoked");
    }

    await page.evaluate(async () => {
      const app = (window as any).__formulaApp;
      if (app && typeof app.whenIdle === "function") {
        await app.whenIdle();
      }
    });

    await expect(page.getByTestId("sheet-switcher")).toHaveValue("Sheet1");
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("Hello", { timeout: 20_000 });

    const openWorkbookPaths = await page.evaluate(() => {
      const invokes = (window as any).__tauriInvokes as Array<{ cmd: string; args: any }> | undefined;
      if (!Array.isArray(invokes)) return null;
      return invokes.filter((i) => i.cmd === "open_workbook").map((i) => i.args?.path);
    });
    expect(openWorkbookPaths).toEqual(["/tmp/fake.xlsx"]);

    const openFileReadyEmitted = await page.evaluate(() =>
      (window as any).__tauriEmittedEvents?.some((entry: any) => entry?.event === "open-file-ready"),
    );
    expect(openFileReadyEmitted).toBe(true);
  });

  test("open-file payload with multiple paths opens them sequentially (last wins)", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const emitted: Array<{ event: string; payload: any }> = [];
      const callOrder: Array<{ kind: "listen" | "listen-registered" | "emit"; name: string; seq: number }> = [];
      const invokes: Array<{ cmd: string; args: any }> = [];
      const dialogs: Array<{ kind: "confirm" | "message" | "alert"; message: string; options?: any }> = [];
      let seq = 0;
      let activeWorkbookPath: string | null = null;
      const a1ValueByPath: Record<string, string> = {
        "/tmp/fake.xlsx": "Hello",
        "/tmp/first.xlsx": "First",
        "/tmp/second.xlsx": "Second",
      };
      (window as any).__tauriListeners = listeners;
      (window as any).__tauriEmittedEvents = emitted;
      (window as any).__tauriCallOrder = callOrder;
      (window as any).__tauriInvokes = invokes;
      (window as any).__tauriDialogCalls = dialogs;

      // eslint-disable-next-line no-alert
      (window as any).confirm = () => true;
      // eslint-disable-next-line no-alert
      (window as any).alert = () => {};

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            switch (cmd) {
              case "open_workbook":
                activeWorkbookPath = typeof args?.path === "string" ? args.path : null;
                return {
                  path: args?.path ?? null,
                  origin_path: args?.path ?? null,
                  sheets: [{ id: "Sheet1", name: "Sheet1" }],
                };

              case "stat_file":
                return { mtimeMs: 0, sizeBytes: 0 };

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
                      const value =
                        (activeWorkbookPath && a1ValueByPath[activeWorkbookPath]) ??
                        a1ValueByPath["/tmp/fake.xlsx"] ??
                        "Hello";
                      return { value, formula: null, display_value: value };
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
              case "fire_workbook_before_close":
                return { ok: true, output: [], updates: [] };

              case "set_tray_status":
              case "power_query_state_get":
              case "power_query_state_set":
              case "power_query_refresh_state_get":
              case "power_query_refresh_state_set":
              case "report_startup_webview_loaded":
              case "report_startup_tti":
              case "mark_saved":
              case "get_workbook_theme_palette":
              case "list_defined_names":
              case "list_tables":
                return null;

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
            callOrder.push({ kind: "listen", name, seq: ++seq });
            await Promise.resolve();
            listeners[name] = handler;
            callOrder.push({ kind: "listen-registered", name, seq: ++seq });
            return () => {
              delete listeners[name];
            };
          },
          emit: async (event: string, payload?: any) => {
            callOrder.push({ kind: "emit", name: event, seq: ++seq });
            emitted.push({ event, payload });
          },
        },
        dialog: {
          confirm: async (message: string, options?: any) => {
            dialogs.push({ kind: "confirm", message, options });
            return true;
          },
          message: async (message: string, options?: any) => {
            dialogs.push({ kind: "message", message, options });
          },
          alert: async (message: string, options?: any) => {
            dialogs.push({ kind: "alert", message, options });
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
      // Payloads come from OS integrations (file associations / argv / single-instance forwarding),
      // so be tolerant of non-string entries or empty strings.
      (window as any).__tauriListeners["open-file"]({
        payload: ["", "/tmp/first.xlsx", "   ", 123, null, "/tmp/second.xlsx"],
      });
    });

    await page.waitForFunction(
      () => {
        const invokes = (window as any).__tauriInvokes as Array<{ cmd: string; args: any }> | undefined;
        const dialogs = (window as any).__tauriDialogCalls as Array<{ kind: string }> | undefined;
        if (Array.isArray(dialogs) && dialogs.some((d) => d?.kind === "message" || d?.kind === "alert")) return true;
        if (!Array.isArray(invokes)) return false;
        const openPaths = invokes.filter((i) => i?.cmd === "open_workbook").map((i) => i?.args?.path);
        return openPaths[openPaths.length - 1] === "/tmp/second.xlsx";
      },
      undefined,
      { timeout: 20_000 },
    );

    const openWorkbookOutcome = await page.evaluate(() => {
      const invokes = (window as any).__tauriInvokes as Array<{ cmd: string; args: any }> | undefined;
      const dialogs = (window as any).__tauriDialogCalls as Array<{ kind: string; message: string }> | undefined;
      const openPaths = Array.isArray(invokes)
        ? invokes.filter((i) => i?.cmd === "open_workbook").map((i) => i?.args?.path)
        : [];
      const errorDialog = Array.isArray(dialogs)
        ? dialogs.find((d) => d?.kind === "message" || d?.kind === "alert")?.message ?? null
        : null;
      return { openPaths, errorDialog };
    });
    if (openWorkbookOutcome.openPaths[openWorkbookOutcome.openPaths.length - 1] !== "/tmp/second.xlsx") {
      throw new Error(openWorkbookOutcome.errorDialog ?? "open_workbook was never invoked for /tmp/second.xlsx");
    }

    await page.evaluate(async () => {
      const app = (window as any).__formulaApp;
      if (app && typeof app.whenIdle === "function") {
        await app.whenIdle();
      }
    });

    const a1Value = await page.evaluate(async () => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1Value).toBe("Second");

    const openWorkbookPaths = await page.evaluate(() => {
      const invokes = (window as any).__tauriInvokes as Array<{ cmd: string; args: any }> | undefined;
      if (!Array.isArray(invokes)) return null;
      return invokes.filter((i) => i.cmd === "open_workbook").map((i) => i.args?.path);
    });
    expect(openWorkbookPaths).toEqual(["/tmp/first.xlsx", "/tmp/second.xlsx"]);
  });
});
