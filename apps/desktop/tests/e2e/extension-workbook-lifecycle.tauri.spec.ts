import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

test.describe("extension workbook lifecycle (tauri)", () => {
  test("workbook.createWorkbook/saveAs delegate to the desktop tauri flows", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            switch (cmd) {
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
                const values = Array.from({ length: rows }, () =>
                  Array.from({ length: cols }, () => ({ value: null, formula: null, display_value: "" })),
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
                // Best-effort: ignore unknown commands so unrelated UI features don't
                // break this test when new backend calls are introduced.
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

      // The desktop extension host now uses a real permission prompt UI. Pre-grant
      // the permissions needed by this ad-hoc test extension so the worker can
      // activate without blocking on an interactive modal.
      try {
        const extensionId = "formula-test.wb-test";
        const key = "formula.extensionHost.permissions";
        const existing = (() => {
          try {
            const raw = localStorage.getItem(key);
            return raw ? JSON.parse(raw) : {};
          } catch {
            return {};
          }
        })();

        existing[extensionId] = {
          ...(existing[extensionId] ?? {}),
          "ui.commands": true,
          "workbook.manage": true,
        };

        localStorage.setItem(key, JSON.stringify(existing));
      } catch {
        // ignore
      }
    });

    await gotoDesktop(page);
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager));

    const result = await page.evaluate(async () => {
      const mgr: any = (window as any).__formulaExtensionHostManager;
      if (!mgr) throw new Error("Missing window.__formulaExtensionHostManager");
      const host = mgr.host;
      if (!host) throw new Error("Missing extension host");

      const commandId = "wbTest.createAndSaveAs";
      const extensionId = "formula-test.wb-test";
      const manifest = {
        name: "wb-test",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Create workbook + Save As" }] },
        permissions: ["ui.commands", "workbook.manage"],
      };

      // Avoid module imports so the worker can keep strict import sandboxing enabled in Vite.
      // BrowserExtensionHost loads the extension API runtime into the worker already.
      const code = `
        export async function activate(context) {
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          const beforeSave = [];
          const opened = [];
          context.subscriptions.push(formula.events.onBeforeSave((e) => beforeSave.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(formula.events.onWorkbookOpened((e) => opened.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
            await formula.workbook.createWorkbook();
            await formula.workbook.saveAs(${JSON.stringify("/tmp/ext-save.xlsx")});
            const wb = await formula.workbook.getActiveWorkbook();
            return {
              workbook: { name: wb.name, path: wb.path, sheets: wb.sheets, activeSheet: wb.activeSheet },
              events: { beforeSave, opened }
            };
          }));
        }
        export default { activate };
      `;

      const blob = new Blob([code], { type: "text/javascript" });
      const mainUrl = URL.createObjectURL(blob);

      await host.loadExtension({
        extensionId,
        extensionPath: "memory://wb-test/",
        manifest,
        mainUrl,
      });

      try {
        return await host.executeCommand(commandId);
      } finally {
        URL.revokeObjectURL(mainUrl);
        await host.unloadExtension(extensionId).catch(() => {});
      }
    });

    expect(result.workbook.path).toBe("/tmp/ext-save.xlsx");
    expect(result.workbook.name).toBe("ext-save.xlsx");
    expect(Array.isArray(result.workbook.sheets)).toBe(true);
    expect(result.workbook.sheets.length).toBeGreaterThan(0);
    expect(result.workbook.activeSheet).toEqual({ id: "Sheet1", name: "Sheet1" });

    expect(result.events.beforeSave).toEqual(["/tmp/ext-save.xlsx"]);
    // createWorkbook should emit exactly one workbookOpened for the synthetic new workbook.
    expect(result.events.opened.length).toBe(1);

    const invokes = await page.evaluate(() => (window as any).__tauriInvokes);
    expect(invokes.some((entry: any) => entry?.cmd === "new_workbook")).toBe(true);
    expect(
      invokes.some((entry: any) => entry?.cmd === "save_workbook" && entry?.args?.path === "/tmp/ext-save.xlsx"),
    ).toBe(true);
  });

  test("workbook.openWorkbook delegates to the desktop tauri open flow", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;

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
                const values = Array.from({ length: rows }, () =>
                  Array.from({ length: cols }, () => ({ value: null, formula: null, display_value: "" })),
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

      // Pre-grant permissions for the ad-hoc test extension.
      try {
        const extensionId = "formula-test.wb-test";
        const key = "formula.extensionHost.permissions";
        const existing = (() => {
          try {
            const raw = localStorage.getItem(key);
            return raw ? JSON.parse(raw) : {};
          } catch {
            return {};
          }
        })();

        existing[extensionId] = {
          ...(existing[extensionId] ?? {}),
          "ui.commands": true,
          "workbook.manage": true,
        };

        localStorage.setItem(key, JSON.stringify(existing));
      } catch {
        // ignore
      }
    });

    await gotoDesktop(page);
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager));

    const result = await page.evaluate(async () => {
      const mgr: any = (window as any).__formulaExtensionHostManager;
      if (!mgr) throw new Error("Missing window.__formulaExtensionHostManager");
      const host = mgr.host;
      if (!host) throw new Error("Missing extension host");

      const commandId = "wbTest.openWorkbook";
      const extensionId = "formula-test.wb-test";
      const manifest = {
        name: "wb-test",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Open workbook" }] },
        permissions: ["ui.commands", "workbook.manage"],
      };

      const code = `
        export async function activate(context) {
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          const opened = [];
          context.subscriptions.push(formula.events.onWorkbookOpened((e) => opened.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
            await formula.workbook.openWorkbook(${JSON.stringify("/tmp/opened.xlsx")});
            const wb = await formula.workbook.getActiveWorkbook();
            return { workbook: { name: wb.name, path: wb.path }, events: { opened } };
          }));
        }
        export default { activate };
      `;

      const blob = new Blob([code], { type: "text/javascript" });
      const mainUrl = URL.createObjectURL(blob);

      await host.loadExtension({
        extensionId,
        extensionPath: "memory://wb-test/",
        manifest,
        mainUrl,
      });

      try {
        return await host.executeCommand(commandId);
      } finally {
        URL.revokeObjectURL(mainUrl);
        await host.unloadExtension(extensionId).catch(() => {});
      }
    });

    expect(result.workbook.path).toBe("/tmp/opened.xlsx");
    expect(result.workbook.name).toBe("opened.xlsx");
    expect(result.events.opened).toEqual(["/tmp/opened.xlsx"]);

    const invokes = await page.evaluate(() => (window as any).__tauriInvokes);
    expect(invokes.some((entry: any) => entry?.cmd === "open_workbook" && entry?.args?.path === "/tmp/opened.xlsx")).toBe(
      true,
    );
  });

  test("UI Ctrl+N creates a new workbook and clears the previous workbook path", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            switch (cmd) {
              case "new_workbook":
                return {
                  path: null,
                  origin_path: null,
                  sheets: [{ id: "Sheet1", name: "Sheet1" }],
                };

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
                const values = Array.from({ length: rows }, () =>
                  Array.from({ length: cols }, () => ({ value: null, formula: null, display_value: "" })),
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
                // Best-effort: ignore unknown commands so unrelated UI features don't
                // break this test when new backend calls are introduced.
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
          open: async () => "/tmp/ui-opened.xlsx",
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
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHost), undefined, { timeout: 30_000 });

    // Ctrl+O triggers the UI "Open workbook" flow (promptOpenWorkbook -> openWorkbookFromPath).
    await page.evaluate(() => {
      window.dispatchEvent(new KeyboardEvent("keydown", { key: "o", ctrlKey: true }));
    });

    await page.waitForFunction(
      async () => {
        const host: any = (window as any).__formulaExtensionHost;
        if (!host || typeof host._getActiveWorkbook !== "function") return false;
        const wb = await host._getActiveWorkbook();
        return wb?.path === "/tmp/ui-opened.xlsx";
      },
      undefined,
      { timeout: 10_000 },
    );

    // Ctrl+N triggers the UI "New workbook" flow (handleNewWorkbook).
    await page.evaluate(() => {
      window.dispatchEvent(new KeyboardEvent("keydown", { key: "n", ctrlKey: true }));
    });

    // Regression check: when the new workbook is unsaved (path=null), the extension host should
    // clear any previously-stored path instead of falling back to the old one.
    await page.waitForFunction(
      async () => {
        const host: any = (window as any).__formulaExtensionHost;
        if (!host || typeof host._getActiveWorkbook !== "function") return false;
        const wb = await host._getActiveWorkbook();
        return wb?.path == null;
      },
      undefined,
      { timeout: 10_000 },
    );

    const invokes = await page.evaluate(() => (window as any).__tauriInvokes);
    expect(invokes.some((entry: any) => entry?.cmd === "open_workbook" && entry?.args?.path === "/tmp/ui-opened.xlsx")).toBe(
      true,
    );
    expect(invokes.some((entry: any) => entry?.cmd === "new_workbook")).toBe(true);
  });

  test("workbook.save delegates to save_workbook when the workbook has a path", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;

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
                const values = Array.from({ length: rows }, () =>
                  Array.from({ length: cols }, () => ({ value: null, formula: null, display_value: "" })),
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
                // Best-effort: ignore unknown commands so unrelated UI features don't
                // break this test when new backend calls are introduced.
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

      // Pre-grant permissions for the ad-hoc test extension.
      try {
        const extensionId = "formula-test.wb-test";
        const key = "formula.extensionHost.permissions";
        const existing = (() => {
          try {
            const raw = localStorage.getItem(key);
            return raw ? JSON.parse(raw) : {};
          } catch {
            return {};
          }
        })();

        existing[extensionId] = {
          ...(existing[extensionId] ?? {}),
          "ui.commands": true,
          "workbook.manage": true,
        };

        localStorage.setItem(key, JSON.stringify(existing));
      } catch {
        // ignore
      }
    });

    await gotoDesktop(page);
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager));

    const result = await page.evaluate(async () => {
      const mgr: any = (window as any).__formulaExtensionHostManager;
      if (!mgr) throw new Error("Missing window.__formulaExtensionHostManager");
      const host = mgr.host;
      if (!host) throw new Error("Missing extension host");

      const commandId = "wbTest.openAndSave";
      const extensionId = "formula-test.wb-test";
      const manifest = {
        name: "wb-test",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Open workbook + save" }] },
        permissions: ["ui.commands", "workbook.manage"],
      };

      const code = `
        export async function activate(context) {
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          const beforeSave = [];
          context.subscriptions.push(formula.events.onBeforeSave((e) => beforeSave.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
            await formula.workbook.openWorkbook(${JSON.stringify("/tmp/book.xlsx")});
            await formula.workbook.save();
            return { beforeSave };
          }));
        }
        export default { activate };
      `;

      const blob = new Blob([code], { type: "text/javascript" });
      const mainUrl = URL.createObjectURL(blob);

      await host.loadExtension({
        extensionId,
        extensionPath: "memory://wb-test/",
        manifest,
        mainUrl,
      });

      try {
        return await host.executeCommand(commandId);
      } finally {
        URL.revokeObjectURL(mainUrl);
        await host.unloadExtension(extensionId).catch(() => {});
      }
    });

    expect(result.beforeSave).toEqual(["/tmp/book.xlsx"]);

    const invokes = await page.evaluate(() => (window as any).__tauriInvokes);
    expect(invokes.some((entry: any) => entry?.cmd === "open_workbook" && entry?.args?.path === "/tmp/book.xlsx")).toBe(
      true,
    );
    expect(invokes.some((entry: any) => entry?.cmd === "save_workbook" && JSON.stringify(entry?.args ?? {}) === "{}")).toBe(
      true,
    );
  });

  test("workbook.save prompts for Save As when the workbook is unsaved and emits beforeSave with the chosen path", async ({
    page,
  }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            switch (cmd) {
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
                const values = Array.from({ length: rows }, () =>
                  Array.from({ length: cols }, () => ({ value: null, formula: null, display_value: "" })),
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
                // Best-effort: ignore unknown commands so unrelated UI features don't
                // break this test when new backend calls are introduced.
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
          open: async () => null,
          save: async () => "/tmp/saved-from-save.xlsx",
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

      // Pre-grant permissions for the ad-hoc test extension.
      try {
        const extensionId = "formula-test.wb-save-pathless";
        const key = "formula.extensionHost.permissions";
        const existing = (() => {
          try {
            const raw = localStorage.getItem(key);
            return raw ? JSON.parse(raw) : {};
          } catch {
            return {};
          }
        })();

        existing[extensionId] = {
          ...(existing[extensionId] ?? {}),
          "ui.commands": true,
          "workbook.manage": true,
        };

        localStorage.setItem(key, JSON.stringify(existing));
      } catch {
        // ignore
      }
    });

    await gotoDesktop(page);
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager));

    const result = await page.evaluate(async () => {
      const mgr: any = (window as any).__formulaExtensionHostManager;
      if (!mgr) throw new Error("Missing window.__formulaExtensionHostManager");
      const host = mgr.host;
      if (!host) throw new Error("Missing extension host");

      const commandId = "wbTest.savePathless";
      const extensionId = "formula-test.wb-save-pathless";
      const manifest = {
        name: "wb-save-pathless",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Save workbook (pathless)" }] },
        permissions: ["ui.commands", "workbook.manage"],
      };

      const code = `
        export async function activate(context) {
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          const beforeSave = [];
          const opened = [];
          context.subscriptions.push(formula.events.onBeforeSave((e) => beforeSave.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(formula.events.onWorkbookOpened((e) => opened.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
            await formula.workbook.createWorkbook();
            await formula.workbook.save();
            const wb = await formula.workbook.getActiveWorkbook();
            return { workbook: { name: wb.name, path: wb.path }, events: { beforeSave, opened } };
          }));
        }
        export default { activate };
      `;

      const blob = new Blob([code], { type: "text/javascript" });
      const mainUrl = URL.createObjectURL(blob);

      await host.loadExtension({
        extensionId,
        extensionPath: "memory://wb-save-pathless/",
        manifest,
        mainUrl,
      });

      try {
        return await host.executeCommand(commandId);
      } finally {
        URL.revokeObjectURL(mainUrl);
        await host.unloadExtension(extensionId).catch(() => {});
      }
    });

    expect(result.workbook.path).toBe("/tmp/saved-from-save.xlsx");
    expect(result.workbook.name).toBe("saved-from-save.xlsx");
    expect(result.events.beforeSave).toEqual(["/tmp/saved-from-save.xlsx"]);
    expect(result.events.opened.length).toBe(1);

    const invokes = await page.evaluate(() => (window as any).__tauriInvokes);
    expect(invokes.some((entry: any) => entry?.cmd === "save_workbook" && entry?.args?.path === "/tmp/saved-from-save.xlsx")).toBe(
      true,
    );
  });

  test("workbook.close delegates to the desktop tauri new workbook flow", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            switch (cmd) {
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
                const values = Array.from({ length: rows }, () =>
                  Array.from({ length: cols }, () => ({ value: null, formula: null, display_value: "" })),
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
                // Best-effort: ignore unknown commands so unrelated UI features don't
                // break this test when new backend calls are introduced.
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

      // Pre-grant permissions for the ad-hoc test extension.
      try {
        const extensionId = "formula-test.wb-close-success";
        const key = "formula.extensionHost.permissions";
        const existing = (() => {
          try {
            const raw = localStorage.getItem(key);
            return raw ? JSON.parse(raw) : {};
          } catch {
            return {};
          }
        })();

        existing[extensionId] = {
          ...(existing[extensionId] ?? {}),
          "ui.commands": true,
          "workbook.manage": true,
        };

        localStorage.setItem(key, JSON.stringify(existing));
      } catch {
        // ignore
      }
    });

    await gotoDesktop(page);
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager));

    const result = await page.evaluate(async () => {
      const mgr: any = (window as any).__formulaExtensionHostManager;
      if (!mgr) throw new Error("Missing window.__formulaExtensionHostManager");
      const host = mgr.host;
      if (!host) throw new Error("Missing extension host");

      const commandId = "wbTest.closeWorkbook";
      const extensionId = "formula-test.wb-close-success";
      const manifest = {
        name: "wb-close-success",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Close workbook" }] },
        permissions: ["ui.commands", "workbook.manage"],
      };

      const code = `
        export async function activate(context) {
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          const opened = [];
          context.subscriptions.push(formula.events.onWorkbookOpened((e) => opened.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
            await formula.workbook.createWorkbook();
            await formula.workbook.close();
            const wb = await formula.workbook.getActiveWorkbook();
            return {
              workbook: { name: wb.name, path: wb.path, activeSheet: wb.activeSheet },
              events: { opened }
            };
          }));
        }
        export default { activate };
      `;

      const blob = new Blob([code], { type: "text/javascript" });
      const mainUrl = URL.createObjectURL(blob);

      await host.loadExtension({
        extensionId,
        extensionPath: "memory://wb-close-success/",
        manifest,
        mainUrl,
      });

      try {
        return await host.executeCommand(commandId);
      } finally {
        URL.revokeObjectURL(mainUrl);
        await host.unloadExtension(extensionId).catch(() => {});
      }
    });

    expect(result.workbook.path).toBe(null);
    expect(result.workbook.name).toBe("Workbook");
    expect(result.workbook.activeSheet).toEqual({ id: "Sheet1", name: "Sheet1" });
    expect(result.events.opened.length).toBe(2);

    const invokes = await page.evaluate(() => (window as any).__tauriInvokes);
    expect(invokes.filter((entry: any) => entry?.cmd === "new_workbook").length).toBe(2);
  });

  test("workbook.openWorkbook rejects when the discard-changes prompt is cancelled", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;

      // Cancel the discard prompt.
      window.confirm = () => false;

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
                const values = Array.from({ length: rows }, () =>
                  Array.from({ length: cols }, () => ({ value: null, formula: null, display_value: "" })),
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
                // Best-effort: ignore unknown commands so unrelated UI features don't
                // break this test when new backend calls are introduced.
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

      // Pre-grant permissions for the ad-hoc test extension.
      try {
        const extensionId = "formula-test.wb-cancel";
        const key = "formula.extensionHost.permissions";
        const existing = (() => {
          try {
            const raw = localStorage.getItem(key);
            return raw ? JSON.parse(raw) : {};
          } catch {
            return {};
          }
        })();

        existing[extensionId] = {
          ...(existing[extensionId] ?? {}),
          "ui.commands": true,
          "workbook.manage": true,
        };

        localStorage.setItem(key, JSON.stringify(existing));
      } catch {
        // ignore
      }
    });

    await gotoDesktop(page);
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager));

    // Make the document dirty so opening a workbook triggers the discard prompt.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = window.__formulaApp as any;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, { row: 0, col: 0 }, 123);
    });

    const result = await page.evaluate(async () => {
      const mgr: any = (window as any).__formulaExtensionHostManager;
      if (!mgr) throw new Error("Missing window.__formulaExtensionHostManager");
      const host = mgr.host;
      if (!host) throw new Error("Missing extension host");

      const commandId = "wbTest.cancelOpen";
      const extensionId = "formula-test.wb-cancel";
      const manifest = {
        name: "wb-cancel",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Cancel workbook open" }] },
        permissions: ["ui.commands", "workbook.manage"],
      };

      const code = `
        export async function activate(context) {
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          const opened = [];
          context.subscriptions.push(formula.events.onWorkbookOpened((e) => opened.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
            try {
              await formula.workbook.openWorkbook(${JSON.stringify("/tmp/cancel-open.xlsx")});
              return { ok: true, opened };
            } catch (err) {
              return { ok: false, opened, error: err?.message ?? String(err), name: err?.name ?? null };
            }
          }));
        }
        export default { activate };
      `;

      const blob = new Blob([code], { type: "text/javascript" });
      const mainUrl = URL.createObjectURL(blob);

      await host.loadExtension({
        extensionId,
        extensionPath: "memory://wb-cancel/",
        manifest,
        mainUrl,
      });

      try {
        return await host.executeCommand(commandId);
      } finally {
        URL.revokeObjectURL(mainUrl);
        await host.unloadExtension(extensionId).catch(() => {});
      }
    });

    expect(result.ok).toBe(false);
    expect(result.name).toBe("AbortError");
    expect(String(result.error ?? "")).toMatch(/cancel/i);
    expect(result.opened).toEqual([]);

    const invokes = await page.evaluate(() => (window as any).__tauriInvokes);
    expect(invokes.some((entry: any) => entry?.cmd === "open_workbook")).toBe(false);
  });

  test("workbook.save rejects when the Save As dialog is cancelled", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            switch (cmd) {
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
                const values = Array.from({ length: rows }, () =>
                  Array.from({ length: cols }, () => ({ value: null, formula: null, display_value: "" })),
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
                // Best-effort: ignore unknown commands so unrelated UI features don't
                // break this test when new backend calls are introduced.
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
          open: async () => null,
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

      // Pre-grant permissions for the ad-hoc test extension.
      try {
        const extensionId = "formula-test.wb-save-cancel";
        const key = "formula.extensionHost.permissions";
        const existing = (() => {
          try {
            const raw = localStorage.getItem(key);
            return raw ? JSON.parse(raw) : {};
          } catch {
            return {};
          }
        })();

        existing[extensionId] = {
          ...(existing[extensionId] ?? {}),
          "ui.commands": true,
          "workbook.manage": true,
        };

        localStorage.setItem(key, JSON.stringify(existing));
      } catch {
        // ignore
      }
    });

    await gotoDesktop(page);
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager));

    const result = await page.evaluate(async () => {
      const mgr: any = (window as any).__formulaExtensionHostManager;
      if (!mgr) throw new Error("Missing window.__formulaExtensionHostManager");
      const host = mgr.host;
      if (!host) throw new Error("Missing extension host");

      const commandId = "wbTest.cancelSave";
      const extensionId = "formula-test.wb-save-cancel";
      const manifest = {
        name: "wb-save-cancel",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Cancel workbook save" }] },
        permissions: ["ui.commands", "workbook.manage"],
      };

      const code = `
        export async function activate(context) {
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          const beforeSave = [];
          context.subscriptions.push(formula.events.onBeforeSave((e) => beforeSave.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
            await formula.workbook.createWorkbook();
            try {
              await formula.workbook.save();
              return { ok: true, beforeSave };
            } catch (err) {
              return { ok: false, beforeSave, error: err?.message ?? String(err), name: err?.name ?? null };
            }
          }));
        }
        export default { activate };
      `;

      const blob = new Blob([code], { type: "text/javascript" });
      const mainUrl = URL.createObjectURL(blob);

      await host.loadExtension({
        extensionId,
        extensionPath: "memory://wb-save-cancel/",
        manifest,
        mainUrl,
      });

      try {
        return await host.executeCommand(commandId);
      } finally {
        URL.revokeObjectURL(mainUrl);
        await host.unloadExtension(extensionId).catch(() => {});
      }
    });

    expect(result.ok).toBe(false);
    expect(result.name).toBe("AbortError");
    expect(result.beforeSave).toEqual([]);
    expect(String(result.error ?? "")).toMatch(/save/i);
    expect(String(result.error ?? "")).toMatch(/cancel/i);

    const invokes = await page.evaluate(() => (window as any).__tauriInvokes);
    expect(invokes.some((entry: any) => entry?.cmd === "save_workbook")).toBe(false);
  });

  test("workbook.createWorkbook rejects when the discard-changes prompt is cancelled", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;

      // Cancel the discard prompt.
      window.confirm = () => false;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            switch (cmd) {
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
                const values = Array.from({ length: rows }, () =>
                  Array.from({ length: cols }, () => ({ value: null, formula: null, display_value: "" })),
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
                // Best-effort: ignore unknown commands so unrelated UI features don't
                // break this test when new backend calls are introduced.
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

      // Pre-grant permissions for the ad-hoc test extension.
      try {
        const extensionId = "formula-test.wb-create-cancel";
        const key = "formula.extensionHost.permissions";
        const existing = (() => {
          try {
            const raw = localStorage.getItem(key);
            return raw ? JSON.parse(raw) : {};
          } catch {
            return {};
          }
        })();

        existing[extensionId] = {
          ...(existing[extensionId] ?? {}),
          "ui.commands": true,
          "workbook.manage": true,
        };

        localStorage.setItem(key, JSON.stringify(existing));
      } catch {
        // ignore
      }
    });

    await gotoDesktop(page);
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager));

    // Make the document dirty so creating a workbook triggers the discard prompt.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = window.__formulaApp as any;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, { row: 0, col: 0 }, 123);
    });

    const result = await page.evaluate(async () => {
      const mgr: any = (window as any).__formulaExtensionHostManager;
      if (!mgr) throw new Error("Missing window.__formulaExtensionHostManager");
      const host = mgr.host;
      if (!host) throw new Error("Missing extension host");

      const commandId = "wbTest.cancelCreate";
      const extensionId = "formula-test.wb-create-cancel";
      const manifest = {
        name: "wb-create-cancel",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Cancel workbook create" }] },
        permissions: ["ui.commands", "workbook.manage"],
      };

      const code = `
        export async function activate(context) {
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          const opened = [];
          context.subscriptions.push(formula.events.onWorkbookOpened((e) => opened.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
            try {
              await formula.workbook.createWorkbook();
              return { ok: true, opened };
            } catch (err) {
              return { ok: false, opened, error: err?.message ?? String(err), name: err?.name ?? null };
            }
          }));
        }
        export default { activate };
      `;

      const blob = new Blob([code], { type: "text/javascript" });
      const mainUrl = URL.createObjectURL(blob);

      await host.loadExtension({
        extensionId,
        extensionPath: "memory://wb-create-cancel/",
        manifest,
        mainUrl,
      });

      try {
        return await host.executeCommand(commandId);
      } finally {
        URL.revokeObjectURL(mainUrl);
        await host.unloadExtension(extensionId).catch(() => {});
      }
    });

    expect(result.ok).toBe(false);
    expect(result.name).toBe("AbortError");
    expect(String(result.error ?? "")).toMatch(/cancel/i);
    expect(result.opened).toEqual([]);

    const invokes = await page.evaluate(() => (window as any).__tauriInvokes);
    expect(invokes.some((entry: any) => entry?.cmd === "new_workbook")).toBe(false);
  });

  test("workbook.close rejects when the discard-changes prompt is cancelled", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      const invokes: Array<{ cmd: string; args: any }> = [];

      (window as any).__tauriListeners = listeners;
      (window as any).__tauriInvokes = invokes;

      // Cancel the discard prompt.
      window.confirm = () => false;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            invokes.push({ cmd, args });
            switch (cmd) {
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
                const values = Array.from({ length: rows }, () =>
                  Array.from({ length: cols }, () => ({ value: null, formula: null, display_value: "" })),
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
                // Best-effort: ignore unknown commands so unrelated UI features don't
                // break this test when new backend calls are introduced.
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

      // Pre-grant permissions for the ad-hoc test extension.
      try {
        const extensionId = "formula-test.wb-close-cancel";
        const key = "formula.extensionHost.permissions";
        const existing = (() => {
          try {
            const raw = localStorage.getItem(key);
            return raw ? JSON.parse(raw) : {};
          } catch {
            return {};
          }
        })();

        existing[extensionId] = {
          ...(existing[extensionId] ?? {}),
          "ui.commands": true,
          "workbook.manage": true,
        };

        localStorage.setItem(key, JSON.stringify(existing));
      } catch {
        // ignore
      }
    });

    await gotoDesktop(page);
    await page.waitForFunction(() => Boolean((window as any).__formulaExtensionHostManager));

    // Make the document dirty so closing triggers the discard prompt.
    await page.evaluate(() => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const app: any = window.__formulaApp as any;
      if (!app) throw new Error("Missing window.__formulaApp (desktop e2e harness)");
      const doc = app.getDocument();
      const sheetId = app.getCurrentSheetId();
      doc.setCellValue(sheetId, { row: 0, col: 0 }, 123);
    });

    const result = await page.evaluate(async () => {
      const mgr: any = (window as any).__formulaExtensionHostManager;
      if (!mgr) throw new Error("Missing window.__formulaExtensionHostManager");
      const host = mgr.host;
      if (!host) throw new Error("Missing extension host");

      const commandId = "wbTest.cancelClose";
      const extensionId = "formula-test.wb-close-cancel";
      const manifest = {
        name: "wb-close-cancel",
        version: "1.0.0",
        publisher: "formula-test",
        main: "./dist/extension.mjs",
        engines: { formula: "^1.0.0" },
        activationEvents: [`onCommand:${commandId}`],
        contributes: { commands: [{ command: commandId, title: "Cancel workbook close" }] },
        permissions: ["ui.commands", "workbook.manage"],
      };

      const code = `
        export async function activate(context) {
          const formula = globalThis[Symbol.for("formula.extensionApi.api")];
          const opened = [];
          context.subscriptions.push(formula.events.onWorkbookOpened((e) => opened.push(e?.workbook?.path ?? null)));
          context.subscriptions.push(await formula.commands.registerCommand(${JSON.stringify(commandId)}, async () => {
            try {
              await formula.workbook.close();
              return { ok: true, opened };
            } catch (err) {
              return { ok: false, opened, error: err?.message ?? String(err), name: err?.name ?? null };
            }
          }));
        }
        export default { activate };
      `;

      const blob = new Blob([code], { type: "text/javascript" });
      const mainUrl = URL.createObjectURL(blob);

      await host.loadExtension({
        extensionId,
        extensionPath: "memory://wb-close-cancel/",
        manifest,
        mainUrl,
      });

      try {
        return await host.executeCommand(commandId);
      } finally {
        URL.revokeObjectURL(mainUrl);
        await host.unloadExtension(extensionId).catch(() => {});
      }
    });

    expect(result.ok).toBe(false);
    expect(result.name).toBe("AbortError");
    expect(String(result.error ?? "")).toMatch(/cancel/i);
    expect(result.opened).toEqual([]);

    const invokes = await page.evaluate(() => (window as any).__tauriInvokes);
    expect(invokes.some((entry: any) => entry?.cmd === "new_workbook")).toBe(false);
  });
});
