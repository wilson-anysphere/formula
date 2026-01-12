import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("sheet reorder", () => {
  test("dragging a sheet tab persists the order through reopen", async ({ page }) => {
    await page.addInitScript(() => {
      const STORAGE_KEY = "__formula_e2e_sheet_order__";
      const DEFAULT_ORDER = ["Sheet1", "Sheet2", "Sheet3"];

      function readOrder(): string[] {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (raw) {
          try {
            const parsed = JSON.parse(raw);
            if (Array.isArray(parsed) && parsed.every((x) => typeof x === "string")) {
              return parsed as string[];
            }
          } catch {
            // ignore
          }
        }
        localStorage.setItem(STORAGE_KEY, JSON.stringify(DEFAULT_ORDER));
        return DEFAULT_ORDER.slice();
      }

      function writeOrder(order: string[]): void {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(order));
      }

      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            switch (cmd) {
              case "open_workbook": {
                const order = readOrder();
                return {
                  path: args?.path ?? null,
                  origin_path: args?.path ?? null,
                  sheets: order.map((id) => ({ id, name: id })),
                };
              }

              case "get_sheet_used_range":
                // No cells in this fixture; avoids needing to implement get_range.
                return null;

              case "move_sheet": {
                const sheetId = String(args?.sheet_id ?? "");
                const toIndex = Number(args?.to_index ?? 0);

                const order = readOrder().slice();
                const from = order.indexOf(sheetId);
                if (from === -1) return null;

                const [moved] = order.splice(from, 1);
                const clamped = Math.max(0, Math.min(Math.trunc(toIndex), order.length));
                order.splice(clamped, 0, moved);
                writeOrder(order);
                return null;
              }

              case "set_macro_ui_context":
                return null;

              case "fire_workbook_open":
                return { ok: true, output: [], updates: [] };

              case "set_cell":
              case "set_range":
              case "save_workbook":
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
            hide: async () => {},
            close: async () => {},
          }),
        },
      };
    });

    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/reorder.xlsx"] });
    });

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();

    const tabOrder = async () =>
      page.evaluate(() => {
        // Scope to actual sheet tab elements; other UI (e.g. context menus) may also use `sheet-tab-*` test ids.
        const tabs = Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")) as HTMLElement[];
        return tabs.map((el) => el.getAttribute("data-sheet-id") ?? "");
      });

    const expectSheetPositionMatchesTabOrder = async () => {
      const [order, activeSheetId] = await Promise.all([
        tabOrder(),
        page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId()),
      ]);
      const idx = order.indexOf(activeSheetId);
      expect(idx).toBeGreaterThanOrEqual(0);
      await expect(page.getByTestId("sheet-position")).toHaveText(`Sheet ${idx + 1} of ${order.length}`);
    };

    expect(await tabOrder()).toEqual(["Sheet1", "Sheet2", "Sheet3"]);
    await expectSheetPositionMatchesTabOrder();

    // Drop on the left half of the target tab so the reorder logic chooses "insert before".
    await page.evaluate(() => {
      const fromId = "Sheet3";
      const target = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
      if (!target) throw new Error("Missing Sheet1 tab");
      const rect = target.getBoundingClientRect();

      const dt = new DataTransfer();
      dt.setData("text/sheet-id", fromId);
      dt.setData("text/plain", fromId);

      const drop = new DragEvent("drop", {
        bubbles: true,
        cancelable: true,
        clientX: rect.left + 1,
        clientY: rect.top + rect.height / 2,
      });
      Object.defineProperty(drop, "dataTransfer", { value: dt });
      target.dispatchEvent(drop);
    });

    await page.waitForFunction(() => {
      const tabs = Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")) as HTMLElement[];
      return tabs[0]?.getAttribute("data-sheet-id") === "Sheet3";
    });
    expect(await tabOrder()).toEqual(["Sheet3", "Sheet1", "Sheet2"]);
    await expectSheetPositionMatchesTabOrder();

    // The UI reorders optimistically before the backend persistence call resolves.
    // Wait for the mocked Tauri backend to record the move before reloading.
    await page.waitForFunction(() => {
      const raw = localStorage.getItem("__formula_e2e_sheet_order__");
      if (!raw) return false;
      try {
        const parsed = JSON.parse(raw);
        return Array.isArray(parsed) && parsed[0] === "Sheet3";
      } catch {
        return false;
      }
    });

    // Reload and re-open; backend should return the same sheet order.
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/reorder.xlsx"] });
    });

    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();
    await page.waitForFunction(() => {
      const tabs = Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")) as HTMLElement[];
      return tabs[0]?.getAttribute("data-sheet-id") === "Sheet3";
    });
    expect(await tabOrder()).toEqual(["Sheet3", "Sheet1", "Sheet2"]);
    await expectSheetPositionMatchesTabOrder();
  });

  test("undoing a sheet tab reorder persists the order through reopen", async ({ page }) => {
    await page.addInitScript(() => {
      const STORAGE_KEY = "__formula_e2e_sheet_order__";
      const DEFAULT_ORDER = ["Sheet1", "Sheet2", "Sheet3"];

      function readOrder(): string[] {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (raw) {
          try {
            const parsed = JSON.parse(raw);
            if (Array.isArray(parsed) && parsed.every((x) => typeof x === "string")) {
              return parsed as string[];
            }
          } catch {
            // ignore
          }
        }
        localStorage.setItem(STORAGE_KEY, JSON.stringify(DEFAULT_ORDER));
        return DEFAULT_ORDER.slice();
      }

      function writeOrder(order: string[]): void {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(order));
      }

      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            switch (cmd) {
              case "open_workbook": {
                const order = readOrder();
                return {
                  path: args?.path ?? null,
                  origin_path: args?.path ?? null,
                  sheets: order.map((id) => ({ id, name: id })),
                };
              }

              case "get_sheet_used_range":
                return null;

              case "move_sheet": {
                const sheetId = String(args?.sheet_id ?? "");
                const toIndex = Number(args?.to_index ?? 0);

                const order = readOrder().slice();
                const from = order.indexOf(sheetId);
                if (from === -1) return null;

                const [moved] = order.splice(from, 1);
                const clamped = Math.max(0, Math.min(Math.trunc(toIndex), order.length));
                order.splice(clamped, 0, moved);
                writeOrder(order);
                return null;
              }

              case "set_macro_ui_context":
                return null;

              case "fire_workbook_open":
                return { ok: true, output: [], updates: [] };

              case "set_cell":
              case "set_range":
              case "save_workbook":
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
            hide: async () => {},
            close: async () => {},
          }),
        },
      };
    });

    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/reorder.xlsx"] });
    });

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();

    const tabOrder = async () =>
      page.evaluate(() => {
        const tabs = Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")) as HTMLElement[];
        return tabs.map((el) => el.getAttribute("data-sheet-id") ?? "");
      });

    expect(await tabOrder()).toEqual(["Sheet1", "Sheet2", "Sheet3"]);

    // Move Sheet3 to the front.
    await page.evaluate(() => {
      const fromId = "Sheet3";
      const target = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
      if (!target) throw new Error("Missing Sheet1 tab");
      const rect = target.getBoundingClientRect();

      const dt = new DataTransfer();
      dt.setData("text/sheet-id", fromId);
      dt.setData("text/plain", fromId);

      const drop = new DragEvent("drop", {
        bubbles: true,
        cancelable: true,
        clientX: rect.left + 1,
        clientY: rect.top + rect.height / 2,
      });
      Object.defineProperty(drop, "dataTransfer", { value: dt });
      target.dispatchEvent(drop);
    });

    await page.waitForFunction(() => {
      const tabs = Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")) as HTMLElement[];
      return tabs[0]?.getAttribute("data-sheet-id") === "Sheet3";
    });
    expect(await tabOrder()).toEqual(["Sheet3", "Sheet1", "Sheet2"]);

    // Wait for the mocked backend to record the move.
    await page.waitForFunction(() => {
      const raw = localStorage.getItem("__formula_e2e_sheet_order__");
      if (!raw) return false;
      try {
        const parsed = JSON.parse(raw);
        return Array.isArray(parsed) && parsed[0] === "Sheet3";
      } catch {
        return false;
      }
    });

    // Undo the reorder (doc-driven, store sync runs with syncingSheetUi=true).
    await page.evaluate(() => {
      (window as any).__formulaApp.undo();
    });

    await page.waitForFunction(() => {
      const tabs = Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")) as HTMLElement[];
      return tabs[0]?.getAttribute("data-sheet-id") === "Sheet1";
    });
    expect(await tabOrder()).toEqual(["Sheet1", "Sheet2", "Sheet3"]);

    // Wait for the backend persistence call triggered by the undo to update storage.
    await page.waitForFunction(() => {
      const raw = localStorage.getItem("__formula_e2e_sheet_order__");
      if (!raw) return false;
      try {
        const parsed = JSON.parse(raw);
        return Array.isArray(parsed) && parsed[0] === "Sheet1";
      } catch {
        return false;
      }
    });

    // Reload and re-open; backend should return the undone order.
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/reorder.xlsx"] });
    });

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await page.waitForFunction(() => {
      const tabs = Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]")) as HTMLElement[];
      return tabs[0]?.getAttribute("data-sheet-id") === "Sheet1";
    });
    expect(await tabOrder()).toEqual(["Sheet1", "Sheet2", "Sheet3"]);
  });
});
