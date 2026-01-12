import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop } from "./helpers";

function installTauriStubForTests() {
  const listeners: Record<string, any> = {};
  (window as any).__tauriListeners = listeners;
  (window as any).__tauriInvokeCalls = [];

  const pushCall = (cmd: string, args: any) => {
    (window as any).__tauriInvokeCalls.push({ cmd, args });
  };

  let nextSheetNumber = 3;

  const baseWorkbookInfo = (path: string | null, origin: string | null) => ({
    path,
    origin_path: origin,
    sheets: [
      { id: "Sheet1", name: "Sheet1", visibility: "visible" },
      { id: "Sheet2", name: "Sheet2", visibility: "visible" },
      { id: "Sheet3", name: "Sheet3", visibility: "visible" },
    ],
  });

  (window as any).__TAURI__ = {
    core: {
      invoke: async (cmd: string, args: any) => {
        pushCall(cmd, args);
        switch (cmd) {
          case "open_workbook":
            return baseWorkbookInfo(args?.path ?? null, args?.path ?? null);
          case "new_workbook":
            return baseWorkbookInfo(null, null);

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

          case "add_sheet": {
            const requestedId = typeof args?.sheet_id === "string" && args.sheet_id.trim() ? args.sheet_id.trim() : null;
            if (requestedId) {
              return { id: requestedId, name: args?.name ?? requestedId, visibility: "visible" };
            }
            nextSheetNumber += 1;
            const id = `Sheet${nextSheetNumber}`;
            return { id, name: args?.name ?? id, visibility: "visible" };
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
          case "rename_sheet":
          case "move_sheet":
          case "delete_sheet":
          case "set_sheet_visibility":
          case "set_sheet_tab_color":
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

async function getInvokeCalls(page: Page): Promise<Array<{ cmd: string; args: any }>> {
  return await page.evaluate(() => (window as any).__tauriInvokeCalls ?? []);
}

async function getVisibleSheetTabOrder(page: Page): Promise<string[]> {
  return await page.evaluate(() => {
    const strip = document.querySelector(".sheet-tabs");
    if (!strip) throw new Error("Missing .sheet-tabs");
    return Array.from(strip.children)
      .map((child) => (child as HTMLElement).dataset.sheetId ?? "")
      .filter(Boolean);
  });
}

test.describe("tauri workbookSync sheet metadata undo/redo", () => {
  test("rename sheet then undo mirrors rename_sheet back to original", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);

    const tab = page.getByTestId("sheet-tab-Sheet1");
    await expect(tab).toBeVisible();

    await tab.dblclick();
    const input = tab.locator("input");
    await expect(input).toBeVisible();
    await input.fill("RenamedSheet1");
    await input.press("Enter");
    await expect(tab.locator(".sheet-tab__name")).toHaveText("RenamedSheet1");

    await expect
      .poll(async () => (await getInvokeCalls(page)).filter((c) => c.cmd === "rename_sheet").length, { timeout: 10_000 })
      .toBe(1);

    // Prefer app-level undo to avoid relying on keyboard focus.
    await page.evaluate(() => (window as any).__formulaApp.undo());

    await expect
      .poll(async () => (await getInvokeCalls(page)).filter((c) => c.cmd === "rename_sheet").length, { timeout: 10_000 })
      .toBe(2);

    const calls = (await getInvokeCalls(page)).filter((c) => c.cmd === "rename_sheet");
    expect(calls[1]?.args).toMatchObject({ sheet_id: "Sheet1", name: "Sheet1" });
  });

  test("hide sheet then undo mirrors set_sheet_visibility back to visible", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);

    await page.evaluate(() => {
      (window as any).__formulaApp.getWorkbookSheetStore().hide("Sheet2");
    });

    await expect
      .poll(
        async () => (await getInvokeCalls(page)).filter((c) => c.cmd === "set_sheet_visibility").length,
        { timeout: 10_000 },
      )
      .toBe(1);

    await page.evaluate(() => (window as any).__formulaApp.undo());

    await expect
      .poll(
        async () => (await getInvokeCalls(page)).filter((c) => c.cmd === "set_sheet_visibility").length,
        { timeout: 10_000 },
      )
      .toBe(2);

    const calls = (await getInvokeCalls(page)).filter((c) => c.cmd === "set_sheet_visibility");
    expect(calls[1]?.args).toMatchObject({ sheet_id: "Sheet2", visibility: "visible" });
  });

  test("set tab color then undo mirrors set_sheet_tab_color back to null", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);

    await page.evaluate(() => {
      (window as any).__formulaApp.getWorkbookSheetStore().setTabColor("Sheet1", { rgb: "FF00FF00" });
    });

    await expect
      .poll(async () => (await getInvokeCalls(page)).filter((c) => c.cmd === "set_sheet_tab_color").length, { timeout: 10_000 })
      .toBe(1);

    await page.evaluate(() => (window as any).__formulaApp.undo());

    await expect
      .poll(async () => (await getInvokeCalls(page)).filter((c) => c.cmd === "set_sheet_tab_color").length, { timeout: 10_000 })
      .toBe(2);

    const calls = (await getInvokeCalls(page)).filter((c) => c.cmd === "set_sheet_tab_color");
    expect(calls[1]?.args).toMatchObject({ sheet_id: "Sheet1", tab_color: null });
  });

  test("reorder sheets then undo mirrors the backend ordering via move_sheet", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();

    await expect.poll(async () => await getVisibleSheetTabOrder(page), { timeout: 10_000 }).toEqual(["Sheet1", "Sheet2", "Sheet3"]);

    // Drag Sheet3 onto Sheet1.
    await page.dragAndDrop('[data-testid="sheet-tab-Sheet3"]', '[data-testid="sheet-tab-Sheet1"]', {
      // Drop near the left edge so the SheetTabStrip interprets this as "insert before".
      targetPosition: { x: 5, y: 5 },
    });
    // Wait for the UI to re-render in the new order. If Playwright's dragAndDrop
    // fails to plumb dataTransfer, fall back to a synthetic drop event.
    try {
      await expect
        .poll(async () => await getVisibleSheetTabOrder(page), { timeout: 2_000 })
        .toEqual(["Sheet3", "Sheet1", "Sheet2"]);
    } catch {
      await page.evaluate(() => {
        const target = document.querySelector('[data-testid="sheet-tab-Sheet1"]');
        if (!target) throw new Error("Missing sheet-tab-Sheet1");
        const dt = new DataTransfer();
        dt.setData("text/plain", "Sheet3");
        const drop = new DragEvent("drop", { bubbles: true, cancelable: true });
        Object.defineProperty(drop, "dataTransfer", { value: dt });
        target.dispatchEvent(drop);
      });

      await expect
        .poll(async () => await getVisibleSheetTabOrder(page), { timeout: 2_000 })
        .toEqual(["Sheet3", "Sheet1", "Sheet2"]);
    }

    await expect
      .poll(async () => (await getInvokeCalls(page)).filter((c) => c.cmd === "move_sheet").length, { timeout: 10_000 })
      .toBeGreaterThan(0);

    await page.evaluate(() => (window as any).__formulaApp.undo());

    await expect.poll(async () => await getVisibleSheetTabOrder(page), { timeout: 10_000 }).toEqual(["Sheet1", "Sheet2", "Sheet3"]);

    // Undo should trigger a backend reorder. Depending on the algorithm, this may be a single call or multiple.
    await expect
      .poll(async () => (await getInvokeCalls(page)).filter((c) => c.cmd === "move_sheet").length, { timeout: 10_000 })
      .toBeGreaterThan(1);

    const moveCalls = (await getInvokeCalls(page)).filter((c) => c.cmd === "move_sheet");
    expect(moveCalls.some((c) => c.args?.sheet_id === "Sheet3" && c.args?.to_index === 2)).toBe(true);
  });
});
