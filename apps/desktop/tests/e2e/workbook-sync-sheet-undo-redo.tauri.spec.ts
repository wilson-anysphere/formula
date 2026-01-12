import { expect, test, type Page } from "@playwright/test";

import { gotoDesktop, openSheetTabContextMenu } from "./helpers";

async function openWorkbookFromFileDrop(page: Page, path: string = "/tmp/fake.xlsx"): Promise<void> {
  await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]), undefined, {
    timeout: 20_000,
  });
  await page.evaluate((workbookPath) => {
    (window as any).__tauriListeners["file-dropped"]({ payload: [workbookPath] });
  }, path);

  // Wait for the workbook snapshot to be applied (our stub returns "Hello" in A1).
  await expect
    .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1")), { timeout: 20_000 })
    .toBe("Hello");

  // Clear any invokes from the open flow so subsequent assertions are scoped to the sheet operation under test.
  await page.evaluate(() => {
    (window as any).__tauriInvokeCalls = [];
  });
}

function installTauriStubForTests() {
  const listeners: Record<string, any> = {};
  (window as any).__tauriListeners = listeners;
  (window as any).__tauriInvokeCalls = [];

  // Avoid modal prompts blocking the test harness.
  (window as any).confirm = () => true;
  (window as any).alert = () => {};

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

          case "add_sheet_with_id":
            return null;

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
    dialog: {
      open: async () => "/tmp/fake.xlsx",
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
    await openWorkbookFromFileDrop(page);

    const tab = page.getByTestId("sheet-tab-Sheet1");
    await expect(tab).toBeVisible();

    await tab.dblclick();
    const input = tab.locator("input");
    await expect(input).toBeVisible();
    await input.fill("RenamedSheet1");
    await input.press("Enter");
    await expect(tab.locator(".sheet-tab__name")).toHaveText("RenamedSheet1");

    await expect
      .poll(
        async () => {
          const calls = (await getInvokeCalls(page)).filter((c) => c.cmd === "rename_sheet");
          return calls.some((c) => c.args?.sheet_id === "Sheet1" && c.args?.name === "RenamedSheet1");
        },
        { timeout: 20_000 },
      )
      .toBe(true);

    // Prefer app-level undo to avoid relying on keyboard focus.
    await page.evaluate(() => {
      (window as any).__tauriInvokeCalls = [];
    });
    await page.evaluate(() => (window as any).__formulaApp.undo());

    await expect
      .poll(
        async () => {
          const calls = (await getInvokeCalls(page)).filter((c) => c.cmd === "rename_sheet");
          return calls.some((c) => c.args?.sheet_id === "Sheet1" && c.args?.name === "Sheet1");
        },
        { timeout: 20_000 },
      )
      .toBe(true);
  });

  test("hide sheet then undo mirrors set_sheet_visibility back to visible", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);
    await openWorkbookFromFileDrop(page);

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    const menu = await openSheetTabContextMenu(page, "Sheet2");
    await menu.getByRole("button", { name: "Hide", exact: true }).click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveCount(0);

    await expect
      .poll(
        async () =>
          (await getInvokeCalls(page)).some(
            (c) => c.cmd === "set_sheet_visibility" && c.args?.sheet_id === "Sheet2" && c.args?.visibility === "hidden",
          ),
        { timeout: 20_000 },
      )
      .toBe(true);

    await page.evaluate(() => {
      (window as any).__tauriInvokeCalls = [];
    });
    await page.evaluate(() => (window as any).__formulaApp.undo());

    await expect
      .poll(
        async () =>
          (await getInvokeCalls(page)).some(
            (c) => c.cmd === "set_sheet_visibility" && c.args?.sheet_id === "Sheet2" && c.args?.visibility === "visible",
          ),
        { timeout: 20_000 },
      )
      .toBe(true);

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
  });

  test("set tab color then undo mirrors set_sheet_tab_color back to null", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);
    await openWorkbookFromFileDrop(page);

    const menu = await openSheetTabContextMenu(page, "Sheet1");
    await menu.getByRole("button", { name: "Tab Color", exact: true }).click();
    await menu.getByRole("button", { name: "Green", exact: true }).click();
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("data-tab-color", "#00b050");

    await expect
      .poll(
        async () =>
          (await getInvokeCalls(page)).some((c) => c.cmd === "set_sheet_tab_color" && c.args?.sheet_id === "Sheet1"),
        { timeout: 20_000 },
      )
      .toBe(true);

    await page.evaluate(() => {
      (window as any).__tauriInvokeCalls = [];
    });
    await page.evaluate(() => (window as any).__formulaApp.undo());

    await expect
      .poll(
        async () =>
          (await getInvokeCalls(page)).some(
            (c) => c.cmd === "set_sheet_tab_color" && c.args?.sheet_id === "Sheet1" && c.args?.tab_color === null,
          ),
        { timeout: 20_000 },
      )
      .toBe(true);

    await expect(page.getByTestId("sheet-tab-Sheet1")).not.toHaveAttribute("data-tab-color");
  });

  test("reorder sheets then undo mirrors the backend ordering", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await gotoDesktop(page);
    await openWorkbookFromFileDrop(page);

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();

    await expect.poll(async () => await getVisibleSheetTabOrder(page), { timeout: 10_000 }).toEqual(["Sheet1", "Sheet2", "Sheet3"]);

    // Drag Sheet3 onto Sheet1.
    //
    // Dispatch a synthetic `drop` event instead of using Playwright's dragAndDrop helper.
    // The desktop shell can be flaky under load when plumbing `DataTransfer` through real
    // pointer gestures, which can lead to hung/slow tests unrelated to sheet reorder correctness.
    await page.evaluate(() => {
      const target = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
      if (!target) throw new Error("Missing sheet-tab-Sheet1");
      const rect = target.getBoundingClientRect();

      const dt = new DataTransfer();
      dt.setData("text/sheet-id", "Sheet3");
      dt.setData("text/plain", "Sheet3");

      const drop = new DragEvent("drop", {
        bubbles: true,
        cancelable: true,
        // Drop near the left edge so the SheetTabStrip interprets this as "insert before".
        clientX: rect.left + 1,
        clientY: rect.top + rect.height / 2,
      });
      Object.defineProperty(drop, "dataTransfer", { value: dt });
      target.dispatchEvent(drop);
    });

    await expect.poll(async () => await getVisibleSheetTabOrder(page), { timeout: 2_000 }).toEqual(["Sheet3", "Sheet1", "Sheet2"]);

    await expect
      .poll(
        async () => {
          const calls = await getInvokeCalls(page);
          const entry = calls.find((c) => c.cmd === "reorder_sheets");
          const ids = entry?.args?.sheet_ids;
          return Array.isArray(ids) ? ids.join("|") : "";
        },
        { timeout: 20_000 },
      )
      .toBe("Sheet3|Sheet1|Sheet2");

    await page.evaluate(() => {
      (window as any).__tauriInvokeCalls = [];
    });
    await page.evaluate(() => (window as any).__formulaApp.undo());

    await expect.poll(async () => await getVisibleSheetTabOrder(page), { timeout: 10_000 }).toEqual(["Sheet1", "Sheet2", "Sheet3"]);

    await expect
      .poll(
        async () => {
          const calls = await getInvokeCalls(page);
          const entry = calls.find((c) => c.cmd === "reorder_sheets");
          const ids = entry?.args?.sheet_ids;
          return Array.isArray(ids) ? ids.join("|") : "";
        },
        { timeout: 20_000 },
      )
      .toBe("Sheet1|Sheet2|Sheet3");
  });

  test("delete renamed sheet then undo recreates it via add_sheet (stable sheet id)", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);
    await page.addInitScript(() => {
      // Avoid native blocking confirm; use a stubbed confirm so nativeDialogs.confirm proceeds.
      (window as any).confirm = () => true;
    });
    await gotoDesktop(page);
    await openWorkbookFromFileDrop(page);

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    const tab = page.getByTestId("sheet-tab-Sheet1");
    await expect(tab).toBeVisible();

    // Rename (changes display name but keeps sheet id stable).
    await tab.dblclick();
    const input = tab.locator("input");
    await expect(input).toBeVisible();
    await input.fill("Budget");
    await input.press("Enter");
    await expect(tab.locator(".sheet-tab__name")).toHaveText("Budget");

    // Delete via context menu.
    const menu = await openSheetTabContextMenu(page, "Sheet1");
    await menu.getByRole("button", { name: "Delete", exact: true }).click();
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveCount(0);

    await expect
      .poll(
        async () => (await getInvokeCalls(page)).some((c) => c.cmd === "delete_sheet" && c.args?.sheet_id === "Sheet1"),
        { timeout: 20_000 },
      )
      .toBe(true);

    await page.evaluate(() => {
      (window as any).__tauriInvokeCalls = [];
    });
    await page.evaluate(() => (window as any).__formulaApp.undo());

    await expect
      .poll(
        async () =>
          (await getInvokeCalls(page)).some(
            (c) =>
              (c.cmd === "add_sheet" || c.cmd === "add_sheet_with_id") && c.args?.sheet_id === "Sheet1" && c.args?.name === "Budget",
          ),
        { timeout: 20_000 },
      )
      .toBe(true);

    const restoredTab = page.getByTestId("sheet-tab-Sheet1");
    await expect(restoredTab).toBeVisible();
    await expect(restoredTab.locator(".sheet-tab__name")).toHaveText("Budget");
  });
});
