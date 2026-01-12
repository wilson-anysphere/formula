import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

function installTauriStubForTests(
  options: {
    usedRange?: { start_row: number; end_row: number; start_col: number; end_col: number };
    sheets?: Array<{ id: string; name: string; visibility?: string }>;
  } = {},
) {
  const listeners: Record<string, any> = {};
  (window as any).__tauriListeners = listeners;
  (window as any).__tauriInvokeCalls = [];
  const usedRange = options.usedRange ?? { start_row: 0, end_row: 0, start_col: 0, end_col: 0 };
  const sheets =
    Array.isArray(options.sheets) && options.sheets.length > 0 ? options.sheets : [{ id: "Sheet1", name: "Sheet1" }];

  const pushCall = (cmd: string, args: any) => {
    (window as any).__tauriInvokeCalls.push({ cmd, args });
  };

  (window as any).__TAURI__ = {
    core: {
      invoke: async (cmd: string, args: any) => {
        pushCall(cmd, args);
        switch (cmd) {
          case "open_workbook":
            return {
              path: args?.path ?? null,
              origin_path: args?.path ?? null,
              sheets,
            };
          case "new_workbook":
            return {
              path: null,
              origin_path: null,
              sheets,
            };

          case "get_sheet_used_range":
            return usedRange;

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

          case "add_sheet":
            // Purposefully return an id that differs from the requested name to ensure
            // the frontend treats the backend response as canonical.
            return { id: String(args?.sheet_id ?? args?.id ?? "Sheet2-backend"), name: String(args?.name ?? "Sheet2") };

          case "list_defined_names":
            return [];
          case "list_tables":
            return [];
          case "get_workbook_theme_palette":
            return null;

          case "get_macro_security_status":
            return { has_macros: false, trust: "trusted_always" };
          case "set_macro_trust":
            return { has_macros: false, trust: args?.decision ?? "trusted_once" };

          case "stat_file":
            return { mtime_ms: 0, size_bytes: 0 };

          case "set_macro_ui_context":
          case "fire_workbook_open":
            return { ok: true, output: [], updates: [] };

          case "mark_saved":
          case "delete_sheet":
          case "rename_sheet":
          case "move_sheet":
          case "set_sheet_visibility":
          case "set_sheet_tab_color":
          case "set_cell":
          case "set_range":
          case "save_workbook":
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
}

test.describe("tauri workbook integration", () => {
  test("file-dropped event opens a workbook and populates the document", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests, {
      sheets: [
        { id: "sheet-1", name: "Budget" },
        { id: "sheet-2", name: "Summary" },
      ],
    });

    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await page.waitForFunction(async () => (await (window as any).__formulaApp.getCellValueA1("A1")) === "Hello");

    await expect(page.getByTestId("sheet-tab-sheet-1")).toHaveText("Budget");
    await expect(page.getByTestId("sheet-tab-sheet-2")).toHaveText("Summary");

    await expect(page.getByTestId("sheet-switcher")).toHaveValue("sheet-1");
    await expect(page.getByTestId("sheet-switcher").locator("option")).toHaveText(["Budget", "Summary"]);
    await expect(page.getByTestId("active-cell")).toHaveText("A1");
    await expect(page.getByTestId("active-value")).toHaveText("Hello");

    const a1 = await page.evaluate(() => (window as any).__formulaApp.getCellValueA1("A1"));
    expect(a1).toBe("Hello");
  });

  test("hidden sheets are excluded from sheet switcher options", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests, {
      sheets: [
        { id: "Sheet1", name: "Sheet1", visibility: "visible" },
        { id: "Sheet2", name: "Sheet2", visibility: "hidden" },
        { id: "Sheet3", name: "Sheet3", visibility: "visible" },
      ],
    });

    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await page.waitForFunction(async () => (await (window as any).__formulaApp.getCellValueA1("A1")) === "Hello");

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveCount(0);
    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();

    const switcher = page.getByTestId("sheet-switcher");
    await expect(switcher).toHaveValue("Sheet1");
    await expect(switcher.locator("option")).toHaveText(["Sheet1", "Sheet3"]);

    // Unhide Sheet2 and ensure it appears again in the correct order.
    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(() => {
      const tab = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
      if (!tab) throw new Error("Missing Sheet1 tab");
      const rect = tab.getBoundingClientRect();
      tab.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          clientX: rect.left + 10,
          clientY: rect.top + 10,
        }),
      );
    });
    const menu = page.getByTestId("sheet-tab-context-menu");
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Unhide…" }).click();
    await menu.getByRole("button", { name: "Sheet2" }).click();

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(switcher.locator("option")).toHaveText(["Sheet1", "Sheet2", "Sheet3"]);
  });

  test("workbooks with zero visible sheets still render a fallback tab + sheet position and allow unhiding", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests, {
      sheets: [
        { id: "Sheet1", name: "Sheet1", visibility: "hidden" },
        { id: "Sheet2", name: "Sheet2", visibility: "hidden" },
      ],
    });

    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await page.waitForFunction(async () => (await (window as any).__formulaApp.getCellValueA1("A1")) === "Hello");

    // Even though all sheets are marked hidden, the UI should remain usable (defensive fallback):
    // show exactly one sheet (prefer the active sheet) so users can reach the Unhide… menu.
    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveCount(0);
    await expect(page.getByTestId("sheet-position")).toHaveText("Sheet 1 of 1");

    const switcher = page.getByTestId("sheet-switcher");
    await expect(switcher.locator("option")).toHaveText(["Sheet1"]);

    // Unhide Sheet2 via the sheet tab context menu.
    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(() => {
      const tab = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
      if (!tab) throw new Error("Missing Sheet1 tab");
      const rect = tab.getBoundingClientRect();
      tab.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          clientX: rect.left + 10,
          clientY: rect.top + 10,
        }),
      );
    });

    const menu = page.getByTestId("sheet-tab-context-menu");
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Unhide…" }).click();
    await expect(menu.getByRole("button", { name: "Sheet2" })).toBeVisible();
    await menu.getByRole("button", { name: "Sheet2" }).click();

    // After unhiding, Sheet2 should become the active/visible sheet (since Sheet1 remains hidden).
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveCount(0);
    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId()))
      .toBe("Sheet2");
    await expect(switcher.locator("option")).toHaveText(["Sheet2"]);

    // Ensure the unhide action persisted via the tauri backend.
    await expect
      .poll(() =>
        page.evaluate(
          () =>
            ((window as any).__tauriInvokeCalls ?? []).filter(
              (c: any) => c?.cmd === "set_sheet_visibility" && c?.args?.sheet_id === "Sheet2" && c?.args?.visibility === "visible",
            ).length,
        ),
      )
      .toBe(1);
  });

  test("unhide via the tab strip background context menu persists through the tauri backend", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests, {
      sheets: [
        { id: "Sheet1", name: "Sheet1", visibility: "visible" },
        { id: "Sheet2", name: "Sheet2", visibility: "hidden" },
      ],
    });

    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await page.waitForFunction(async () => (await (window as any).__formulaApp.getCellValueA1("A1")) === "Hello");

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveCount(0);

    // Open the tab strip background menu (Excel-like "Unhide..." entry point).
    await page.evaluate(() => {
      const strip = document.querySelector<HTMLElement>("#sheet-tabs .sheet-tabs");
      if (!strip) throw new Error("Missing sheet tab strip");
      const rect = strip.getBoundingClientRect();
      strip.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          clientX: rect.left + rect.width - 4,
          clientY: rect.top + rect.height / 2,
        }),
      );
    });

    const menu = page.getByTestId("sheet-tab-context-menu");
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Unhide…" }).click();
    await page.getByTestId("quick-pick").getByRole("button", { name: "Sheet2" }).click();

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    await expect
      .poll(() =>
        page.evaluate(
          () =>
            ((window as any).__tauriInvokeCalls ?? []).filter(
              (c: any) => c?.cmd === "set_sheet_visibility" && c?.args?.sheet_id === "Sheet2" && c?.args?.visibility === "visible",
            ).length,
        ),
      )
      .toBe(1);
  });

  test("drag reordering visible tabs calls move_sheet with absolute indices (including hidden sheets)", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests, {
      sheets: [
        { id: "Sheet1", name: "Sheet1", visibility: "visible" },
        { id: "Sheet2", name: "Sheet2", visibility: "hidden" },
        { id: "Sheet3", name: "Sheet3", visibility: "visible" },
        { id: "Sheet4", name: "Sheet4", visibility: "visible" },
      ],
    });

    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await page.waitForFunction(async () => (await (window as any).__formulaApp.getCellValueA1("A1")) === "Hello");

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveCount(0);
    await expect(page.getByTestId("sheet-tab-Sheet3")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet4")).toBeVisible();

    // Drag Sheet4 before Sheet3 (Sheet2 is hidden but should still affect the absolute index).
    // Use a synthetic drop event for determinism (Playwright drag/drop can be flaky in the desktop shell).
    await page.evaluate(() => {
      const target = document.querySelector('[data-testid="sheet-tab-Sheet3"]') as HTMLElement | null;
      if (!target) throw new Error("Missing sheet-tab-Sheet3");

      const rect = target.getBoundingClientRect();
      const dt = new DataTransfer();
      dt.setData("text/plain", "Sheet4");

      const drop = new DragEvent("drop", {
        bubbles: true,
        cancelable: true,
        clientX: rect.left + 1,
        clientY: rect.top + rect.height / 2,
      });
      Object.defineProperty(drop, "dataTransfer", { value: dt });
      target.dispatchEvent(drop);
    });

    // Visible order should now be: [Sheet1, Sheet4, Sheet3] (Sheet2 remains hidden).
    await expect.poll(() =>
      page.evaluate(() =>
        Array.from(document.querySelectorAll("#sheet-tabs .sheet-tabs [data-sheet-id]"))
          .map((el) => (el as HTMLElement).getAttribute("data-sheet-id"))
          .filter(Boolean),
      ),
    ).toEqual(["Sheet1", "Sheet4", "Sheet3"]);

    // Ensure the backend was called with the full-workbook insertion index:
    // [Sheet1, Sheet2(hidden), Sheet3, Sheet4] -> move Sheet4 before Sheet3 => to_index === 2.
    await page.waitForFunction(
      () => (window as any).__tauriInvokeCalls?.some?.((call: any) => call?.cmd === "move_sheet"),
      undefined,
      { timeout: 5_000 },
    );
    const moveCall = await page.evaluate(() => {
      const calls = (window as any).__tauriInvokeCalls ?? [];
      return calls.find((call: any) => call?.cmd === "move_sheet") ?? null;
    });
    expect(moveCall?.args).toEqual({ sheet_id: "Sheet4", to_index: 2 });
  });

  test("veryHidden sheets are excluded and not offered in the unhide menu", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests, {
      sheets: [
        { id: "Sheet1", name: "Sheet1", visibility: "visible" },
        { id: "Sheet2", name: "Sheet2", visibility: "hidden" },
        { id: "Sheet3", name: "Sheet3", visibility: "veryHidden" },
        { id: "Sheet4", name: "Sheet4", visibility: "visible" },
      ],
    });

    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await page.waitForFunction(async () => (await (window as any).__formulaApp.getCellValueA1("A1")) === "Hello");

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveCount(0);
    await expect(page.getByTestId("sheet-tab-Sheet3")).toHaveCount(0);
    await expect(page.getByTestId("sheet-tab-Sheet4")).toBeVisible();

    const switcher = page.getByTestId("sheet-switcher");
    await expect(switcher.locator("option")).toHaveText(["Sheet1", "Sheet4"]);

    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(() => {
      const tab = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
      if (!tab) throw new Error("Missing Sheet1 tab");
      const rect = tab.getBoundingClientRect();
      tab.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          clientX: rect.left + 10,
          clientY: rect.top + 10,
        }),
      );
    });
    const menu = page.getByTestId("sheet-tab-context-menu");
    await expect(menu).toBeVisible();

    await expect(menu.getByRole("button", { name: "Unhide…" })).toBeVisible();
    await menu.getByRole("button", { name: "Unhide…" }).click();
    await expect(menu.getByRole("button", { name: "Sheet2" })).toBeVisible();
    await expect(menu.getByRole("button", { name: "Sheet3" })).toHaveCount(0);

    await menu.getByRole("button", { name: "Sheet2" }).click();

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet3")).toHaveCount(0);
    await expect(switcher.locator("option")).toHaveText(["Sheet1", "Sheet2", "Sheet4"]);
  });

  test("warns when workbook exceeds the current load limit", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests, {
      // Exceeds the default maxRows cap (10,000). Keep end_col small so the test avoids
      // large range allocations.
      usedRange: { start_row: 0, end_row: 10_000, start_col: 0, end_col: 0 },
    });

    await gotoDesktop(page);

    // Clear any startup toasts so we can assert on the truncation warning cleanly.
    await page.evaluate(() => {
      document.getElementById("toast-root")?.replaceChildren();
    });

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await expect(page.getByTestId("toast-root")).toContainText(
      "Workbook partially loaded (limited to rows 1-10,000, cols 1-200).",
      { timeout: 30_000 },
    );
  });

  test("load limits can be overridden via query params", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests, {
      usedRange: { start_row: 0, end_row: 10, start_col: 0, end_col: 0 },
    });

    await gotoDesktop(page, "/?loadMaxRows=5&loadMaxCols=6");

    await page.evaluate(() => {
      document.getElementById("toast-root")?.replaceChildren();
    });

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await expect(page.getByTestId("toast-root")).toContainText("Workbook partially loaded (limited to rows 1-5, cols 1-6).", {
      timeout: 30_000,
    });
  });

  test("sheet add button calls backend add_sheet and uses returned id for subsequent sync", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests);

    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await expect(page.getByTestId("sheet-switcher")).toHaveValue("Sheet1");

    await page.getByTestId("sheet-add").click();

    // Frontend should trust the backend id.
    await expect(page.getByTestId("sheet-tab-Sheet2-backend")).toBeVisible();
    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCurrentSheetId()))
      .toBe("Sheet2-backend");

    // Ensure the backend command was invoked with the next SheetN name.
    await expect
      .poll(() => page.evaluate(() => (window as any).__tauriInvokeCalls?.filter((c: any) => c.cmd === "add_sheet")?.length ?? 0))
      .toBe(1);
    await expect
      .poll(() => page.evaluate(() => (window as any).__tauriInvokeCalls?.find((c: any) => c.cmd === "add_sheet")?.args?.name))
      .toBe("Sheet2");
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).__tauriInvokeCalls?.find((c: any) => c.cmd === "add_sheet")?.args?.after_sheet_id),
      )
      .toBe("Sheet1");

    // Mutate the document to ensure workbook sync uses the backend sheet id (not the requested name).
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const sheetId = app.getCurrentSheetId();
      app.getDocument().setCellValue(sheetId, "A1", "Hello from Sheet2");
    });

    await page.waitForFunction(() =>
      ((window as any).__tauriInvokeCalls ?? []).some(
        (c: any) =>
          (c.cmd === "set_cell" || c.cmd === "set_range") &&
          (c.args?.sheet_id ?? c.args?.sheetId) === "Sheet2-backend",
      ),
    );
  });

  test("sheet hide/unhide + tab color persist through tauri backend", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests, {
      sheets: [
        { id: "Sheet1", name: "Sheet1", visibility: "visible" },
        { id: "Sheet2", name: "Sheet2", visibility: "visible" },
      ],
    });

    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await page.waitForFunction(async () => (await (window as any).__formulaApp.getCellValueA1("A1")) === "Hello");

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Hide Sheet1.
    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(() => {
      const tab = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
      if (!tab) throw new Error("Missing Sheet1 tab");
      const rect = tab.getBoundingClientRect();
      tab.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          clientX: rect.left + 10,
          clientY: rect.top + 10,
        }),
      );
    });
    const menu = page.getByTestId("sheet-tab-context-menu");
    await expect(menu).toBeVisible();
    // "Unhide…" also contains "Hide" in its accessible name; require an exact match.
    await menu.getByRole("button", { name: "Hide", exact: true }).click();

    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveCount(0);
    await expect(page.getByTestId("sheet-switcher")).toHaveValue("Sheet2");

    await expect
      .poll(() =>
        page.evaluate(
          () =>
            ((window as any).__tauriInvokeCalls ?? []).filter(
              (c: any) => c?.cmd === "set_sheet_visibility" && c?.args?.sheet_id === "Sheet1" && c?.args?.visibility === "hidden",
            ).length,
        ),
      )
      .toBe(1);

    // Unhide Sheet1.
    await page.evaluate(() => {
      const tab = document.querySelector('[data-testid="sheet-tab-Sheet2"]') as HTMLElement | null;
      if (!tab) throw new Error("Missing Sheet2 tab");
      const rect = tab.getBoundingClientRect();
      tab.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          clientX: rect.left + 10,
          clientY: rect.top + 10,
        }),
      );
    });
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Unhide…", exact: true }).click();
    await menu.getByRole("button", { name: "Sheet1" }).click();

    await expect(page.getByTestId("sheet-tab-Sheet1")).toBeVisible();

    await expect
      .poll(() =>
        page.evaluate(
          () =>
            ((window as any).__tauriInvokeCalls ?? []).filter(
              (c: any) => c?.cmd === "set_sheet_visibility" && c?.args?.sheet_id === "Sheet1" && c?.args?.visibility === "visible",
            ).length,
        ),
      )
      .toBe(1);

    // Set tab color for Sheet1.
    // Open via synthetic contextmenu event and use arrow navigation to avoid pointer-driven
    // scrolling that can close the menu (ContextMenu closes on window scroll/wheel events).
    await page
      .getByTestId("sheet-tab-Sheet1")
      .dispatchEvent("contextmenu", { button: 2, clientX: 0, clientY: 0, bubbles: true, cancelable: true });
    await expect(menu).toBeVisible();
    await page.keyboard.press("ArrowDown"); // Rename -> Hide
    await page.keyboard.press("ArrowDown"); // Hide -> Tab Color
    await page.keyboard.press("ArrowRight"); // Open submenu.
    const submenu = menu.locator(".context-menu__submenu");
    await expect(submenu).toBeVisible();
    await page.keyboard.press("ArrowDown"); // No Color -> Red
    await page.keyboard.press("Enter");

    await expect
      .poll(() =>
        page.evaluate(
          () =>
            ((window as any).__tauriInvokeCalls ?? []).find(
              (c: any) => c?.cmd === "set_sheet_tab_color" && c?.args?.sheet_id === "Sheet1",
            )?.args?.tab_color?.rgb ?? null,
        ),
      )
      .toBe("FFFF0000");
  });

  test("undo restore of deleted renamed sheet re-adds backend sheet using stable id", async ({ page }) => {
    await page.addInitScript(installTauriStubForTests, {
      sheets: [
        { id: "Sheet1", name: "Budget" },
        { id: "Sheet2", name: "Sheet2" },
      ],
    });

    await gotoDesktop(page);

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await page.waitForFunction(async () => (await (window as any).__formulaApp.getCellValueA1("A1")) === "Hello");

    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveText("Budget");
    await expect(page.getByTestId("sheet-tab-Sheet2")).toHaveText("Sheet2");

    // Delete the renamed sheet (Sheet1 -> "Budget").
    // Avoid flaky right-click handling in the desktop shell; dispatch a deterministic contextmenu event.
    await page.evaluate(() => {
      const tab = document.querySelector('[data-testid="sheet-tab-Sheet1"]') as HTMLElement | null;
      if (!tab) throw new Error("Missing Sheet1 tab");
      const rect = tab.getBoundingClientRect();
      tab.dispatchEvent(
        new MouseEvent("contextmenu", {
          bubbles: true,
          cancelable: true,
          clientX: rect.left + 10,
          clientY: rect.top + 10,
        }),
      );
    });
    const menu = page.getByTestId("sheet-tab-context-menu");
    await expect(menu).toBeVisible();

    await menu.getByRole("button", { name: "Delete" }).click();
    // The desktop web harness uses the non-blocking quick-pick dialog instead of the
    // browser-native `window.confirm` prompt.
    await page.getByTestId("quick-pick").getByTestId("quick-pick-item-0").click();

    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveCount(0);

    // Confirm backend delete_sheet invoked.
    await expect
      .poll(() => page.evaluate(() => (window as any).__tauriInvokeCalls?.filter((c: any) => c.cmd === "delete_sheet")?.length ?? 0))
      .toBe(1);
    await expect
      .poll(() => page.evaluate(() => (window as any).__tauriInvokeCalls?.find((c: any) => c.cmd === "delete_sheet")?.args?.sheet_id))
      .toBe("Sheet1");

    // Undo should restore the sheet, and main.ts should reconcile the backend by creating the sheet with the same id.
    await page.evaluate(async () => {
      const registry = window.__formulaCommandRegistry as any;
      await registry.executeCommand("edit.undo");
      await (window.__formulaApp as any).whenIdle();
    });

    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveText("Budget");

    await expect
      .poll(() =>
        page.evaluate(
          () =>
            (window as any).__tauriInvokeCalls?.some(
              (c: any) => c.cmd === "add_sheet" && c.args?.sheet_id === "Sheet1" && c.args?.name === "Budget",
            ) ?? false,
        ),
      )
      .toBe(true);

    // Optional: a subsequent cell edit on the restored sheet should flush to the backend using the stable sheet id.
    await page.getByTestId("sheet-tab-Sheet1").click();
    const beforeEdits = await page.evaluate(() => (window as any).__tauriInvokeCalls?.length ?? 0);
    await page.evaluate(() => {
      const app = (window as any).__formulaApp;
      const doc = app.getDocument();
      doc.setCellValue("Sheet1", "A1", "AfterUndo");
      app.refresh();
    });
    await page.waitForFunction(
      (beforeEdits) =>
        ((window as any).__tauriInvokeCalls ?? []).slice(beforeEdits).some(
          (c: any) =>
            (c.cmd === "set_cell" || c.cmd === "set_range") && (c.args?.sheet_id ?? c.args?.sheetId) === "Sheet1",
        ),
      beforeEdits,
    );
  });
});
