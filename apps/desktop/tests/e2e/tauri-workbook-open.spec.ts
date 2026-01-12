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
            return { id: "Sheet2-backend", name: String(args?.name ?? "Sheet2") };

          case "stat_file":
            return { mtime_ms: 0, size_bytes: 0 };

          case "set_macro_ui_context":
          case "fire_workbook_open":
          case "mark_saved":
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
    await page.getByTestId("sheet-tab-Sheet1").click({ button: "right" });
    const menu = page.getByTestId("sheet-tab-context-menu");
    await expect(menu).toBeVisible();
    await menu.getByRole("button", { name: "Unhide…" }).click();
    await menu.getByRole("button", { name: "Sheet2" }).click();

    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();
    await expect(switcher.locator("option")).toHaveText(["Sheet1", "Sheet2", "Sheet3"]);
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

    await page.getByTestId("sheet-tab-Sheet1").click({ button: "right" });
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
      "Workbook partially loaded (limited to 10,000 rows × 200 cols).",
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

    await expect(page.getByTestId("toast-root")).toContainText(
      "Workbook partially loaded (limited to 5 rows × 6 cols).",
      { timeout: 30_000 },
    );
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
});
