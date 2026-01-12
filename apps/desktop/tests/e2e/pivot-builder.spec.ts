import { expect, test } from "@playwright/test";

import { gotoDesktop, waitForDesktopReady } from "./helpers";

test.describe("pivot builder", () => {
  test("creates and refreshes a pivot table from the current selection", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;

      window.confirm = () => true;

      const pivotTables: any[] = [];

      function displayValue(value: any): string {
        if (value == null) return "";
        if (typeof value === "number") return Number.isFinite(value) ? String(value) : "";
        return String(value);
      }

      function computePivotUpdates(request: any): any[] {
        const doc = (window as any).__formulaApp?.getDocument?.();
        if (!doc) return [];

        const sheetId = request.source_sheet_id;
        const startRow = request.source_range.start_row;
        const endRow = request.source_range.end_row;
        const startCol = request.source_range.start_col;
        const endCol = request.source_range.end_col;

        const headers: string[] = [];
        for (let c = startCol; c <= endCol; c += 1) {
          headers.push(String(doc.getCell(sheetId, { row: startRow, col: c }).value ?? ""));
        }

        const rowField = request.config?.rowFields?.[0]?.sourceField;
        const valueField = request.config?.valueFields?.[0];
        const valueAgg = valueField?.aggregation ?? "sum";
        const valueName = valueField?.name ?? "Value";
        const includeRowGrand = Boolean(request.config?.grandTotals?.rows ?? true);
        const includeColGrand = Boolean(request.config?.grandTotals?.columns ?? true);

        const rowFieldCol = headers.indexOf(rowField);
        const valueFieldCol = headers.indexOf(valueField?.sourceField);

        const sums = new Map<string, number>();
        for (let r = startRow + 1; r <= endRow; r += 1) {
          const key = String(doc.getCell(sheetId, { row: r, col: startCol + rowFieldCol }).value ?? "");
          const raw = doc.getCell(sheetId, { row: r, col: startCol + valueFieldCol }).value;
          const num = typeof raw === "number" ? raw : Number(raw);
          const current = sums.get(key) ?? 0;
          if (valueAgg === "count") {
            sums.set(key, current + 1);
          } else if (valueAgg === "average") {
            // Keep it simple for this test (not used).
            sums.set(key, current + (Number.isFinite(num) ? num : 0));
          } else {
            sums.set(key, current + (Number.isFinite(num) ? num : 0));
          }
        }

        const rows = Array.from(sums.keys()).sort((a, b) => a.localeCompare(b));
        const destSheetId = request.destination.sheet_id;
        const destRow = request.destination.row;
        const destCol = request.destination.col;

        const grid: any[][] = [];
        const headerRow: any[] = [rowField, valueName];
        if (includeColGrand) headerRow.push(`Grand Total - ${valueName}`);
        grid.push(headerRow);

        let grand = 0;
        for (const key of rows) {
          const value = sums.get(key) ?? 0;
          grand += value;
          const row: any[] = [key, value];
          if (includeColGrand) row.push(value);
          grid.push(row);
        }

        if (includeRowGrand) {
          const row: any[] = ["Grand Total", grand];
          if (includeColGrand) row.push(grand);
          grid.push(row);
        }

        const updates: any[] = [];
        for (let r = 0; r < grid.length; r += 1) {
          for (let c = 0; c < grid[r]!.length; c += 1) {
            const value = grid[r]![c];
            updates.push({
              sheet_id: destSheetId,
              row: destRow + r,
              col: destCol + c,
              value,
              formula: null,
              display_value: displayValue(value),
            });
          }
        }
        return updates;
      }

      (window as any).__TAURI__ = {
        core: {
          invoke: async (cmd: string, args: any) => {
            switch (cmd) {
              case "list_tables":
                return [];

              case "add_sheet":
                return { id: args?.name ?? "Sheet2", name: args?.name ?? "Sheet2" };

              case "list_pivot_tables":
                return pivotTables.map(({ id, name, source_sheet_id, source_range, destination }) => ({
                  id,
                  name,
                  source_sheet_id,
                  source_range,
                  destination,
                }));

              case "create_pivot_table": {
                const request = args?.request;
                if (!request) throw new Error("create_pivot_table missing request");
                const pivot_id = `pivot-${pivotTables.length + 1}`;
                (window as any).__pivotLastRequest = request;
                pivotTables.push({ id: pivot_id, ...request });
                return { pivot_id, updates: computePivotUpdates(request) };
              }

              case "refresh_pivot_table": {
                const pivotId = args?.request?.pivot_id;
                const pivot = pivotTables.find((p) => p.id === pivotId);
                if (!pivot) throw new Error(`unknown pivot: ${pivotId}`);
                return computePivotUpdates(pivot);
              }

              // Host sync calls (no-op in this test harness).
              case "set_cell":
              case "set_range":
              case "save_workbook":
              case "mark_saved":
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
    await page.evaluate(() => localStorage.clear());
    await page.reload({ waitUntil: "domcontentloaded" });
    await waitForDesktopReady(page);

    // Seed a small dataset.
    await page.evaluate(() => {
      const doc = (window as any).__formulaApp.getDocument();
      doc.beginBatch({ label: "seed pivot data" });
      doc.setRangeValues("Sheet1", "A1", [
        ["Category", "Amount"],
        ["A", 10],
        ["A", 20],
        ["B", 5],
      ]);
      doc.endBatch();

      (window as any).__formulaApp.selectRange({
        sheetId: "Sheet1",
        range: { startRow: 0, startCol: 0, endRow: 3, endCol: 1 },
      });
    });

    const modifier = process.platform === "darwin" ? "Meta" : "Control";
    await page.keyboard.press(`${modifier}+Shift+P`);
    await expect(page.getByTestId("command-palette-input")).toBeVisible();
    await page.getByTestId("command-palette-input").fill("pivot");
    // Bonus: verify category grouping renders in the list (stable UI affordance).
    await expect(page.getByTestId("command-palette-list")).toContainText("Insert");
    await page.keyboard.press("Enter");

    const panel = page.getByTestId("dock-left").getByTestId("panel-pivotBuilder");
    await expect(panel).toBeVisible();

    // Destination: existing sheet, starting at D1.
    await panel.getByTestId("pivot-destination-existing").click();
    await panel.getByTestId("pivot-destination-cell").fill("D1");

    // Configure fields via drag-and-drop.
    // Use a synthetic drop event for determinism (Playwright drag/drop can be flaky under load
    // when plumbing HTML5 DataTransfer through real pointer gestures).
    await page.evaluate(() => {
      const zone = document.querySelector('[data-testid="pivot-drop-rows"]');
      if (!zone) throw new Error("Missing pivot-drop-rows");
      const dt = new DataTransfer();
      dt.setData("text/plain", "Category");
      const drop = new DragEvent("drop", { bubbles: true, cancelable: true });
      Object.defineProperty(drop, "dataTransfer", { value: dt });
      zone.dispatchEvent(drop);
    });
    await expect(panel.getByTestId("pivot-drop-rows")).toContainText("Category");

    await page.evaluate(() => {
      const zone = document.querySelector('[data-testid="pivot-drop-values"]');
      if (!zone) throw new Error("Missing pivot-drop-values");
      const dt = new DataTransfer();
      dt.setData("text/plain", "Amount");
      const drop = new DragEvent("drop", { bubbles: true, cancelable: true });
      Object.defineProperty(drop, "dataTransfer", { value: dt });
      zone.dispatchEvent(drop);
    });
    await expect(panel.getByTestId("pivot-value-aggregation-0")).toBeVisible();

    // Disable column grand totals to keep the output shape deterministic for assertions.
    await panel.getByTestId("pivot-grand-totals-columns").uncheck();

    await panel.getByTestId("pivot-create").click();

    // Validate output.
    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("D1")), { timeout: 5_000 })
      .toBe("Category");
    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E2")), { timeout: 5_000 })
      .toBe("30");
    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E4")), { timeout: 5_000 })
      .toBe("35");

    // Update source data and refresh pivot.
    await page.evaluate(() => {
      const doc = (window as any).__formulaApp.getDocument();
      doc.setCellValue("Sheet1", "B2", 15);
    });

    await panel.getByTestId("pivot-refresh-pivot-1").click();

    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E2")), { timeout: 5_000 })
      .toBe("35");
    await expect
      .poll(() => page.evaluate(() => (window as any).__formulaApp.getCellValueA1("E4")), { timeout: 5_000 })
      .toBe("40");

    // Confirm config mapping made it into the backend request (grandTotals.columns disabled).
    const lastRequest = await page.evaluate(() => (window as any).__pivotLastRequest);
    expect(lastRequest?.config?.grandTotals?.columns).toBe(false);
  });
});
