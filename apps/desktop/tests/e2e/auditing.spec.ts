import { expect, test } from "@playwright/test";

import { gotoDesktop } from "./helpers";

async function waitForIdle(page: import("@playwright/test").Page): Promise<void> {
  // Vite may occasionally trigger a one-time full reload after dependency optimization.
  // Retry once if the execution context is destroyed mid-wait.
  for (let attempt = 0; attempt < 2; attempt += 1) {
    try {
      await page.waitForFunction(() => Boolean((window.__formulaApp as any)?.whenIdle), null, { timeout: 10_000 });
      await page.evaluate(() => (window.__formulaApp as any).whenIdle());
      return;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (attempt === 0 && message.includes("Execution context was destroyed")) {
        await page.waitForLoadState("domcontentloaded");
        continue;
      }
      throw err;
    }
  }
}

test.describe("formula auditing overlays", () => {
  test.setTimeout(60_000);

  test("trace precedents/dependents uses backend dependency graph", async ({ page }) => {
    await page.addInitScript(() => {
      const listeners: Record<string, any> = {};
      (window as any).__tauriListeners = listeners;

      const formulas = new Map<string, string>();
      const precedentsByCell = new Map<string, Array<{ sheetId: string; row: number; col: number }>>();

      const cellKey = (sheetId: string, row: number, col: number) => `${sheetId}:${row},${col}`;

      const colName = (col0: number) => {
        let col = Math.floor(col0) + 1;
        let out = "";
        while (col > 0) {
          const rem = (col - 1) % 26;
          out = String.fromCharCode(65 + rem) + out;
          col = Math.floor((col - 1) / 26);
        }
        return out;
      };

      const toA1 = (row0: number, col0: number) => `${colName(col0)}${row0 + 1}`;

      const nameToCol = (name: string) => {
        const up = name.toUpperCase();
        if (!/^[A-Z]+$/.test(up)) return null;
        let col = 0;
        for (let i = 0; i < up.length; i++) {
          col = col * 26 + (up.charCodeAt(i) - 64);
        }
        return col - 1;
      };

      const parseRefs = (sheetId: string, formula: string) => {
        const text = formula.startsWith("=") ? formula.slice(1) : formula;
        const refs: Array<{ sheetId: string; row: number; col: number }> = [];
        const re = /([A-Za-z]+)(\d+)/g;
        let m: RegExpExecArray | null;
        while ((m = re.exec(text))) {
          const col = nameToCol(m[1] ?? "");
          const row1 = Number(m[2]);
          if (col == null) continue;
          if (!Number.isFinite(row1) || row1 <= 0) continue;
          refs.push({ sheetId, row: row1 - 1, col });
        }
        return refs;
      };

      const updateCell = (sheetId: string, row: number, col: number, formula: string | null) => {
        const key = cellKey(sheetId, row, col);
        if (typeof formula === "string" && formula.trim() !== "") {
          formulas.set(key, formula);
          precedentsByCell.set(key, parseRefs(sheetId, formula));
        } else {
          formulas.delete(key);
          precedentsByCell.delete(key);
        }
      };

      const directPrecedents = (sheetId: string, row: number, col: number) => {
        const refs = precedentsByCell.get(cellKey(sheetId, row, col)) ?? [];
        const out = new Set<string>();
        for (const ref of refs) out.add(toA1(ref.row, ref.col));
        return Array.from(out);
      };

      const transitivePrecedents = (sheetId: string, row: number, col: number) => {
        const start = cellKey(sheetId, row, col);
        const visited = new Set<string>();
        const out = new Set<string>();
        const queue: string[] = [start];
        visited.add(start);

        while (queue.length > 0) {
          const cur = queue.shift()!;
          const refs = precedentsByCell.get(cur) ?? [];
          for (const ref of refs) {
            out.add(toA1(ref.row, ref.col));
            const refKey = cellKey(ref.sheetId, ref.row, ref.col);
            if (!visited.has(refKey) && precedentsByCell.has(refKey)) {
              visited.add(refKey);
              queue.push(refKey);
            }
          }
        }

        return Array.from(out);
      };

      const directDependents = (sheetId: string, row: number, col: number) => {
        const target = cellKey(sheetId, row, col);
        const out = new Set<string>();
        for (const [cell, refs] of precedentsByCell.entries()) {
          if (!refs.some((r) => cellKey(r.sheetId, r.row, r.col) === target)) continue;
          const [, rc] = cell.split(":");
          const [r, c] = (rc ?? "").split(",").map((n) => Number(n));
          if (Number.isFinite(r) && Number.isFinite(c)) out.add(toA1(r, c));
        }
        return Array.from(out);
      };

      const transitiveDependents = (sheetId: string, row: number, col: number) => {
        const start = cellKey(sheetId, row, col);
        const visited = new Set<string>();
        const out = new Set<string>();
        const queue: string[] = [start];
        visited.add(start);

        while (queue.length > 0) {
          const cur = queue.shift()!;
          const [, rc] = cur.split(":");
          const [r, c] = (rc ?? "").split(",").map((n) => Number(n));
          if (!Number.isFinite(r) || !Number.isFinite(c)) continue;

          for (const depA1 of directDependents(sheetId, r, c)) {
            const match = /^([A-Z]+)(\d+)$/.exec(depA1);
            if (!match) continue;
            const depCol = nameToCol(match[1] ?? "");
            const depRow = Number(match[2]) - 1;
            if (depCol == null || !Number.isFinite(depRow)) continue;

            const depKey = cellKey(sheetId, depRow, depCol);
            if (visited.has(depKey)) continue;
            visited.add(depKey);
            out.add(depA1);
            queue.push(depKey);
          }
        }

        return Array.from(out);
      };

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

              case "set_cell": {
                const sheetId = String(args?.sheet_id ?? "Sheet1");
                const row = Number(args?.row ?? 0);
                const col = Number(args?.col ?? 0);
                const formula = typeof args?.formula === "string" ? String(args.formula) : null;
                updateCell(sheetId, row, col, formula);
                return null;
              }

              case "set_range": {
                const sheetId = String(args?.sheet_id ?? "Sheet1");
                const startRow = Number(args?.start_row ?? 0);
                const startCol = Number(args?.start_col ?? 0);
                const values = Array.isArray(args?.values) ? args.values : [];
                for (let r = 0; r < values.length; r += 1) {
                  const rowVals = Array.isArray(values[r]) ? values[r] : [];
                  for (let c = 0; c < rowVals.length; c += 1) {
                    const cell = rowVals[c] as any;
                    const formula = typeof cell?.formula === "string" ? String(cell.formula) : null;
                    updateCell(sheetId, startRow + r, startCol + c, formula);
                  }
                }
                return null;
              }

              case "get_precedents": {
                const sheetId = String(args?.sheet_id ?? "Sheet1");
                const row = Number(args?.row ?? 0);
                const col = Number(args?.col ?? 0);
                const transitive = Boolean(args?.transitive);
                const out = transitive ? transitivePrecedents(sheetId, row, col) : directPrecedents(sheetId, row, col);
                out.sort();
                return out;
              }

              case "get_dependents": {
                const sheetId = String(args?.sheet_id ?? "Sheet1");
                const row = Number(args?.row ?? 0);
                const col = Number(args?.col ?? 0);
                const transitive = Boolean(args?.transitive);
                const out = transitive ? transitiveDependents(sheetId, row, col) : directDependents(sheetId, row, col);
                out.sort();
                return out;
              }

              case "mark_saved":
              case "save_workbook":
              case "recalculate":
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
            hide: async () => {},
            close: async () => {},
          }),
        },
      };

      (window as any).__auditingStub = { formulas };
    });

    await gotoDesktop(page, "/", { waitForIdle: false, waitForContextMenu: false });

    await page.waitForFunction(() => Boolean((window as any).__tauriListeners?.["file-dropped"]));
    await page.evaluate(() => {
      (window as any).__tauriListeners["file-dropped"]({ payload: ["/tmp/fake.xlsx"] });
    });

    await page.waitForFunction(async () => (await (window.__formulaApp as any).getCellValueA1("A1")) === "Hello");
    await waitForIdle(page);

    await page.click("#grid", { position: { x: 60, y: 40 } });

    const editor = page.locator("textarea.cell-editor");

    await page.keyboard.press("F2");
    await expect(editor).toBeVisible();
    await editor.fill("1");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await page.keyboard.press("ArrowUp");
    await page.keyboard.press("ArrowRight");
    await page.keyboard.press("F2");
    await expect(editor).toBeVisible();
    await editor.fill("=A1+1");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await page.keyboard.press("ArrowUp");
    await page.keyboard.press("ArrowRight");
    await page.keyboard.press("F2");
    await expect(editor).toBeVisible();
    await editor.fill("=B1+1");
    await page.keyboard.press("Enter");
    await waitForIdle(page);

    await page.keyboard.press("ArrowUp");
    await page.keyboard.press("ArrowLeft");

    await page.getByTestId("ribbon-root").getByTestId("audit-precedents").click();
    await page.getByTestId("ribbon-root").getByTestId("audit-dependents").click();
    await waitForIdle(page);

    await expect(page.getByTestId("active-cell")).toHaveText("B1");

    const highlightsB1 = await page.evaluate(() => (window.__formulaApp as any).getAuditingHighlights());
    expect(highlightsB1.mode).toBe("both");
    expect(highlightsB1.precedents).toEqual(["A1"]);
    expect(highlightsB1.dependents).toEqual(["C1"]);

    await page.keyboard.press("ArrowRight");
    await waitForIdle(page);
    const highlightsC1 = await page.evaluate(() => (window.__formulaApp as any).getAuditingHighlights());
    expect(highlightsC1.precedents).toEqual(["B1"]);
    expect(highlightsC1.dependents).toEqual([]);

    await page.getByTestId("ribbon-root").getByTestId("audit-transitive").click();
    await waitForIdle(page);
    const highlightsC1Transitive = await page.evaluate(() => (window.__formulaApp as any).getAuditingHighlights());
    expect(highlightsC1Transitive.transitive).toBe(true);
    expect(highlightsC1Transitive.precedents.sort()).toEqual(["A1", "B1"]);

    // Ribbon wiring: Formulas → Formula Auditing commands should mirror the same spreadsheet
    // auditing capabilities (precedents/dependents + clear).
    await page.keyboard.press("ArrowLeft"); // back to B1 (has both precedents and dependents)
    await waitForIdle(page);

    const formulasTab = page.getByRole("tab", { name: "Formulas", exact: true });
    await expect(formulasTab).toBeVisible();
    await formulasTab.click();

    await page.locator('button[data-command-id="formulas.formulaAuditing.tracePrecedents"]').click();
    await waitForIdle(page);
    const ribbonPrecedents = await page.evaluate(() => (window.__formulaApp as any).getAuditingHighlights());
    expect(ribbonPrecedents.mode).toBe("precedents");
    expect(ribbonPrecedents.precedents).toEqual(["A1"]);
    expect(ribbonPrecedents.dependents).toEqual([]);

    await page.locator('button[data-command-id="formulas.formulaAuditing.traceDependents"]').click();
    await waitForIdle(page);
    const ribbonDependents = await page.evaluate(() => (window.__formulaApp as any).getAuditingHighlights());
    expect(ribbonDependents.mode).toBe("dependents");
    expect(ribbonDependents.precedents).toEqual([]);
    expect(ribbonDependents.dependents).toEqual(["C1"]);

    await page.locator('button[data-command-id="formulas.formulaAuditing.removeArrows"]').click();
    await waitForIdle(page);
    const ribbonCleared = await page.evaluate(() => (window.__formulaApp as any).getAuditingHighlights());
    expect(ribbonCleared.mode).toBe("off");
    expect(ribbonCleared.precedents).toEqual([]);
    expect(ribbonCleared.dependents).toEqual([]);
  });
});
