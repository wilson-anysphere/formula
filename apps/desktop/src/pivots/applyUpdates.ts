import type { DocumentController } from "../document/documentController.js";
import type { CellUpdate } from "../tauri/pivotBackend.js";

export function applyPivotCellUpdates(document: DocumentController, updates: readonly CellUpdate[]): void {
  if (!Array.isArray(updates) || updates.length === 0) return;

  const setCellInputs = (document as any).setCellInputs as
    | ((inputs: Array<{ sheetId: string; row: number; col: number; value: unknown; formula: string | null }>, options?: any) => void)
    | undefined;

  if (typeof setCellInputs === "function") {
    const inputs: Array<{ sheetId: string; row: number; col: number; value: unknown; formula: string | null }> = [];
    for (const update of updates) {
      const sheetId = String(update?.sheet_id ?? "").trim();
      if (!sheetId) continue;

      const row = Number(update?.row);
      const col = Number(update?.col);
      if (!Number.isInteger(row) || row < 0) continue;
      if (!Number.isInteger(col) || col < 0) continue;

      const trimmedFormula = typeof update.formula === "string" ? update.formula.trim() : "";
      const formula = trimmedFormula.length > 0 ? trimmedFormula : null;
      const value = formula ? null : (update.value ?? null);

      inputs.push({ sheetId, row, col, value, formula });
    }
    // Pivot updates originate in the backend pivot engine. Tag them so the desktop workbook sync
    // bridge does not echo the edits back via `set_cell` / `set_range`.
    // `setCellInputs` is a method on the DocumentController; ensure it is invoked with the
    // controller instance as its `this` context.
    setCellInputs.call(document, inputs, { source: "pivot" });
    return;
  }

  for (const update of updates) {
    const sheetId = String(update?.sheet_id ?? "").trim();
    if (!sheetId) continue;

    const row = Number(update?.row);
    const col = Number(update?.col);
    if (!Number.isInteger(row) || row < 0) continue;
    if (!Number.isInteger(col) || col < 0) continue;

    const trimmedFormula = typeof update.formula === "string" ? update.formula.trim() : "";
    const formula = trimmedFormula.length > 0 ? trimmedFormula : null;
    const value = formula ? null : (update.value ?? null);

    // Pivot updates originate in the backend pivot engine. Tag them so the desktop workbook sync
    // bridge does not echo the edits back via `set_cell` / `set_range`.
    document.setCellInput(sheetId, { row, col }, { value, formula }, { source: "pivot" });
  }
}
