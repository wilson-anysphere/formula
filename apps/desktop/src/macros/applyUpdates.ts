import type { DocumentController } from "../document/documentController.js";
import type { MacroCellUpdate } from "./types";

export function applyMacroCellUpdates(document: DocumentController, updates: readonly MacroCellUpdate[]): void {
  if (!Array.isArray(updates) || updates.length === 0) return;

  for (const update of updates) {
    const sheetId = String(update?.sheetId ?? "").trim();
    if (!sheetId) continue;

    const row = Number(update?.row);
    const col = Number(update?.col);
    if (!Number.isInteger(row) || row < 0) continue;
    if (!Number.isInteger(col) || col < 0) continue;

    const trimmedFormula = typeof update.formula === "string" ? update.formula.trim() : "";
    const formula = trimmedFormula.length > 0 ? trimmedFormula : null;
    const value = formula ? null : (update.value ?? null);

    document.setCellInput(sheetId, { row, col }, { value, formula });
  }
}
