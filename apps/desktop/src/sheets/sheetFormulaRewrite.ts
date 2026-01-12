import { rewriteDeletedSheetReferencesInFormula, rewriteSheetNamesInFormula } from "../workbook/formulaRewrite";

export { rewriteDeletedSheetReferencesInFormula };

type DocumentControllerLike = {
  setCellInputs?: (
    inputs: ReadonlyArray<{ sheetId: string; row: number; col: number; value: unknown; formula: string | null }>,
    options?: { source?: string; label?: string },
  ) => void;
  // The real DocumentController has a `model` property with a sparse sheet map.
  // This helper intentionally uses an `any` view so it can operate on the runtime
  // JS implementation without importing internal types.
  model?: any;
};

function parseRowColKey(key: string): { row: number; col: number } | null {
  const comma = key.indexOf(",");
  if (comma === -1) return null;
  const row = Number(key.slice(0, comma));
  const col = Number(key.slice(comma + 1));
  if (!Number.isInteger(row) || row < 0) return null;
  if (!Number.isInteger(col) || col < 0) return null;
  return { row, col };
}

function collectFormulaEdits(params: {
  doc: DocumentControllerLike;
  rewrite: (formula: string) => string;
}): Array<{ sheetId: string; row: number; col: number; value: null; formula: string }> {
  const { doc, rewrite } = params;
  /** @type {Array<{ sheetId: string; row: number; col: number; value: null; formula: string }>} */
  const edits: Array<{ sheetId: string; row: number; col: number; value: null; formula: string }> = [];

  const sheets: Map<string, any> | undefined = doc?.model?.sheets;
  if (!sheets || typeof sheets.entries !== "function") return edits;

  for (const [sheetId, sheet] of sheets.entries()) {
    const cells: Map<string, any> | undefined = sheet?.cells;
    if (!cells || typeof cells.entries !== "function") continue;

    for (const [key, cell] of cells.entries()) {
      const formula = cell?.formula;
      if (formula == null) continue;
      const next = rewrite(String(formula));
      if (next === formula) continue;
      const coord = parseRowColKey(String(key));
      if (!coord) continue;
      edits.push({ sheetId, row: coord.row, col: coord.col, value: null, formula: next });
    }
  }

  return edits;
}

export function rewriteDocumentFormulasForSheetRename(doc: DocumentControllerLike, oldName: string, newName: string): void {
  const edits = collectFormulaEdits({
    doc,
    rewrite: (formula) => rewriteSheetNamesInFormula(formula, oldName, newName),
  });
  if (edits.length === 0) return;
  if (typeof doc.setCellInputs !== "function") return;

  doc.setCellInputs(edits, { label: "Rename Sheet", source: "sheetRename" });
}

export function rewriteDocumentFormulasForSheetDelete(doc: DocumentControllerLike, deletedName: string, sheetOrder: string[]): void {
  const edits = collectFormulaEdits({
    doc,
    rewrite: (formula) => rewriteDeletedSheetReferencesInFormula(formula, deletedName, sheetOrder),
  });
  if (edits.length === 0) return;
  if (typeof doc.setCellInputs !== "function") return;

  doc.setCellInputs(edits, { label: "Delete Sheet", source: "sheetDelete" });
}
