import type { CellChange } from "./protocol.ts";
import { toA1 } from "./backend/a1.ts";
import { normalizeFormulaTextOpt } from "./backend/formula.ts";

export type EngineCellScalar = number | string | boolean | null;

export type EngineSheetJson = {
  /**
   * Sparse cell map keyed by A1 address.
   */
  cells: Record<string, EngineCellScalar>;
  /**
   * Optional logical worksheet dimensions (row count).
   *
   * When set, this controls how whole-column/row references like `A:A` / `1:1`
   * are expanded by the WASM engine.
   */
  rowCount?: number;
  /**
   * Optional logical worksheet dimensions (column count).
   */
  colCount?: number;
};

export type EngineWorkbookJson = {
  sheets: Record<string, EngineSheetJson>;
};

export type DocumentCellState = {
  value: unknown;
  formula: string | null;
  styleId: number;
};

export type DocumentCellDelta = {
  sheetId: string;
  row: number;
  col: number;
  before: DocumentCellState;
  after: DocumentCellState;
};

export interface EngineSyncTarget {
  loadWorkbookFromJson: (json: string) => Promise<void> | void;
  setCell: (address: string, value: EngineCellScalar, sheet?: string) => Promise<void> | void;
  setCells?: (
    updates: Array<{ address: string; value: EngineCellScalar; sheet?: string }>,
  ) => Promise<void> | void;
  recalculate: (sheet?: string) => Promise<CellChange[]> | CellChange[];
}

function parseRowColKey(key: string): { row: number; col: number } | null {
  const [rowStr, colStr] = key.split(",");
  const row = Number(rowStr);
  const col = Number(colStr);
  if (!Number.isInteger(row) || row < 0) return null;
  if (!Number.isInteger(col) || col < 0) return null;
  return { row, col };
}

function isRichTextValue(value: unknown): value is { text: string } {
  return Boolean(
    value &&
      typeof value === "object" &&
      "text" in value &&
      typeof (value as { text?: unknown }).text === "string",
  );
}

function coerceDocumentValueToScalar(value: unknown): EngineCellScalar | null {
  if (value == null) return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return value;
  if (isRichTextValue(value)) return value.text;
  return null;
}

function cellStateToEngineInput(cell: DocumentCellState): EngineCellScalar | null {
  if (typeof cell.formula === "string") {
    const normalized = normalizeFormulaTextOpt(cell.formula);
    if (normalized != null) return normalized;
  }
  return coerceDocumentValueToScalar(cell.value);
}

/**
 * Export the current DocumentController workbook into the JSON format consumed by
 * `crates/formula-wasm` (`WasmWorkbook.fromJson`).
 *
 * Note: empty/cleared cells are omitted from the JSON entirely (sparse semantics).
 */
export function exportDocumentToEngineWorkbookJson(doc: any): EngineWorkbookJson {
  const sheets: Record<string, EngineSheetJson> = {};

  const sheetIds: string[] =
    typeof doc?.getSheetIds === "function" ? (doc.getSheetIds() as string[]) : [];
  const ids = sheetIds.length > 0 ? sheetIds : ["Sheet1"];

  for (const sheetId of ids) {
    const cells: Record<string, EngineCellScalar> = {};
    const sheet = doc?.model?.sheets?.get?.(sheetId);

    if (sheet?.cells?.entries) {
      for (const [key, cell] of sheet.cells.entries() as Iterable<[string, DocumentCellState]>) {
        const coord = parseRowColKey(key);
        if (!coord) continue;

        const input = cellStateToEngineInput(cell);
        if (input == null) continue;

        const address = toA1(coord.row, coord.col);
        cells[address] = input;
      }
    }

    sheets[sheetId] = { cells };
  }

  return { sheets };
}

export async function engineHydrateFromDocument(engine: EngineSyncTarget, doc: any): Promise<CellChange[]> {
  const workbookJson = exportDocumentToEngineWorkbookJson(doc);
  await engine.loadWorkbookFromJson(JSON.stringify(workbookJson));
  return await engine.recalculate();
}

export async function engineApplyDeltas(
  engine: EngineSyncTarget,
  deltas: readonly DocumentCellDelta[],
  options: { recalculate?: boolean } = {},
): Promise<CellChange[]> {
  const shouldRecalculate = options.recalculate ?? true;
  const updates: Array<{ address: string; value: EngineCellScalar; sheet?: string }> = [];

  for (const delta of deltas) {
    const beforeInput = cellStateToEngineInput(delta.before);
    const afterInput = cellStateToEngineInput(delta.after);

    // Ignore formatting-only edits and rich-text run edits that don't change the plain input.
    if (beforeInput === afterInput) continue;

    const address = toA1(delta.row, delta.col);
    updates.push({ address, value: afterInput, sheet: delta.sheetId });
  }

  if (updates.length === 0) return [];

  if (engine.setCells) {
    await engine.setCells(updates);
  } else {
    await Promise.all(updates.map((u) => engine.setCell(u.address, u.value, u.sheet)));
  }
  if (!shouldRecalculate) return [];
  return await engine.recalculate();
}
