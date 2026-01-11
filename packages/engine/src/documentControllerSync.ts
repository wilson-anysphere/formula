export type EngineCellScalar = number | string | boolean | null;

export type EngineWorkbookJson = {
  sheets: Record<string, { cells: Record<string, EngineCellScalar> }>;
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
  recalculate: (sheet?: string) => Promise<unknown> | unknown;
}

function colToName(col0: number): string {
  let n = col0 + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function toA1(row0: number, col0: number): string {
  return `${colToName(col0)}${row0 + 1}`;
}

function parseRowColKey(key: string): { row: number; col: number } | null {
  const [rowStr, colStr] = key.split(",");
  const row = Number(rowStr);
  const col = Number(colStr);
  if (!Number.isInteger(row) || row < 0) return null;
  if (!Number.isInteger(col) || col < 0) return null;
  return { row, col };
}

function normalizeFormulaText(formula: string): string {
  const trimmed = formula.trimStart();
  if (trimmed.startsWith("=")) return trimmed;
  return `=${trimmed}`;
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
  if (typeof cell.formula === "string" && cell.formula.trim() !== "") {
    return normalizeFormulaText(cell.formula);
  }
  return coerceDocumentValueToScalar(cell.value);
}

/**
 * Export the current DocumentController workbook into the JSON format consumed by
 * `crates/formula-core` / `WasmWorkbook.fromJson`.
 */
export function exportDocumentToEngineWorkbookJson(doc: any): EngineWorkbookJson {
  const sheets: Record<string, { cells: Record<string, EngineCellScalar> }> = {};

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

export async function engineHydrateFromDocument(engine: EngineSyncTarget, doc: any): Promise<void> {
  const workbookJson = exportDocumentToEngineWorkbookJson(doc);
  await engine.loadWorkbookFromJson(JSON.stringify(workbookJson));
  await engine.recalculate();
}

export async function engineApplyDeltas(
  engine: EngineSyncTarget,
  deltas: readonly DocumentCellDelta[],
): Promise<void> {
  const updates: Array<{ sheetId: string; address: string; input: EngineCellScalar }> = [];

  for (const delta of deltas) {
    const beforeInput = cellStateToEngineInput(delta.before);
    const afterInput = cellStateToEngineInput(delta.after);

    // Ignore formatting-only edits and rich-text run edits that don't change the plain input.
    if (beforeInput === afterInput) continue;

    const address = toA1(delta.row, delta.col);
    updates.push({ sheetId: delta.sheetId, address, input: afterInput });
  }

  if (updates.length === 0) return;

  await Promise.all(updates.map((u) => engine.setCell(u.address, u.input, u.sheetId)));
  await engine.recalculate();
}
