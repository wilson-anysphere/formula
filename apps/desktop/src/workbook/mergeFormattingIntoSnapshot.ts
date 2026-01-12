export type SnapshotCell = {
  row: number;
  col: number;
  value: unknown | null;
  formula: string | null;
  format: unknown | null;
};

export type CellFormatClampBounds = Readonly<{
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
}>;

export type SheetFormattingSnapshot = Readonly<{
  defaultFormat?: unknown | null;
  rowFormats?: unknown;
  colFormats?: unknown;
  formatRunsByCol?: unknown;
  cellFormats?: unknown;
}>;

type CellFormatEntry = Readonly<{
  row: number;
  col: number;
  format: unknown | null;
}>;

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function normalizeCellFormatEntries(raw: unknown): CellFormatEntry[] {
  if (!raw) return [];

  if (Array.isArray(raw)) {
    const out: CellFormatEntry[] = [];
    for (const entry of raw) {
      const row = Number((entry as any)?.row);
      const col = Number((entry as any)?.col);
      if (!Number.isInteger(row) || row < 0) continue;
      if (!Number.isInteger(col) || col < 0) continue;
      out.push({ row, col, format: (entry as any)?.format ?? null });
    }
    return out;
  }

  if (isPlainObject(raw)) {
    // Allow `{ "row,col": format }` or `{ "row,col": { format: ... } }` shapes.
    const out: CellFormatEntry[] = [];
    for (const [key, value] of Object.entries(raw)) {
      const comma = key.indexOf(",");
      if (comma === -1) continue;
      const row = Number(key.slice(0, comma));
      const col = Number(key.slice(comma + 1));
      if (!Number.isInteger(row) || row < 0) continue;
      if (!Number.isInteger(col) || col < 0) continue;
      const format = isPlainObject(value) && "format" in value ? (value as any).format : value;
      out.push({ row, col, format: (format as any) ?? null });
    }
    return out;
  }

  return [];
}

function withinBounds(row: number, col: number, bounds: CellFormatClampBounds): boolean {
  return row >= bounds.startRow && row <= bounds.endRow && col >= bounds.startCol && col <= bounds.endCol;
}

/**
 * Merges a per-sheet persisted formatting snapshot into the sheet snapshot used by
 * `SpreadsheetApp.restoreDocumentState`.
 *
 * - Formats for cells already present in `cells` are attached to those entries.
 * - Format-only cells are added when they have a non-null format.
 * - When `clampCellFormatsTo` is provided, `cellFormats` entries outside those bounds
 *   are ignored (to match workbook-open truncation behavior).
 */
export function mergeFormattingIntoSnapshot(options: Readonly<{
  cells: SnapshotCell[];
  formatting: SheetFormattingSnapshot | null | undefined;
  clampCellFormatsTo?: CellFormatClampBounds | null;
}>): {
  cells: SnapshotCell[];
  defaultFormat?: unknown | null;
  rowFormats?: unknown;
  colFormats?: unknown;
  formatRunsByCol?: unknown;
} {
  const formatting = options.formatting ?? null;
  if (!formatting) return { cells: options.cells };

  const clampBounds = options.clampCellFormatsTo ?? null;

  const cellFormats = normalizeCellFormatEntries(formatting.cellFormats);

  // Fast path: no per-cell formatting to merge.
  if (cellFormats.length === 0) {
    const out: {
      cells: SnapshotCell[];
      defaultFormat?: unknown | null;
      rowFormats?: unknown;
      colFormats?: unknown;
      formatRunsByCol?: unknown;
    } = { cells: options.cells };

    if ("defaultFormat" in formatting) out.defaultFormat = formatting.defaultFormat ?? null;
    if ("rowFormats" in formatting) out.rowFormats = formatting.rowFormats;
    if ("colFormats" in formatting) out.colFormats = formatting.colFormats;
    if ("formatRunsByCol" in formatting) out.formatRunsByCol = formatting.formatRunsByCol;
    return out;
  }

  const cellsByKey = new Map<string, SnapshotCell>();
  const outCells: SnapshotCell[] = [];

  for (const cell of options.cells) {
    const row = Number((cell as any)?.row);
    const col = Number((cell as any)?.col);
    if (!Number.isInteger(row) || row < 0) continue;
    if (!Number.isInteger(col) || col < 0) continue;
    const key = `${row},${col}`;
    if (cellsByKey.has(key)) continue;
    cellsByKey.set(key, cell);
    outCells.push(cell);
  }

  for (const entry of cellFormats) {
    if (clampBounds && !withinBounds(entry.row, entry.col, clampBounds)) continue;

    const key = `${entry.row},${entry.col}`;
    const existing = cellsByKey.get(key);
    if (existing) {
      existing.format = entry.format;
      continue;
    }

    if (entry.format == null) continue;
    const newCell: SnapshotCell = { row: entry.row, col: entry.col, value: null, formula: null, format: entry.format };
    cellsByKey.set(key, newCell);
    outCells.push(newCell);
  }

  const out: {
    cells: SnapshotCell[];
    defaultFormat?: unknown | null;
    rowFormats?: unknown;
    colFormats?: unknown;
    formatRunsByCol?: unknown;
  } = { cells: outCells };

  if ("defaultFormat" in formatting) out.defaultFormat = formatting.defaultFormat ?? null;
  if ("rowFormats" in formatting) out.rowFormats = formatting.rowFormats;
  if ("colFormats" in formatting) out.colFormats = formatting.colFormats;
  if ("formatRunsByCol" in formatting) out.formatRunsByCol = formatting.formatRunsByCol;

  return out;
}
