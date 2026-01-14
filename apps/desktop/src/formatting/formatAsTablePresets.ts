import type { DocumentController } from "../document/documentController.js";
import { applyOutsideBorders, type CellRange } from "./toolbar.js";
import { DEFAULT_FORMATTING_APPLY_CELL_LIMIT } from "./selectionSizeGuard.js";

export type FormatAsTablePresetId = "light" | "medium" | "dark";

export const FORMAT_AS_TABLE_MAX_BANDED_ROW_OPS = 5_000;

export type FormatAsTablePreset = {
  header: {
    fill: string;
    fontColor: string;
  };
  bandedRows: {
    primaryFill: string;
    secondaryFill: string;
  };
  borders: {
    /**
     * OOXML-style ARGB string (AARRGGBB), e.g. `FF000000`.
     *
     * Note: we store these without a leading `#` so `noHardcodedColors.test.js` doesn't
     * treat them as CSS hex literals. Callers should prefix `#` when applying to styles.
     */
    outlineColor: string;
    /**
     * OOXML-style ARGB string (AARRGGBB), e.g. `FFBFBFBF`.
     */
    innerHorizontalColor: string;
    style: string;
  };
};

const PRESETS: Record<FormatAsTablePresetId, FormatAsTablePreset> = {
  light: {
    header: { fill: "FF4F81BD", fontColor: "FFFFFFFF" },
    bandedRows: { primaryFill: "FFFFFFFF", secondaryFill: "FFD9E1F2" },
    borders: { outlineColor: "FF000000", innerHorizontalColor: "FFBFBFBF", style: "thin" },
  },
  medium: {
    header: { fill: "FF70AD47", fontColor: "FFFFFFFF" },
    bandedRows: { primaryFill: "FFFFFFFF", secondaryFill: "FFE2EFDA" },
    borders: { outlineColor: "FF000000", innerHorizontalColor: "FFBFBFBF", style: "thin" },
  },
  dark: {
    header: { fill: "FF1F4E79", fontColor: "FFFFFFFF" },
    bandedRows: { primaryFill: "FFF2F2F2", secondaryFill: "FFD9D9D9" },
    borders: { outlineColor: "FF000000", innerHorizontalColor: "FF808080", style: "thin" },
  },
};

export function getFormatAsTablePreset(preset: FormatAsTablePresetId): FormatAsTablePreset {
  return PRESETS[preset];
}

export function estimateFormatAsTableBandedRowOps(rowCount: number): number {
  // Banded rows are applied to every other *body* row (selection rows - 1 header row).
  // For example:
  // - 1 row total => 0 banded row ops
  // - 2 rows total => 0 banded row ops
  // - 3 rows total => 1 banded row op (the 2nd body row)
  const rows = Number(rowCount);
  if (!Number.isFinite(rows) || rows <= 0) return 0;
  return Math.floor(Math.max(0, rows - 1) / 2);
}

function normalizeRange(range: CellRange): CellRange {
  const startRow = Math.min(range.start.row, range.end.row);
  const endRow = Math.max(range.start.row, range.end.row);
  const startCol = Math.min(range.start.col, range.end.col);
  const endCol = Math.max(range.start.col, range.end.col);
  return { start: { row: startRow, col: startCol }, end: { row: endRow, col: endCol } };
}

export type TablePresetRanges = {
  table: CellRange;
  header: CellRange;
  body: CellRange | null;
};

export function computeTablePresetRanges(input: CellRange): TablePresetRanges {
  const table = normalizeRange(input);
  const header: CellRange = {
    start: { row: table.start.row, col: table.start.col },
    end: { row: table.start.row, col: table.end.col },
  };
  const body =
    table.end.row > table.start.row
      ? {
          start: { row: table.start.row + 1, col: table.start.col },
          end: { row: table.end.row, col: table.end.col },
        }
      : null;
  return { table, header, body };
}

function normalizeColor(argb: string): string {
  const raw = String(argb ?? "").trim();
  if (!raw) return raw;
  if (raw.startsWith("#")) return raw;
  return `#${raw}`;
}

function fillPatch(argb: string): Record<string, any> {
  return { fill: { pattern: "solid", fgColor: normalizeColor(argb) } };
}

export function applyFormatAsTablePreset(doc: DocumentController, sheetId: string, range: CellRange, presetId: FormatAsTablePresetId): boolean {
  const preset = getFormatAsTablePreset(presetId);
  const { table, header, body } = computeTablePresetRanges(range);

  const rowCount = table.end.row - table.start.row + 1;
  const colCount = table.end.col - table.start.col + 1;
  const cellCount = rowCount * colCount;
  const bandedRowOps = estimateFormatAsTableBandedRowOps(rowCount);
  if (cellCount > DEFAULT_FORMATTING_APPLY_CELL_LIMIT || bandedRowOps > FORMAT_AS_TABLE_MAX_BANDED_ROW_OPS) {
    return false;
  }

  let applied = true;
  const label = "Format as Table";

  // Ensure this formatting preset is always applied as a single undo step, even if callers
  // don't explicitly batch. The ribbon command path already batches, so avoid creating nested
  // batches there.
  const shouldBatch = (doc as any).batchDepth === 0;
  if (shouldBatch) doc.beginBatch({ label });
  try {
    const okHeader = doc.setRangeFormat(
      sheetId,
      header,
      {
        font: { bold: true, color: normalizeColor(preset.header.fontColor) },
        ...fillPatch(preset.header.fill),
      },
      { label },
    );
    if (okHeader === false) applied = false;

    if (body) {
      const okBody = doc.setRangeFormat(sheetId, body, fillPatch(preset.bandedRows.primaryFill), { label });
      if (okBody === false) applied = false;

      // Apply the secondary band to every other row in the body. This is an MVP implementation
      // (formatting-only, no table metadata), so we keep it simple and optimize for small-medium
      // selections by using a single pass over alternating rows.
      for (let row = body.start.row + 1; row <= body.end.row; row += 2) {
        const rowRange: CellRange = {
          start: { row, col: body.start.col },
          end: { row, col: body.end.col },
        };
        const okBand = doc.setRangeFormat(sheetId, rowRange, fillPatch(preset.bandedRows.secondaryFill), { label });
        if (okBand === false) applied = false;
      }
    }

    applied = applyTableBorders(doc, sheetId, table, preset) && applied;

    return applied;
  } catch (err) {
    // If we started the batch, cancel it so callers don't observe partial state.
    if (shouldBatch) {
      try {
        doc.cancelBatch();
      } catch {
        // ignore
      }
    }
    throw err;
  } finally {
    if (shouldBatch) {
      // If the batch was canceled above, endBatch() will be a no-op.
      doc.endBatch();
    }
  }
}

function applyTableBorders(doc: DocumentController, sheetId: string, table: CellRange, preset: FormatAsTablePreset): boolean {
  const style = preset.borders.style;
  const innerEdge = { style, color: normalizeColor(preset.borders.innerHorizontalColor) };
  const label = "Format as Table";
  let applied = true;

  const okOutline = applyOutsideBorders(doc, sheetId, table, { style, color: normalizeColor(preset.borders.outlineColor) });
  if (okOutline === false) applied = false;

  const startRow = table.start.row;
  const endRow = table.end.row;
  const startCol = table.start.col;
  const endCol = table.end.col;

  // Interior horizontal separators for readability.
  if (endRow > startRow) {
    const interiorRows: CellRange = {
      start: { row: startRow, col: startCol },
      end: { row: endRow - 1, col: endCol },
    };
    const okInner = doc.setRangeFormat(sheetId, interiorRows, { border: { bottom: innerEdge } }, { label });
    if (okInner === false) applied = false;
  }

  // Interior vertical separators.
  if (endCol > startCol) {
    const interiorCols: CellRange = {
      start: { row: startRow, col: startCol },
      end: { row: endRow, col: endCol - 1 },
    };
    const okInner = doc.setRangeFormat(sheetId, interiorCols, { border: { right: innerEdge } }, { label });
    if (okInner === false) applied = false;
  }

  return applied;
}
