import type { DocumentController } from "../document/documentController.js";
import type { Range } from "../selection/types";
import { getStyleFontSizePt, getStyleNumberFormat, getStyleWrapText } from "../formatting/styleFieldAccess.js";

export type SelectionHorizontalAlign = "left" | "center" | "right" | "mixed";

export type SelectionVerticalAlign = "top" | "center" | "bottom" | "mixed";

export type SelectionNumberFormat = string | "mixed" | null;

export type SelectionFontName = string | "mixed" | null;
export type SelectionFontSize = number | "mixed" | null;
export type SelectionFontVariantPosition = "subscript" | "superscript" | "mixed" | null;

export type SelectionFormatState = {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  fontName: SelectionFontName;
  fontSize: SelectionFontSize;
  fontVariantPosition: SelectionFontVariantPosition;
  wrapText: boolean;
  align: SelectionHorizontalAlign;
  verticalAlign: SelectionVerticalAlign;
  numberFormat: SelectionNumberFormat;
};

type NormalizedRange = { startRow: number; endRow: number; startCol: number; endCol: number };

function normalizeRange(range: Range): NormalizedRange {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return { startRow, endRow, startCol, endCol };
}

function safeCellCount(range: NormalizedRange): number {
  const rows = range.endRow - range.startRow + 1;
  const cols = range.endCol - range.startCol + 1;
  return rows * cols;
}

function sampleAxisIndices(start: number, end: number, maxSamples: number): number[] {
  const len = end - start + 1;
  if (len <= 0) return [];
  if (len <= maxSamples) {
    return Array.from({ length: len }, (_, i) => start + i);
  }
  if (maxSamples <= 1) return [start];

  const out: number[] = [];
  for (let i = 0; i < maxSamples; i++) {
    const idx = start + Math.floor((i * (len - 1)) / (maxSamples - 1));
    if (out[out.length - 1] !== idx) out.push(idx);
  }
  return out;
}

function* sampleRangeCells(range: NormalizedRange, maxCells: number): Generator<{ row: number; col: number }, void> {
  if (maxCells <= 0) return;

  const rows = range.endRow - range.startRow + 1;
  const cols = range.endCol - range.startCol + 1;
  if (rows <= 0 || cols <= 0) return;

  // Sampling strategy:
  // - For small ranges, enumerate every cell.
  // - For large ranges, sample a coarse grid across the full area.
  //   This intentionally does *not* try to be exhaustive; callers can treat
  //   results as best-effort.
  const total = rows * cols;
  if (total <= maxCells) {
    for (let row = range.startRow; row <= range.endRow; row++) {
      for (let col = range.startCol; col <= range.endCol; col++) {
        yield { row, col };
      }
    }
    return;
  }

  const approxPerAxis = Math.max(2, Math.floor(Math.sqrt(maxCells)));
  const rowSamples = sampleAxisIndices(range.startRow, range.endRow, approxPerAxis);
  const colSamples = sampleAxisIndices(range.startCol, range.endCol, approxPerAxis);

  let emitted = 0;
  for (const row of rowSamples) {
    for (const col of colSamples) {
      yield { row, col };
      emitted += 1;
      if (emitted >= maxCells) return;
    }
  }
}

type AggregationState = {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  fontName: string | null | "mixed" | undefined;
  fontSize: number | null | "mixed" | undefined;
  fontVariantPosition: "subscript" | "superscript" | "mixed" | null | undefined;
  wrapText: boolean;
  align: "left" | "center" | "right" | "mixed" | null;
  verticalAlign: "top" | "center" | "bottom" | "mixed" | null;
  numberFormat: string | null | "mixed" | undefined;
  inspected: number;
  exhaustive: boolean;
};

/**
 * Compute a lightweight formatting state summary for the current selection.
 *
 * The intent is to drive Ribbon toggle state (Bold/Italic/Underline/etc.) without
 * iterating every cell in very large selections.
 *
 * Performance semantics:
 * - If the total selection cell count is <= `maxInspectCells`, all cells are inspected.
 * - Otherwise, we inspect up to `maxInspectCells` sampled cells across the selection.
 * - For properties that require certainty, callers can treat sampled results as "mixed".
 */
export function computeSelectionFormatState(
  doc: DocumentController,
  sheetId: string,
  selectionRanges: Range[],
  options: { maxInspectCells?: number } = {},
): SelectionFormatState {
  const maxInspectCells = options.maxInspectCells ?? 256;

  const ranges = selectionRanges.map(normalizeRange).filter((r) => safeCellCount(r) > 0);
  if (ranges.length === 0) {
    return {
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      fontName: null,
      fontSize: null,
      fontVariantPosition: null,
      wrapText: false,
      align: "left",
      verticalAlign: "bottom",
      numberFormat: null,
    };
  }

  // Avoid materializing "phantom" sheets when a stale sheet id is passed in (e.g. after a sheet
  // delete/applyState). DocumentController lazily creates sheets when referenced by
  // `getCellFormat()` / `getCell()`, but this function is called frequently to drive Ribbon UI
  // state and should be side-effect free.
  //
  // Treat a missing sheet as having default formatting.
  const anyDoc = doc as any;
  const model = anyDoc?.model;
  const sheetMap = model?.sheets;
  const sheetMeta = anyDoc?.sheetMeta;
  const canCheckExists =
    Boolean(sheetMap && typeof sheetMap.has === "function") || Boolean(sheetMeta && typeof sheetMeta.has === "function");
  if (canCheckExists) {
    const exists =
      (sheetMap && typeof sheetMap.has === "function" && sheetMap.has(sheetId)) ||
      (sheetMeta && typeof sheetMeta.has === "function" && sheetMeta.has(sheetId));
    if (!exists) {
      return {
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontName: null,
        fontSize: null,
        fontVariantPosition: null,
        wrapText: false,
        align: "left",
        verticalAlign: "bottom",
        numberFormat: null,
      };
    }
  }

  // Determine whether we can inspect the selection exhaustively.
  let totalCells = 0;
  for (const r of ranges) {
    totalCells += safeCellCount(r);
    if (totalCells > maxInspectCells) break;
  }
  const exhaustive = totalCells <= maxInspectCells;

  /** @type {AggregationState} */
  const state: AggregationState = {
    bold: true,
    italic: true,
    underline: true,
    strikethrough: true,
    fontName: undefined,
    fontSize: undefined,
    fontVariantPosition: undefined,
    wrapText: true,
    align: null,
    verticalAlign: null,
    numberFormat: undefined,
    inspected: 0,
    exhaustive,
  };

  const visited = new Set<string>();
  const hasOwn = (obj: unknown, key: string): boolean => {
    if (!obj || typeof obj !== "object") return false;
    return Object.prototype.hasOwnProperty.call(obj, key);
  };

  const mergeAlign = (raw: unknown) => {
    const value = raw === "center" || raw === "right" || raw === "left" ? (raw as "left" | "center" | "right") : "left";
    if (state.align == null) state.align = value;
    else if (state.align !== "mixed" && state.align !== value) state.align = "mixed";
  };

  const mergeVerticalAlign = (raw: unknown) => {
    const normalized = typeof raw === "string" ? raw.toLowerCase() : "";
    // Excel/OOXML uses `center` for the vertical "middle" alignment option.
    const value: "top" | "center" | "bottom" =
      normalized === "top" ? "top" : normalized === "center" || normalized === "middle" ? "center" : "bottom";
    if (state.verticalAlign == null) state.verticalAlign = value;
    else if (state.verticalAlign !== "mixed" && state.verticalAlign !== value) state.verticalAlign = "mixed";
  };

  const mergeNumberFormat = (raw: unknown) => {
    const value = typeof raw === "string" ? raw : null;
    if (state.numberFormat === undefined) state.numberFormat = value;
    else if (state.numberFormat !== "mixed" && state.numberFormat !== value) state.numberFormat = "mixed";
  };

  const mergeFontName = (raw: unknown) => {
    const value = typeof raw === "string" ? raw : null;
    if (state.fontName === undefined) state.fontName = value;
    else if (state.fontName !== "mixed" && state.fontName !== value) state.fontName = "mixed";
  };

  const mergeFontSize = (raw: unknown) => {
    const value = typeof raw === "number" && Number.isFinite(raw) ? raw : null;
    if (state.fontSize === undefined) state.fontSize = value;
    else if (state.fontSize !== "mixed" && state.fontSize !== value) state.fontSize = "mixed";
  };
  const mergeFontVariantPosition = (raw: unknown) => {
    const normalized = typeof raw === "string" ? raw.toLowerCase() : null;
    const value = normalized === "subscript" || normalized === "superscript" ? (normalized as "subscript" | "superscript") : null;
    if (state.fontVariantPosition === undefined) state.fontVariantPosition = value;
    else if (state.fontVariantPosition !== "mixed" && state.fontVariantPosition !== value) state.fontVariantPosition = "mixed";
  };
  const hasGetCellFormat = typeof anyDoc.getCellFormat === "function";
  const hasGetCellFormatStyleIds = typeof anyDoc.getCellFormatStyleIds === "function";
  const effectiveStyleCache = new Map<string, any>();

  const inspectCell = (row: number, col: number) => {
    const key = `${row},${col}`;
    if (visited.has(key)) return;
    visited.add(key);
    state.inspected += 1;

    // Prefer effective formatting (layered sheet/row/col/cell styles) when available.
    // Fall back to legacy cell-level styleId for older controller implementations.
    const style = (() => {
      if (hasGetCellFormat) {
        // `getCellFormat()` does a deep-merge across style layers. Cache by contributing
        // style ids so large selections with uniform formatting don't repeatedly merge.
        if (hasGetCellFormatStyleIds) {
          const styleIds = anyDoc.getCellFormatStyleIds(sheetId, { row, col }) as any;
          const cacheKey = Array.isArray(styleIds) ? styleIds.join("|") : String(styleIds);
          const cached = effectiveStyleCache.get(cacheKey);
          if (cached !== undefined) return cached;
          const computed = anyDoc.getCellFormat(sheetId, { row, col }) as any;
          effectiveStyleCache.set(cacheKey, computed);
          return computed;
        }
        return anyDoc.getCellFormat(sheetId, { row, col }) as any;
      }

      const cell =
        typeof anyDoc.peekCell === "function" ? (anyDoc.peekCell(sheetId, { row, col }) as any) : (doc.getCell(sheetId, { row, col }) as any);
      return doc.styleTable.get(cell?.styleId ?? 0) as any;
    })();

    const font = style?.font ?? null;
    mergeFontVariantPosition(font?.vertAlign);

    const bold =
      typeof font?.bold === "boolean"
        ? font.bold
        : typeof style?.bold === "boolean"
          ? style.bold
          : false;
    state.bold = state.bold && bold;

    const italic =
      typeof font?.italic === "boolean"
        ? font.italic
        : typeof style?.italic === "boolean"
          ? style.italic
          : false;
    state.italic = state.italic && italic;

    const underlineRaw =
      typeof font?.underline === "boolean" || typeof font?.underline === "string"
        ? font.underline
        : typeof style?.underline === "boolean" || typeof style?.underline === "string"
          ? style.underline
          : undefined;
    const underline = underlineRaw === true || (typeof underlineRaw === "string" && underlineRaw !== "none");
    state.underline = state.underline && underline;

    const strike =
      typeof font?.strike === "boolean"
        ? font.strike
        : typeof style?.strike === "boolean"
          ? style.strike
          : false;
    state.strikethrough = state.strikethrough && strike;

    const alignment = style?.alignment ?? null;
    const wrapText = getStyleWrapText(style);
    state.wrapText = state.wrapText && wrapText;

    const fontName =
      typeof font?.name === "string"
        ? font.name
        : typeof style?.fontFamily === "string"
          ? style.fontFamily
          : typeof style?.font_family === "string"
            ? style.font_family
            : typeof style?.fontName === "string"
              ? style.fontName
              : typeof style?.font_name === "string"
                ? style.font_name
                : undefined;
    mergeFontName(fontName);

    mergeFontSize(getStyleFontSizePt(style));

    // Alignment keys may exist in camelCase or snake_case forms depending on provenance.
    // Treat an explicit `alignment.horizontal` (including null) as authoritative so callers can
    // clear imported `horizontal_alignment` values.
    let horizontalRaw: unknown = undefined;
    if (hasOwn(alignment, "horizontal")) horizontalRaw = (alignment as any).horizontal;
    else if (hasOwn(alignment, "horizontal_align")) horizontalRaw = (alignment as any).horizontal_align;
    else if (hasOwn(alignment, "horizontal_alignment")) horizontalRaw = (alignment as any).horizontal_alignment;
    else if (hasOwn(alignment, "horizontalAlign")) horizontalRaw = (alignment as any).horizontalAlign;
    else if (hasOwn(alignment, "horizontalAlignment")) horizontalRaw = (alignment as any).horizontalAlignment;
    else if (hasOwn(style, "horizontalAlign")) horizontalRaw = (style as any).horizontalAlign;
    else if (hasOwn(style, "horizontal_align")) horizontalRaw = (style as any).horizontal_align;
    else if (hasOwn(style, "horizontalAlignment")) horizontalRaw = (style as any).horizontalAlignment;
    else if (hasOwn(style, "horizontal_alignment")) horizontalRaw = (style as any).horizontal_alignment;
    else horizontalRaw = (alignment as any)?.horizontal;
    mergeAlign(horizontalRaw);

    // Like horizontal alignment, vertical alignment may arrive in camelCase or snake_case forms.
    // Treat an explicit `alignment.vertical` (including null) as authoritative so callers can
    // clear imported `vertical_alignment` values.
    let verticalRaw: unknown = undefined;
    if (hasOwn(alignment, "vertical")) verticalRaw = (alignment as any).vertical;
    else if (hasOwn(alignment, "vertical_align")) verticalRaw = (alignment as any).vertical_align;
    else if (hasOwn(alignment, "vertical_alignment")) verticalRaw = (alignment as any).vertical_alignment;
    else if (hasOwn(alignment, "verticalAlign")) verticalRaw = (alignment as any).verticalAlign;
    else if (hasOwn(alignment, "verticalAlignment")) verticalRaw = (alignment as any).verticalAlignment;
    else if (hasOwn(style, "verticalAlign")) verticalRaw = (style as any).verticalAlign;
    else if (hasOwn(style, "vertical_align")) verticalRaw = (style as any).vertical_align;
    else if (hasOwn(style, "verticalAlignment")) verticalRaw = (style as any).verticalAlignment;
    else if (hasOwn(style, "vertical_alignment")) verticalRaw = (style as any).vertical_alignment;
    else verticalRaw = (alignment as any)?.vertical;
    mergeVerticalAlign(verticalRaw);
    mergeNumberFormat(getStyleNumberFormat(style));
  };

  outer: for (const range of ranges) {
    if (exhaustive) {
      for (let row = range.startRow; row <= range.endRow; row++) {
        for (let col = range.startCol; col <= range.endCol; col++) {
          inspectCell(row, col);
          if (state.inspected >= maxInspectCells) break outer;
        }
      }
    } else {
      const remaining = maxInspectCells - state.inspected;
      for (const cell of sampleRangeCells(range, remaining)) {
        inspectCell(cell.row, cell.col);
        if (state.inspected >= maxInspectCells) break outer;
      }
    }
  }

  return {
    // Note: When the selection is very large we sample rather than scan every cell.
    // These values should be treated as "best effort" and may miss rare outliers,
    // but should still be responsive for typical UI usage.
    bold: state.bold,
    italic: state.italic,
    underline: state.underline,
    strikethrough: state.strikethrough,
    fontName: state.fontName === undefined ? null : state.fontName,
    fontSize: state.fontSize === undefined ? null : state.fontSize,
    fontVariantPosition: state.fontVariantPosition === undefined ? null : state.fontVariantPosition,
    wrapText: state.wrapText,
    align: state.align ?? "left",
    verticalAlign: state.verticalAlign ?? "bottom",
    numberFormat: state.numberFormat === undefined ? null : state.numberFormat,
  };
}
