import { DocumentController } from "../../document/documentController.js";
import { applyStylePatch } from "../../formatting/styleTable.js";

import { normalizeFormulaTextOpt } from "@formula/engine/backend/formula";

import { formatA1Range, type CellAddress, type RangeAddress } from "../../../../../packages/ai-tools/src/spreadsheet/a1.ts";
import type { CellEntry, SpreadsheetApi } from "../../../../../packages/ai-tools/src/spreadsheet/api.ts";
import { isCellEmpty, type CellData, type CellFormat } from "../../../../../packages/ai-tools/src/spreadsheet/types.ts";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver.js";

type DocumentControllerStyle = Record<string, any>;

function cloneCellValue(value: any): any {
  if (value == null || typeof value !== "object") return value;
  // `structuredClone` is available in modern browsers + Node, but TypeScript's DOM libs
  // don't always include it on `globalThis` depending on configuration.
  const structuredCloneFn =
    typeof (globalThis as any).structuredClone === "function" ? ((globalThis as any).structuredClone as any) : null;
  return structuredCloneFn ? structuredCloneFn(value) : JSON.parse(JSON.stringify(value));
}

function isPlainObject(value: unknown): value is Record<string, any> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function cellValuesEqual(left: any, right: any): boolean {
  if (left === right) return true;
  if (left == null || right == null) return left === right;
  if (typeof left !== typeof right) return false;
  if (typeof left === "object") {
    try {
      return JSON.stringify(left) === JSON.stringify(right);
    } catch {
      return false;
    }
  }
  return false;
}

function normalizeFormula(raw: unknown): string | undefined {
  if (raw == null) return undefined;
  const normalized = normalizeFormulaTextOpt(String(raw));
  return normalized === null ? undefined : normalized;
}

function styleToCellFormat(style: DocumentControllerStyle | null | undefined): CellFormat | undefined {
  if (!style || typeof style !== "object") return undefined;

  const out: CellFormat = {};

  const font = style.font;
  if (font && typeof font === "object") {
    if (typeof font.bold === "boolean") out.bold = font.bold;
    if (typeof font.italic === "boolean") out.italic = font.italic;
    if (typeof font.size === "number") out.font_size = font.size;
    if (typeof font.color === "string") out.font_color = font.color;
  }

  const fill = style.fill;
  if (fill && typeof fill === "object") {
    const color = typeof fill.fgColor === "string" ? fill.fgColor : typeof fill.background === "string" ? fill.background : null;
    if (color != null) out.background_color = color;
  }

  if (typeof style.numberFormat === "string") out.number_format = style.numberFormat;

  const alignment = style.alignment;
  if (alignment && typeof alignment === "object") {
    const horizontal = alignment.horizontal;
    if (horizontal === "left" || horizontal === "center" || horizontal === "right") {
      out.horizontal_align = horizontal;
    }
  }

  // Back-compat: older adapters/snapshots may store ai-tools `CellFormat` fields
  // directly on the style object (flat snake_case or camelCase keys).
  const styleAny = style as any;

  if (out.bold === undefined && typeof styleAny.bold === "boolean") out.bold = styleAny.bold;
  if (out.italic === undefined && typeof styleAny.italic === "boolean") out.italic = styleAny.italic;

  if (out.font_size === undefined) {
    if (typeof styleAny.font_size === "number") out.font_size = styleAny.font_size;
    else if (typeof styleAny.fontSize === "number") out.font_size = styleAny.fontSize;
  }

  if (out.font_color === undefined) {
    if (typeof styleAny.font_color === "string") out.font_color = styleAny.font_color;
    else if (typeof styleAny.fontColor === "string") out.font_color = styleAny.fontColor;
  }

  if (out.background_color === undefined) {
    if (typeof styleAny.background_color === "string") out.background_color = styleAny.background_color;
    else if (typeof styleAny.backgroundColor === "string") out.background_color = styleAny.backgroundColor;
  }

  if (out.number_format === undefined) {
    if (typeof styleAny.number_format === "string") out.number_format = styleAny.number_format;
    else if (typeof styleAny.numberFormat === "string") out.number_format = styleAny.numberFormat;
  }

  if (out.horizontal_align === undefined) {
    const align = styleAny.horizontal_align ?? styleAny.horizontalAlign;
    if (align === "left" || align === "center" || align === "right") {
      out.horizontal_align = align;
    }
  }

  return Object.keys(out).length > 0 ? out : undefined;
}

function cellFormatToStylePatch(format: Partial<CellFormat> | null | undefined): DocumentControllerStyle | null {
  if (!format) return null;

  /** @type {DocumentControllerStyle} */
  const patch: DocumentControllerStyle = {};

  const setFont = (key: string, value: unknown) => {
    patch.font ??= {};
    patch.font[key] = value;
  };

  if (typeof format.bold === "boolean") setFont("bold", format.bold);
  if (typeof format.italic === "boolean") setFont("italic", format.italic);
  if (typeof format.font_size === "number") setFont("size", format.font_size);
  if (typeof format.font_color === "string") setFont("color", format.font_color);

  if (typeof format.background_color === "string") {
    patch.fill = { pattern: "solid", fgColor: format.background_color };
  }

  if (typeof format.number_format === "string") {
    patch.numberFormat = format.number_format;
  }

  if (format.horizontal_align === "left" || format.horizontal_align === "center" || format.horizontal_align === "right") {
    patch.alignment ??= {};
    patch.alignment.horizontal = format.horizontal_align;
  }

  return Object.keys(patch).length > 0 ? patch : null;
}

function styleForWrite(baseStyle: DocumentControllerStyle, format: CellFormat | null | undefined): DocumentControllerStyle {
  const style = (isPlainObject(baseStyle) ? cloneCellValue(baseStyle) : {}) as DocumentControllerStyle;

  // Remove supported keys before applying the next format, so writeRange can move
  // formatting without "contaminating" target cells with stale attributes.
  if (isPlainObject(style.font)) {
    delete style.font.bold;
    delete style.font.italic;
    delete style.font.size;
    delete style.font.color;
    if (Object.keys(style.font).length === 0) delete style.font;
  } else {
    delete style.font;
  }

  if (isPlainObject(style.fill)) {
    delete style.fill.fgColor;
    delete style.fill.background;
    delete style.fill.pattern;
    if (Object.keys(style.fill).length === 0) delete style.fill;
  } else {
    delete style.fill;
  }

  delete style.numberFormat;

  if (isPlainObject(style.alignment)) {
    delete style.alignment.horizontal;
    if (Object.keys(style.alignment).length === 0) delete style.alignment;
  } else {
    delete style.alignment;
  }

  // Legacy flat keys from older adapters/snapshots.
  delete (style as any).bold;
  delete (style as any).italic;
  delete (style as any).font_size;
  delete (style as any).fontSize;
  delete (style as any).font_color;
  delete (style as any).fontColor;
  delete (style as any).background_color;
  delete (style as any).backgroundColor;
  delete (style as any).number_format;
  delete (style as any).numberFormat;
  delete (style as any).horizontal_align;
  delete (style as any).horizontalAlign;

  if (!format || Object.keys(format).length === 0) return style;

  if (typeof format.bold === "boolean") {
    style.font ??= {};
    style.font.bold = format.bold;
  }
  if (typeof format.italic === "boolean") {
    style.font ??= {};
    style.font.italic = format.italic;
  }
  if (typeof format.font_size === "number") {
    style.font ??= {};
    style.font.size = format.font_size;
  }
  if (typeof format.font_color === "string") {
    style.font ??= {};
    style.font.color = format.font_color;
  }

  if (typeof format.background_color === "string") {
    style.fill ??= {};
    style.fill.pattern = "solid";
    style.fill.fgColor = format.background_color;
  }

  if (typeof format.number_format === "string") {
    style.numberFormat = format.number_format;
  }

  if (format.horizontal_align === "left" || format.horizontal_align === "center" || format.horizontal_align === "right") {
    style.alignment ??= {};
    style.alignment.horizontal = format.horizontal_align;
  }

  return style;
}

function toCellData(
  controller: DocumentController,
  sheetId: string,
  coord: { row: number; col: number },
  cellState: any
): CellData {
  const style =
    typeof (controller as any).getCellFormat === "function"
      ? (controller as any).getCellFormat(sheetId, coord)
      : (() => {
          const styleId = typeof cellState?.styleId === "number" ? cellState.styleId : 0;
          return styleId === 0 ? null : controller.styleTable.get(styleId);
        })();
  const format = styleToCellFormat(style);

  const rawFormula = cellState?.formula;
  const normalizedFormula = normalizeFormula(rawFormula);

  // Real spreadsheet models may store both a formula string and a cached/computed value.
  // Preserve `cellState.value` even when `formula` is present so `@formula/ai-tools`
  // ToolExecutor can optionally surface/use computed formula values (opt-in via
  // `include_formula_values`).
  const value = cloneCellValue(cellState?.value ?? null);

  return {
    value,
    ...(normalizedFormula ? { formula: normalizedFormula } : {}),
    ...(format ? { format } : {})
  };
}

function toControllerCoord(address: { row: number; col: number }): { row: number; col: number } {
  return { row: address.row - 1, col: address.col - 1 };
}

function toControllerRange(range: RangeAddress): { start: { row: number; col: number }; end: { row: number; col: number } } {
  return {
    start: { row: range.startRow - 1, col: range.startCol - 1 },
    end: { row: range.endRow - 1, col: range.endCol - 1 }
  };
}

function parseSemanticDiffCellKey(key: string): { row: number; col: number } {
  const match = /^r(\d+)c(\d+)$/.exec(key);
  if (!match) throw new Error(`Invalid semantic diff cell key: ${key}`);
  const row = Number(match[1]);
  const col = Number(match[2]);
  if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) {
    throw new Error(`Invalid semantic diff cell key: ${key}`);
  }
  return { row, col };
}

function toCellDataFromExportedCell(exportedCell: any): CellData {
  const rawFormula = exportedCell?.formula;
  const normalizedFormula = normalizeFormula(rawFormula);

  const format = styleToCellFormat(exportedCell?.format ?? null);
  const value = normalizedFormula ? null : cloneCellValue(exportedCell?.value ?? null);

  return {
    value,
    ...(normalizedFormula ? { formula: normalizedFormula } : {}),
    ...(format ? { format } : {})
  };
}

function parseControllerCellKey(key: string): { row: number; col: number } {
  const commaIdx = key.indexOf(",");
  if (commaIdx === -1) throw new Error(`Invalid cell key: ${key}`);
  const row = Number(key.slice(0, commaIdx));
  const col = Number(key.slice(commaIdx + 1));
  if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) {
    throw new Error(`Invalid cell key: ${key}`);
  }
  return { row, col };
}

function toCellDataFromCellState(controller: DocumentController, cellState: any): CellData {
  const rawFormula = cellState?.formula;
  const normalizedFormula = normalizeFormula(rawFormula);

  const styleId = typeof cellState?.styleId === "number" ? cellState.styleId : 0;
  const format = styleId === 0 ? undefined : styleToCellFormat(controller.styleTable.get(styleId));

  const value = cloneCellValue(cellState?.value ?? null);

  return {
    value,
    ...(normalizedFormula ? { formula: normalizedFormula } : {}),
    ...(format ? { format } : {})
  };
}

/**
 * Adapter that lets `packages/ai-tools` execute tool calls against the real
 * `DocumentController` workbook model (used by the desktop app).
 *
 * This makes it possible to reuse the tool executor + preview engine on top of
 * the UI controller before the Rust calc engine integration exists.
 */
export class DocumentControllerSpreadsheetApi implements SpreadsheetApi {
  readonly controller: DocumentController;
  readonly createChart?: SpreadsheetApi["createChart"];
  readonly sheetNameResolver: SheetNameResolver | null;

  constructor(
    controller: DocumentController,
    options: { createChart?: SpreadsheetApi["createChart"]; sheetNameResolver?: SheetNameResolver | null } = {}
  ) {
    this.controller = controller;
    this.createChart = options.createChart;
    this.sheetNameResolver = options.sheetNameResolver ?? null;
  }

  listSheets(): string[] {
    const ids = this.controller.getSheetIds();
    // DocumentController creates sheets lazily; expose the default sheet name so
    // downstream consumers (RAG context, etc) can still reason about "Sheet1"
    // even before any edits.
    const sheetIds = ids.length > 0 ? ids : ["Sheet1"];
    return sheetIds.map((id) => this.getSheetNameById(id));
  }

  listNonEmptyCells(sheet?: string): CellEntry[] {
    const sheetIds = sheet ? [this.resolveSheetIdOrThrow(sheet)] : this.controller.getSheetIds();
    const entries: CellEntry[] = [];
    for (const sheetId of sheetIds) {
      const displayName = this.getSheetNameById(sheetId);
      const sheetModel = this.controller.model.sheets.get(sheetId);
      if (!sheetModel?.cells) continue;
      for (const [key, cellState] of sheetModel.cells.entries()) {
        if (!cellState) continue;
        const value = cellState.value ?? null;
        const formula = cellState.formula ?? null;
        if (value == null && formula == null) continue;
        const { row, col } = parseControllerCellKey(key);
        entries.push({
          address: { sheet: displayName, row: row + 1, col: col + 1 },
          cell: {
            value: value != null && typeof value === "object" ? cloneCellValue(value) : value,
            ...(formula != null ? { formula } : {})
          }
        });
      }
    }
    return entries;
  }

  getCell(address: CellAddress): CellData {
    const sheetId = this.resolveSheetIdOrThrow(address.sheet);
    const coord = toControllerCoord(address);
    const state = this.controller.getCell(sheetId, coord);
    return toCellData(this.controller, sheetId, coord, state);
  }

  setCell(address: CellAddress, cell: CellData): void {
    const sheetId = this.resolveSheetIdOrThrow(address.sheet);
    const coord = toControllerCoord(address);
    const beforeCell = this.getCell(address);
    const desiredFormula = normalizeFormula(cell.formula);
    const desiredValue = desiredFormula ? null : (cell.value ?? null);
    const shouldUpdateContent =
      !cellValuesEqual(beforeCell.value, desiredValue) || (beforeCell.formula ?? undefined) !== desiredFormula;

    if (!shouldUpdateContent && (!cell.format || Object.keys(cell.format).length === 0)) {
      return;
    }

    this.controller.beginBatch({ label: "AI write_cell" });
    try {
      if (shouldUpdateContent && desiredFormula) {
        this.controller.setCellFormula(sheetId, coord, desiredFormula, { label: "AI write_cell" });
      } else if (shouldUpdateContent) {
        this.controller.setCellValue(sheetId, coord, cell.value ?? null, { label: "AI write_cell" });
      }

      if (cell.format && Object.keys(cell.format).length > 0) {
        const patch = cellFormatToStylePatch(cell.format);
        if (patch) {
          this.controller.setRangeFormat(sheetId, { start: coord, end: coord }, patch, {
            label: "AI apply_formatting"
          });
        }
      }
    } finally {
      this.controller.endBatch();
    }
  }

  readRange(range: RangeAddress): CellData[][] {
    const sheetId = this.resolveSheetIdOrThrow(range.sheet);
    const rowCount = Math.max(0, range.endRow - range.startRow + 1);
    const colCount = Math.max(0, range.endCol - range.startCol + 1);

    const rows: CellData[][] = new Array(rowCount);

    // Hot path: many callers (WorkbookContextBuilder) read 5k-10k cells at once.
    // Avoid per-cell `DocumentController.getCell()` calls which clone cell state,
    // normalize formulas, and perform style lookups.
    const sheetModel = this.controller.model.sheets.get(sheetId);
    const cellMap: Map<string, any> | undefined = sheetModel?.cells;

    const hasLayeredFormattingReadPath = typeof (this.controller as any).getCellFormat === "function";

    const formatCache = new Map<string | number, CellFormat | undefined>();

    const getFormatForStyleIds = (styleIds: [number, number, number, number, number]): CellFormat | undefined => {
      const [sheetStyleId, rowStyleId, colStyleId, runStyleId, cellStyleId] = styleIds;
      if (sheetStyleId === 0 && rowStyleId === 0 && colStyleId === 0 && runStyleId === 0 && cellStyleId === 0) return undefined;
      const cacheKey = `${sheetStyleId},${rowStyleId},${colStyleId},${runStyleId},${cellStyleId}`;
      if (formatCache.has(cacheKey)) return formatCache.get(cacheKey);

      const sheetStyle = this.controller.styleTable.get(sheetStyleId);
      const colStyle = this.controller.styleTable.get(colStyleId);
      const rowStyle = this.controller.styleTable.get(rowStyleId);
      const runStyle = this.controller.styleTable.get(runStyleId);
      const cellStyle = this.controller.styleTable.get(cellStyleId);

      // Precedence: sheet < col < row < range-run < cell.
      const sheetCol = applyStylePatch(sheetStyle, colStyle);
      const sheetColRow = applyStylePatch(sheetCol, rowStyle);
      const sheetColRowRun = applyStylePatch(sheetColRow, runStyle);
      const effectiveStyle = applyStylePatch(sheetColRowRun, cellStyle);

      const format = styleToCellFormat(effectiveStyle);
      formatCache.set(cacheKey, format);
      return format;
    };

    const getFormatForLegacyStyleId = (styleId: number): CellFormat | undefined => {
      if (styleId === 0) return undefined;
      if (formatCache.has(styleId)) return formatCache.get(styleId);
      const style = this.controller.styleTable.get(styleId);
      const format = styleToCellFormat(style);
      formatCache.set(styleId, format);
      return format;
    };

    const startRow0 = range.startRow - 1;
    const startCol0 = range.startCol - 1;

    type FormatRun = { startRow: number; endRowExclusive: number; styleId: number };
    const runListsByCol: Array<FormatRun[] | null> = new Array(colCount).fill(null);
    const runIndexByCol = new Array<number>(colCount).fill(0);

    if (hasLayeredFormattingReadPath && sheetModel?.formatRunsByCol?.get) {
      for (let c = 0; c < colCount; c++) {
        const col0 = startCol0 + c;
        const runs = sheetModel.formatRunsByCol.get(col0) as FormatRun[] | undefined;
        if (!Array.isArray(runs) || runs.length === 0) continue;
        runListsByCol[c] = runs;
        if (startRow0 <= 0) continue;
        // Initialize to the first run whose endRowExclusive is > startRow0.
        let lo = 0;
        let hi = runs.length - 1;
        let idx = runs.length;
        while (lo <= hi) {
          const mid = (lo + hi) >> 1;
          const run = runs[mid]!;
          if (startRow0 < run.endRowExclusive) {
            idx = mid;
            hi = mid - 1;
          } else {
            lo = mid + 1;
          }
        }
        runIndexByCol[c] = idx;
      }
    }

    for (let r = 0; r < rowCount; r++) {
      const row = new Array<CellData>(colCount);
      const row0 = startRow0 + r;
      for (let c = 0; c < colCount; c++) {
        const col0 = startCol0 + c;
        const cellState = cellMap?.get(`${row0},${col0}`);
        let runStyleId = 0;
        const runs = runListsByCol[c];
        if (runs) {
          let idx = runIndexByCol[c]!;
          while (idx < runs.length && row0 >= runs[idx]!.endRowExclusive) idx++;
          runIndexByCol[c] = idx;
          const run = idx < runs.length ? runs[idx] : null;
          if (run && row0 >= run.startRow && row0 < run.endRowExclusive) {
            runStyleId = typeof run.styleId === "number" ? run.styleId : 0;
          }
        }

        const styleIds: [number, number, number, number, number] = hasLayeredFormattingReadPath
          ? [
              typeof sheetModel?.defaultStyleId === "number" ? sheetModel.defaultStyleId : 0,
              typeof sheetModel?.rowStyleIds?.get === "function" ? (sheetModel.rowStyleIds.get(row0) ?? 0) : 0,
              typeof sheetModel?.colStyleIds?.get === "function" ? (sheetModel.colStyleIds.get(col0) ?? 0) : 0,
              runStyleId,
              typeof cellState?.styleId === "number" ? cellState.styleId : 0
            ]
          : [0, 0, 0, 0, typeof cellState?.styleId === "number" ? cellState.styleId : 0];

        const format = hasLayeredFormattingReadPath
          ? getFormatForStyleIds(styleIds)
          : getFormatForLegacyStyleId(styleIds[4]);

        if (!cellState) {
          row[c] = { value: null, ...(format ? { format } : {}) };
          continue;
        }

        const rawFormula = cellState.formula;
        if (rawFormula != null) {
          const normalizedFormula = normalizeFormula(rawFormula);
          if (normalizedFormula) {
            const rawValue = cellState.value ?? null;
            row[c] = {
              value: rawValue != null && typeof rawValue === "object" ? cloneCellValue(rawValue) : rawValue,
              formula: normalizedFormula,
              ...(format ? { format } : {})
            };
            continue;
          }
        }

        const rawValue = cellState.value ?? null;
        row[c] = {
          value: rawValue != null && typeof rawValue === "object" ? cloneCellValue(rawValue) : rawValue,
          ...(format ? { format } : {})
        };
      }
      rows[r] = row;
    }

    return rows;
  }

  writeRange(range: RangeAddress, cells: CellData[][]): void {
    const sheetId = this.resolveSheetIdOrThrow(range.sheet);
    const rowCount = range.endRow - range.startRow + 1;
    const colCount = range.endCol - range.startCol + 1;

    if (cells.length !== rowCount) {
      throw new Error(
        `writeRange expected ${rowCount} rows but got ${cells.length} rows for ${range.sheet}!R${range.startRow}C${range.startCol}:R${range.endRow}C${range.endCol}`
      );
    }

    for (const row of cells) {
      if (row.length !== colCount) {
        throw new Error(
          `writeRange expected ${colCount} columns but got ${row.length} columns for ${range.sheet}!R${range.startRow}C${range.startCol}:R${range.endRow}C${range.endCol}`
        );
      }
    }

    const hasAnyFormat = cells.some((row) => row.some((cell) => Boolean(cell?.format && Object.keys(cell.format).length > 0)));

    this.controller.beginBatch({ label: "AI set_range" });
    try {
      if (!hasAnyFormat) {
        const values = cells.map((row) =>
          row.map((cell) => (cell?.formula ? { formula: cell.formula } : cell?.value ?? null))
        );
        this.controller.setRangeValues(sheetId, toControllerRange(range), values, { label: "AI set_range" });
        return;
      }

      const sheetModel = this.controller.model.sheets.get(sheetId);
      const cellMap: Map<string, any> | undefined = sheetModel?.cells;
      const hasLayeredFormattingWritePath = typeof (this.controller as any).getCellFormat === "function";

      // Cache supported-format projections for styleIds since writeRange might touch thousands of cells (sort_range).
      const styleIdFormatCache = new Map<number, CellFormat | undefined>();
      const getFormatForStyleId = (styleId: number): CellFormat | undefined => {
        if (styleId === 0) return undefined;
        if (styleIdFormatCache.has(styleId)) return styleIdFormatCache.get(styleId);
        const style = this.controller.styleTable.get(styleId);
        const format = styleToCellFormat(style);
        styleIdFormatCache.set(styleId, format);
        return format;
      };

      const inheritedFormatCache = new Map<string, CellFormat | undefined>();
      const getInheritedFormat = (styleIds: [number, number, number, number]): CellFormat | undefined => {
        const [sheetStyleId, rowStyleId, colStyleId, runStyleId] = styleIds;
        if (sheetStyleId === 0 && rowStyleId === 0 && colStyleId === 0 && runStyleId === 0) return undefined;
        const key = `${sheetStyleId},${rowStyleId},${colStyleId},${runStyleId}`;
        if (inheritedFormatCache.has(key)) return inheritedFormatCache.get(key);

        const sheetFormat = getFormatForStyleId(sheetStyleId);
        const colFormat = getFormatForStyleId(colStyleId);
        const rowFormat = getFormatForStyleId(rowStyleId);
        const runFormat = getFormatForStyleId(runStyleId);

        // Precedence: sheet < col < row < range-run.
        const merged: CellFormat = { ...(sheetFormat ?? {}), ...(colFormat ?? {}), ...(rowFormat ?? {}), ...(runFormat ?? {}) };
        const out = Object.keys(merged).length > 0 ? merged : undefined;
        inheritedFormatCache.set(key, out);
        return out;
      };

      const startRow0 = range.startRow - 1;
      const startCol0 = range.startCol - 1;

      type FormatRun = { startRow: number; endRowExclusive: number; styleId: number };
      const runListsByCol: Array<FormatRun[] | null> = new Array(colCount).fill(null);
      const runIndexByCol = new Array<number>(colCount).fill(0);

      if (hasLayeredFormattingWritePath && sheetModel?.formatRunsByCol?.get) {
        for (let c = 0; c < colCount; c++) {
          const col0 = startCol0 + c;
          const runs = sheetModel.formatRunsByCol.get(col0) as FormatRun[] | undefined;
          if (!Array.isArray(runs) || runs.length === 0) continue;
          runListsByCol[c] = runs;
          if (startRow0 <= 0) continue;
          // Initialize to the first run whose endRowExclusive is > startRow0.
          let lo = 0;
          let hi = runs.length - 1;
          let idx = runs.length;
          while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            const run = runs[mid]!;
            if (startRow0 < run.endRowExclusive) {
              idx = mid;
              hi = mid - 1;
            } else {
              lo = mid + 1;
            }
          }
          runIndexByCol[c] = idx;
        }
      }

      const inputs: any[][] = [];
      for (let r = 0; r < rowCount; r++) {
        const srcRow = cells[r] ?? [];
        const outRow: any[] = [];
        for (let c = 0; c < colCount; c++) {
          const cell = srcRow[c] ?? { value: null };
          const row0 = startRow0 + r;
          const col0 = startCol0 + c;

          let runStyleId = 0;
          const runs = runListsByCol[c];
          if (runs) {
            let idx = runIndexByCol[c]!;
            while (idx < runs.length && row0 >= runs[idx]!.endRowExclusive) idx++;
            runIndexByCol[c] = idx;
            const run = idx < runs.length ? runs[idx] : null;
            if (run && row0 >= run.startRow && row0 < run.endRowExclusive) {
              runStyleId = typeof run.styleId === "number" ? run.styleId : 0;
            }
          }

          const cellState = cellMap?.get(`${row0},${col0}`);
          const cellStyleId = typeof cellState?.styleId === "number" ? cellState.styleId : 0;
          const baseCellStyle = cellStyleId === 0 ? {} : this.controller.styleTable.get(cellStyleId);

          const styleIds: [number, number, number, number, number] = hasLayeredFormattingWritePath
            ? [
                typeof sheetModel?.defaultStyleId === "number" ? sheetModel.defaultStyleId : 0,
                typeof sheetModel?.rowStyleIds?.get === "function" ? (sheetModel.rowStyleIds.get(row0) ?? 0) : 0,
                typeof sheetModel?.colStyleIds?.get === "function" ? (sheetModel.colStyleIds.get(col0) ?? 0) : 0,
                runStyleId,
                cellStyleId
              ]
            : [0, 0, 0, 0, cellStyleId];

          const inheritedFormat = hasLayeredFormattingWritePath
            ? getInheritedFormat([styleIds[0], styleIds[1], styleIds[2], styleIds[3]])
            : undefined;

          const requestedFormat = cell.format && Object.keys(cell.format).length > 0 ? cell.format : null;

          // CellData.format in writeRange is treated like the ai-tools InMemoryWorkbook: it is a
          // per-cell override for supported keys. We still want to preserve layered defaults
          // (sheet/row/col) without materializing them into per-cell styles. To do that, drop any
          // requested keys that are already satisfied by inherited formatting *when the cell does
          // not already have an explicit per-cell override for that key*.
          //
          // Note: we must preserve existing per-cell overrides even if they match inherited
          // formatting, otherwise writeRange round-trips (e.g. sort_range) could accidentally
          // "promote" direct cell formatting into column/row defaults, changing how formatting
          // moves if the inherited layer is later cleared.
          let formatToWrite: CellFormat | null | undefined = null;
          if (!requestedFormat) {
            // No format specified for this cell => clear ai-tools supported keys (preserve other style keys).
            formatToWrite = null;
          } else {
            const override: CellFormat = {};
            const inherited = inheritedFormat ?? {};
            const explicit = getFormatForStyleId(cellStyleId) ?? {};
            for (const key of Object.keys(requestedFormat) as Array<keyof CellFormat>) {
              const value = requestedFormat[key];
              if (value === undefined) continue;
              // Avoid materializing inherited formatting into per-cell styles for cells that don't
              // already have an explicit override, but keep any existing explicit overrides.
              if ((inherited as any)[key] === value && (explicit as any)[key] === undefined) continue;
              (override as any)[key] = value;
            }
            formatToWrite = Object.keys(override).length > 0 ? override : {};
          }

          const nextStyle = styleForWrite(baseCellStyle, formatToWrite);
          const input: any = { format: nextStyle };
          if (cell.formula) input.formula = cell.formula;
          else input.value = cell.value ?? null;
          outRow.push(input);
        }
        inputs.push(outRow);
      }

      this.controller.setRangeValues(sheetId, toControllerRange(range), inputs, { label: "AI set_range" });
    } finally {
      this.controller.endBatch();
    }
  }

  applyFormatting(range: RangeAddress, format: Partial<CellFormat>): number {
    const sheetId = this.resolveSheetIdOrThrow(range.sheet);
    const patch = cellFormatToStylePatch(format);
    const rows = range.endRow - range.startRow + 1;
    const cols = range.endCol - range.startCol + 1;
    if (!patch) return 0;

    const applied = this.controller.setRangeFormat(sheetId, toControllerRange(range), patch, { label: "AI apply_formatting" });
    if (applied === false) {
      const rangeForUser = { ...range, sheet: this.getSheetNameById(sheetId) };
      throw new Error(
        `Formatting could not be applied to ${formatA1Range(rangeForUser)}. Try selecting fewer cells/rows.`,
      );
    }
    return rows * cols;
  }

  getLastUsedRow(sheet: string): number {
    let max = 0;
    const sheetId = this.resolveSheetIdOrThrow(sheet);
    const sheetModel = this.controller.model.sheets.get(sheetId);
    if (!sheetModel?.cells) return 0;
    for (const [key, cellState] of sheetModel.cells.entries()) {
      const { row, col } = parseControllerCellKey(key);
      const cell = toCellData(this.controller, sheetId, { row, col }, cellState);
      if (isCellEmpty(cell)) continue;
      max = Math.max(max, row + 1);
    }
    return max;
  }

  clone(): SpreadsheetApi {
    const snapshot = this.controller.encodeState();
    const cloned = new DocumentController();
    cloned.applyState(snapshot);
    // `PreviewEngine` executes tool plans against clones. When charts are enabled
    // we still want create_chart previews to succeed, but we must avoid mutating
    // the live chart layer. Provide a throwaway chart implementation that only
    // returns ids.
    const createChart = this.createChart
      ? (() => {
          let counter = 0;
          return () => ({ chart_id: `preview_chart_${++counter}` });
        })()
      : undefined;
    return new DocumentControllerSpreadsheetApi(cloned, { createChart, sheetNameResolver: this.sheetNameResolver });
  }

  private getSheetNameById(sheetId: string): string {
    if (!this.sheetNameResolver) return sheetId;
    return this.sheetNameResolver.getSheetNameById(sheetId) ?? sheetId;
  }

  private resolveSheetIdOrThrow(sheet: string): string {
    const name = String(sheet ?? "").trim();
    if (!name) {
      throw new Error("Sheet name is required.");
    }

    // Prefer the shared resolver when available (handles renamed sheet display names).
    if (this.sheetNameResolver) {
      const resolved = this.sheetNameResolver.getSheetIdByName(name);
      if (resolved) return resolved;
    }

    // Fallback: avoid creating phantom sheets by only allowing known sheet ids.
    const knownSheetIds = this.controller.getSheetIds();
    const candidates = knownSheetIds.length > 0 ? knownSheetIds : ["Sheet1"];
    const match = candidates.find((id) => id.toLowerCase() === name.toLowerCase());
    if (match) return match;

    throw new Error(`Unknown sheet "${name}".`);
  }
}
