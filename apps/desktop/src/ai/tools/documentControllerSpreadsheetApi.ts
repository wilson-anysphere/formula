import { DocumentController } from "../../document/documentController.js";

import type { CellAddress, RangeAddress } from "../../../../../packages/ai-tools/src/spreadsheet/a1.js";
import type { CellEntry, SpreadsheetApi } from "../../../../../packages/ai-tools/src/spreadsheet/api.js";
import { isCellEmpty, type CellData, type CellFormat } from "../../../../../packages/ai-tools/src/spreadsheet/types.js";

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

function normalizeFormula(raw: string): string {
  const trimmed = raw.trimStart();
  if (!trimmed) return "=";
  return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
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

function toCellData(controller: DocumentController, cellState: any): CellData {
  const styleId = typeof cellState?.styleId === "number" ? cellState.styleId : 0;
  const style = styleId === 0 ? null : controller.styleTable.get(styleId);
  const format = styleToCellFormat(style);

  const rawFormula = cellState?.formula;
  const normalizedFormula =
    rawFormula == null || rawFormula === "" ? undefined : normalizeFormula(String(rawFormula));

  const value = normalizedFormula ? null : cloneCellValue(cellState?.value ?? null);

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
  const normalizedFormula =
    rawFormula == null || rawFormula === "" ? undefined : normalizeFormula(String(rawFormula));

  const format = styleToCellFormat(exportedCell?.format ?? null);
  const value = normalizedFormula ? null : cloneCellValue(exportedCell?.value ?? null);

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

  constructor(controller: DocumentController, options: { createChart?: SpreadsheetApi["createChart"] } = {}) {
    this.controller = controller;
    this.createChart = options.createChart;
  }

  listSheets(): string[] {
    const ids = this.controller.getSheetIds();
    // DocumentController creates sheets lazily; expose the default sheet name so
    // downstream consumers (RAG context, etc) can still reason about "Sheet1"
    // even before any edits.
    return ids.length > 0 ? ids : ["Sheet1"];
  }

  listNonEmptyCells(sheet?: string): CellEntry[] {
    const sheetIds = sheet ? [sheet] : this.controller.getSheetIds();
    const entries: CellEntry[] = [];
    for (const sheetId of sheetIds) {
      const exported = this.controller.exportSheetForSemanticDiff(sheetId);
      for (const [key, exportedCell] of exported.cells?.entries?.() ?? []) {
        const { row, col } = parseSemanticDiffCellKey(key);
        const cell = toCellDataFromExportedCell(exportedCell);
        if (isCellEmpty(cell)) continue;
        entries.push({
          address: { sheet: sheetId, row: row + 1, col: col + 1 },
          cell
        });
      }
    }
    return entries;
  }

  getCell(address: CellAddress): CellData {
    const state = this.controller.getCell(address.sheet, toControllerCoord(address));
    return toCellData(this.controller, state);
  }

  setCell(address: CellAddress, cell: CellData): void {
    const coord = toControllerCoord(address);
    const beforeCell = this.getCell(address);
    const desiredFormula = cell.formula ? normalizeFormula(String(cell.formula)) : undefined;
    const desiredValue = desiredFormula ? null : (cell.value ?? null);
    const shouldUpdateContent =
      !cellValuesEqual(beforeCell.value, desiredValue) || (beforeCell.formula ?? undefined) !== desiredFormula;

    if (!shouldUpdateContent && (!cell.format || Object.keys(cell.format).length === 0)) {
      return;
    }

    this.controller.beginBatch({ label: "AI write_cell" });
    try {
      if (shouldUpdateContent && cell.formula) {
        this.controller.setCellFormula(address.sheet, coord, cell.formula, { label: "AI write_cell" });
      } else if (shouldUpdateContent) {
        this.controller.setCellValue(address.sheet, coord, cell.value ?? null, { label: "AI write_cell" });
      }

      if (cell.format && Object.keys(cell.format).length > 0) {
        const patch = cellFormatToStylePatch(cell.format);
        if (patch) {
          this.controller.setRangeFormat(address.sheet, { start: coord, end: coord }, patch, {
            label: "AI apply_formatting"
          });
        }
      }
    } finally {
      this.controller.endBatch();
    }
  }

  readRange(range: RangeAddress): CellData[][] {
    const rows: CellData[][] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const row: CellData[] = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        row.push(this.getCell({ sheet: range.sheet, row: r, col: c }));
      }
      rows.push(row);
    }
    return rows;
  }

  writeRange(range: RangeAddress, cells: CellData[][]): void {
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
        this.controller.setRangeValues(range.sheet, toControllerRange(range), values, { label: "AI set_range" });
        return;
      }

      const inputs: any[][] = [];
      for (let r = 0; r < rowCount; r++) {
        const srcRow = cells[r] ?? [];
        const outRow: any[] = [];
        for (let c = 0; c < colCount; c++) {
          const cell = srcRow[c] ?? { value: null };
          const coord = { row: range.startRow - 1 + r, col: range.startCol - 1 + c };
          const before = this.controller.getCell(range.sheet, coord);
          const baseStyle = before.styleId === 0 ? {} : this.controller.styleTable.get(before.styleId);
          const nextStyle = styleForWrite(baseStyle, cell.format);
          const input: any = { format: nextStyle };
          if (cell.formula) input.formula = cell.formula;
          else input.value = cell.value ?? null;
          outRow.push(input);
        }
        inputs.push(outRow);
      }

      this.controller.setRangeValues(range.sheet, toControllerRange(range), inputs, { label: "AI set_range" });
    } finally {
      this.controller.endBatch();
    }
  }

  applyFormatting(range: RangeAddress, format: Partial<CellFormat>): number {
    const patch = cellFormatToStylePatch(format);
    if (patch) {
      this.controller.setRangeFormat(range.sheet, toControllerRange(range), patch, { label: "AI apply_formatting" });
    }
    const rows = range.endRow - range.startRow + 1;
    const cols = range.endCol - range.startCol + 1;
    return patch ? rows * cols : 0;
  }

  getLastUsedRow(sheet: string): number {
    let max = 0;
    const exported = this.controller.exportSheetForSemanticDiff(sheet);
    for (const [key, exportedCell] of exported.cells?.entries?.() ?? []) {
      const { row } = parseSemanticDiffCellKey(key);
      const cell = toCellDataFromExportedCell(exportedCell);
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
    return new DocumentControllerSpreadsheetApi(cloned, { createChart });
  }
}
