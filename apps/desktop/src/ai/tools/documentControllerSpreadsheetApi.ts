import { DocumentController } from "../../document/documentController.js";

import type { CellAddress, RangeAddress } from "../../../../../packages/ai-tools/src/spreadsheet/a1.js";
import type { CellEntry, SpreadsheetApi } from "../../../../../packages/ai-tools/src/spreadsheet/api.js";
import { isCellEmpty, type CellData, type CellFormat } from "../../../../../packages/ai-tools/src/spreadsheet/types.js";

type DocumentControllerStyle = Record<string, any>;

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

function toCellData(controller: DocumentController, cellState: any): CellData {
  const styleId = typeof cellState?.styleId === "number" ? cellState.styleId : 0;
  const style = styleId === 0 ? null : controller.styleTable.get(styleId);
  const format = styleToCellFormat(style);

  const rawFormula = cellState?.formula;
  const normalizedFormula =
    rawFormula == null || rawFormula === "" ? undefined : normalizeFormula(String(rawFormula));

  return {
    value: normalizedFormula ? null : (cellState?.value ?? null),
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

function parseRowColKey(key: string): { row: number; col: number } {
  const [rowStr, colStr] = key.split(",");
  const row = Number(rowStr);
  const col = Number(colStr);
  if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) {
    throw new Error(`Invalid sheet cell key: ${key}`);
  }
  return { row, col };
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
    const sheets = (this.controller as any).model?.sheets;
    if (!sheets || typeof sheets.get !== "function") return [];

    const sheetIds = sheet ? [sheet] : this.controller.getSheetIds();
    const entries: CellEntry[] = [];
    for (const sheetId of sheetIds) {
      const sheetModel = sheets.get(sheetId);
      if (!sheetModel) continue;
      for (const [key, state] of sheetModel.cells?.entries?.() ?? []) {
        const { row, col } = parseRowColKey(key);
        const cell = toCellData(this.controller, state);
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
    this.controller.beginBatch({ label: "AI write_cell" });
    try {
      if (cell.formula) {
        this.controller.setCellFormula(address.sheet, coord, cell.formula, { label: "AI write_cell" });
      } else {
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

    this.controller.beginBatch({ label: "AI set_range" });
    try {
      const values = cells.map((row) =>
        row.map((cell) => (cell?.formula ? { formula: cell.formula } : cell?.value ?? null))
      );
      this.controller.setRangeValues(range.sheet, toControllerRange(range), values, { label: "AI set_range" });

      for (let r = 0; r < rowCount; r++) {
        const row = cells[r] ?? [];
        for (let c = 0; c < colCount; c++) {
          const format = row[c]?.format;
          if (!format || Object.keys(format).length === 0) continue;
          const patch = cellFormatToStylePatch(format);
          if (!patch) continue;
          const coord = { row: range.startRow - 1 + r, col: range.startCol - 1 + c };
          this.controller.setRangeFormat(range.sheet, { start: coord, end: coord }, patch, { label: "AI apply_formatting" });
        }
      }
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
    const sheets = (this.controller as any).model?.sheets;
    const sheetModel = sheets?.get?.(sheet);
    if (!sheetModel) return 0;
    let max = 0;
    for (const [key, state] of sheetModel.cells?.entries?.() ?? []) {
      const { row } = parseRowColKey(key);
      const cell = toCellData(this.controller, state);
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
