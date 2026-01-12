import type { CellData, CellProvider, CellProviderUpdate, CellRange, CellStyle } from "@formula/grid";
import { LruCache } from "@formula/grid";
import type { DocumentController } from "../../document/documentController.js";
import { resolveCssVar } from "../../theme/cssVars.js";
import { formatValueWithNumberFormat } from "../../formatting/numberFormat.js";

type RichTextValue = { text: string; runs?: Array<{ start: number; end: number; style?: Record<string, unknown> }> };

type DocStyle = Record<string, any>;

function isPlainObject(value: unknown): value is Record<string, any> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

/**
 * Convert `#AARRGGBB` strings into `rgba(r,g,b,a)` strings usable by canvas.
 * Returns `undefined` when the input is invalid.
 */
function argbToCanvasCss(argb: unknown): string | undefined {
  if (typeof argb !== "string") return undefined;
  const match = /^#([0-9a-fA-F]{8})$/.exec(argb.trim());
  if (!match) return undefined;

  const a = Number.parseInt(match[1].slice(0, 2), 16);
  const r = Number.parseInt(match[1].slice(2, 4), 16);
  const g = Number.parseInt(match[1].slice(4, 6), 16);
  const b = Number.parseInt(match[1].slice(6, 8), 16);
  if (![a, r, g, b].every((n) => Number.isFinite(n))) return undefined;

  const alpha = Math.min(1, Math.max(0, Math.round((a / 255) * 1000) / 1000));
  return `rgba(${r},${g},${b},${alpha})`;
}

function isRichTextValue(value: unknown): value is RichTextValue {
  if (typeof value !== "object" || value == null) return false;
  const v = value as { text?: unknown; runs?: unknown };
  if (typeof v.text !== "string") return false;
  if (v.runs == null) return true;
  return Array.isArray(v.runs);
}

function toColumnName(col0: number): string {
  let value = col0 + 1;
  let name = "";
  while (value > 0) {
    const rem = (value - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

export class DocumentCellProvider implements CellProvider {
  private readonly headerStyle: CellStyle = { fontWeight: "600", textAlign: "center" };
  private readonly rowHeaderStyle: CellStyle = { fontWeight: "600", textAlign: "end" };

  private readonly cache: LruCache<string, CellData | null>;
  private readonly styleCache = new Map<number, CellStyle | undefined>();
  private readonly listeners = new Set<(update: CellProviderUpdate) => void>();
  private unsubscribeDoc: (() => void) | null = null;

  constructor(
    private readonly options: {
      document: DocumentController;
      /**
       * Active sheet id for the grid view.
       *
       * The provider is only ever asked for the currently-rendered sheet; callers
       * should update this when switching sheets.
       */
      getSheetId: () => string;
      headerRows: number;
      headerCols: number;
      rowCount: number;
      colCount: number;
      showFormulas: () => boolean;
      getComputedValue: (cell: { row: number; col: number }) => string | number | boolean | null;
      getCommentMeta?: (cellRef: string) => { resolved: boolean } | null;
    }
  ) {
    // Cache covers cell metadata + value formatting work. Keep it bounded to avoid
    // memory blow-ups on huge scrolls.
    this.cache = new LruCache<string, CellData | null>(50_000);
  }

  private resolveStyle(styleId: unknown): CellStyle | undefined {
    const id = typeof styleId === "number" && Number.isInteger(styleId) && styleId >= 0 ? styleId : 0;
    if (this.styleCache.has(id)) return this.styleCache.get(id);

    const docStyle: DocStyle = this.options.document?.styleTable?.get?.(id) ?? {};
    const style = this.convertDocStyleToGridStyle(docStyle);
    this.styleCache.set(id, style);
    return style;
  }

  private convertDocStyleToGridStyle(docStyle: unknown): CellStyle | undefined {
    if (!isPlainObject(docStyle)) return undefined;

    // Note: `@formula/grid` CellStyle is evolving; the shared-grid rendering pipeline reads
    // additional formatting primitives (borders, underline, etc.) off this object at runtime.
    // We intentionally build this as a plain object and cast at the end to avoid tight coupling
    // to the exact type shape.
    const out: any = {};

    const fill = isPlainObject(docStyle.fill) ? docStyle.fill : null;
    const fillColor = argbToCanvasCss(fill?.fgColor ?? fill?.background);
    if (fillColor) out.fill = fillColor;

    const font = isPlainObject(docStyle.font) ? docStyle.font : null;
    if (font?.bold === true) out.fontWeight = "700";
    if (font?.italic === true) out.fontStyle = "italic";
    if (font?.underline === true) out.underline = true;
    if (font?.strike === true) out.strike = true;
    if (typeof font?.name === "string" && font.name.trim() !== "") out.fontFamily = font.name;
    if (typeof font?.size === "number" && Number.isFinite(font.size)) {
      out.fontSize = (font.size * 96) / 72;
    }
    const fontColor = argbToCanvasCss(font?.color);
    if (fontColor) out.color = fontColor;

    const alignment = isPlainObject(docStyle.alignment) ? docStyle.alignment : null;
    const horizontal = alignment?.horizontal;
    if (horizontal === "center") out.textAlign = "center";
    else if (horizontal === "left") out.textAlign = "left";
    else if (horizontal === "right") out.textAlign = "right";
    // "general"/undefined: leave undefined so renderer can pick based on value type.
    if (alignment?.wrapText === true) out.wrapMode = "word";

    const border = isPlainObject(docStyle.border) ? docStyle.border : null;
    if (border) {
      // Use a theme token for default border colors so dark mode remains legible.
      // `resolveCssVar()` returns computed values at runtime, but falls back
      // gracefully in unit tests / non-DOM environments.
      const defaultBorderColor = resolveCssVar("--text-primary", { fallback: "CanvasText" });
      const mapExcelBorderStyle = (style: unknown): { width: number; style: string } | null => {
        if (typeof style !== "string") return null;
        switch (style) {
          case "thin":
            return { width: 1, style: "solid" };
          case "medium":
            return { width: 2, style: "solid" };
          case "thick":
            return { width: 3, style: "solid" };
          case "dashed":
            return { width: 1, style: "dashed" };
          case "dotted":
            return { width: 1, style: "dotted" };
          case "double":
            return { width: 3, style: "double" };
          default:
            return null;
        }
      };

      const mapEdge = (edge: any) => {
        if (!isPlainObject(edge)) return undefined;
        const mapped = mapExcelBorderStyle(edge.style);
        if (!mapped) return undefined;
        const color = argbToCanvasCss(edge.color) ?? defaultBorderColor;
        return { width: mapped.width, style: mapped.style, color };
      };

      const borders: any = {};
      const left = mapEdge(border.left);
      const right = mapEdge(border.right);
      const top = mapEdge(border.top);
      const bottom = mapEdge(border.bottom);
      if (left) borders.left = left;
      if (right) borders.right = right;
      if (top) borders.top = top;
      if (bottom) borders.bottom = bottom;
      if (Object.keys(borders).length > 0) out.borders = borders;
    }

    return Object.keys(out).length > 0 ? (out as CellStyle) : undefined;
  }

  invalidateAll(): void {
    this.cache.clear();
    for (const listener of this.listeners) listener({ type: "invalidateAll" });
  }

  invalidateDocCells(range: { startRow: number; endRow: number; startCol: number; endCol: number }): void {
    const { headerRows, headerCols } = this.options;
    const gridRange: CellRange = {
      startRow: range.startRow + headerRows,
      endRow: range.endRow + headerRows,
      startCol: range.startCol + headerCols,
      endCol: range.endCol + headerCols
    };

    // Best-effort cache eviction for the affected region.
    const cellCount = Math.max(0, gridRange.endRow - gridRange.startRow) * Math.max(0, gridRange.endCol - gridRange.startCol);
    if (cellCount <= 1000) {
      const sheetId = this.options.getSheetId();
      for (let r = gridRange.startRow; r < gridRange.endRow; r++) {
        for (let c = gridRange.startCol; c < gridRange.endCol; c++) {
          this.cache.delete(`${sheetId}:${r},${c}`);
        }
      }
    } else {
      // If the range is large, clear the whole cache to avoid spending too much time
      // iterating keys.
      this.cache.clear();
    }

    for (const listener of this.listeners) listener({ type: "cells", range: gridRange });
  }

  prefetch(range: CellRange): void {
    // Prefetch is a hint used by async providers. We use it to warm our in-memory cache
    // but cap work so fast scrolls don't block the UI thread.
    const maxCells = 2_000;
    const rows = Math.max(0, range.endRow - range.startRow);
    const cols = Math.max(0, range.endCol - range.startCol);
    const total = rows * cols;
    if (total === 0) return;

    const budget = Math.max(0, Math.min(maxCells, total));
    const step = total / budget;

    let idx = 0;
    while (idx < total) {
      const i = Math.floor(idx);
      const r = range.startRow + Math.floor(i / cols);
      const c = range.startCol + (i % cols);
      this.getCell(r, c);
      idx += step;
    }
  }

  getCell(row: number, col: number): CellData | null {
    const { rowCount, colCount, headerRows, headerCols } = this.options;
    if (row < 0 || col < 0 || row >= rowCount || col >= colCount) return null;

    const sheetId = this.options.getSheetId();
    const key = `${sheetId}:${row},${col}`;
    const cached = this.cache.get(key);
    if (cached !== undefined) return cached;

    const headerRow = row < headerRows;
    const headerCol = col < headerCols;

    if (headerRow || headerCol) {
      let value: string | number | null = null;
      let style: CellStyle | undefined;
      if (headerRow && headerCol) {
        value = "";
        style = this.headerStyle;
      } else if (headerRow) {
        const docCol = col - headerCols;
        value = docCol >= 0 ? toColumnName(docCol) : "";
        style = this.headerStyle;
      } else {
        const docRow = row - headerRows;
        value = docRow >= 0 ? docRow + 1 : 0;
        style = this.rowHeaderStyle;
      }

      const cell: CellData = { row, col, value, style };
      this.cache.set(key, cell);
      return cell;
    }

    const docRow = row - headerRows;
    const docCol = col - headerCols;

    const state = this.options.document.getCell(sheetId, { row: docRow, col: docCol }) as {
      value: unknown;
      formula: string | null;
      styleId?: number;
    };
    if (!state) {
      this.cache.set(key, null);
      return null;
    }

    let value: string | number | boolean | null = null;
    let richText: RichTextValue | undefined;
    if (state.formula != null) {
      if (this.options.showFormulas()) {
        value = state.formula;
      } else {
        value = this.options.getComputedValue({ row: docRow, col: docCol });
      }
    } else if (state.value != null) {
      if (isRichTextValue(state.value)) {
        richText = state.value;
        value = richText.text;
      } else {
        value = state.value as any;
      }
      if (value !== null && typeof value !== "string" && typeof value !== "number" && typeof value !== "boolean") {
        value = String(state.value);
      }
    }

    const styleId = typeof state.styleId === "number" ? state.styleId : 0;
    const docStyle: DocStyle = this.options.document?.styleTable?.get?.(styleId) ?? {};
    const styleAny = docStyle as any;
    const numberFormat =
      typeof styleAny?.numberFormat === "string"
        ? (styleAny.numberFormat as string)
        : typeof styleAny?.number_format === "string"
          ? (styleAny.number_format as string)
          : null;

    let style = this.resolveStyle(state.styleId);

    if (typeof value === "number" && numberFormat !== null) {
      value = formatValueWithNumberFormat(value, numberFormat);
      // Preserve spreadsheet-like default alignment for numeric values even though we
      // render them as formatted strings (CanvasGridRenderer defaults to left-aligning strings).
      if (style?.textAlign === undefined) {
        style = { ...(style ?? {}), textAlign: "end" };
      }
    }

    const comment = (() => {
      const metaProvider = this.options.getCommentMeta;
      if (!metaProvider) return null;
      const cellRef = `${toColumnName(docCol)}${docRow + 1}`;
      const meta = metaProvider(cellRef);
      if (!meta) return null;
      return { resolved: meta.resolved };
    })();

    const cell: CellData = richText ? { row, col, value, richText, style, comment } : { row, col, value, style, comment };
    this.cache.set(key, cell);
    return cell;
  }

  subscribe(listener: (update: CellProviderUpdate) => void): () => void {
    this.listeners.add(listener);

    if (!this.unsubscribeDoc) {
      // Coalesce document mutations into provider updates so the renderer can redraw
      // minimal dirty regions.
      this.unsubscribeDoc = this.options.document.on("change", (payload: any) => {
        const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
        if (deltas.length === 0) {
          // Sheet view deltas (frozen panes, row/col sizes, etc.) do not affect cell contents.
          // Avoid evicting the provider cache in those cases; the renderer will be updated by
          // the view sync code (e.g. `syncFrozenPanes` / shared grid axis sync).
          const sheetViewDeltas = Array.isArray(payload?.sheetViewDeltas) ? payload.sheetViewDeltas : [];
          if (sheetViewDeltas.length > 0 && payload?.recalc !== true) {
            return;
          }
          this.invalidateAll();
          return;
        }

        const sheetId = this.options.getSheetId();
        let minRow = Infinity;
        let maxRow = -Infinity;
        let minCol = Infinity;
        let maxCol = -Infinity;
        let saw = false;

        for (const delta of deltas) {
          if (!delta) continue;
          if (String(delta.sheetId ?? "") !== sheetId) continue;
          const row = Number(delta.row);
          const col = Number(delta.col);
          if (!Number.isInteger(row) || row < 0) continue;
          if (!Number.isInteger(col) || col < 0) continue;
          saw = true;
          minRow = Math.min(minRow, row);
          maxRow = Math.max(maxRow, row);
          minCol = Math.min(minCol, col);
          maxCol = Math.max(maxCol, col);
        }

        if (!saw) {
          this.invalidateAll();
          return;
        }

        this.invalidateDocCells({
          startRow: minRow,
          endRow: maxRow + 1,
          startCol: minCol,
          endCol: maxCol + 1
        });
      });
    }

    return () => {
      this.listeners.delete(listener);
      if (this.listeners.size === 0 && this.unsubscribeDoc) {
        this.unsubscribeDoc();
        this.unsubscribeDoc = null;
      }
    };
  }
}
