import type { CellData, CellProvider, CellProviderUpdate, CellRange, CellStyle } from "@formula/grid";
import { LruCache } from "@formula/grid";
import type { DocumentController } from "../../document/documentController.js";
import { resolveCssVar } from "../../theme/cssVars.js";
import { formatValueWithNumberFormat } from "../../formatting/numberFormat.js";

type RichTextValue = { text: string; runs?: Array<{ start: number; end: number; style?: Record<string, unknown> }> };

type DocStyle = Record<string, any>;

const CACHE_KEY_COL_STRIDE = 65_536;
const SHEET_CACHE_MAX_SIZE = 50_000;

function isPlainObject(value: unknown): value is Record<string, any> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

/**
 * Normalize a hex color to a CSS color string.
 *
 * This intentionally matches the clipboard’s permissive behavior:
 * - accepts `#RRGGBB` / `RRGGBB`
 * - accepts Excel/OOXML ARGB `#AARRGGBB` / `AARRGGBB` (converts to `#RRGGBB` or `rgba(...)`)
 */
function normalizeHexToCssColor(hex: string): string | null {
  const normalized = hex.trim().replace(/^#/, "");
  if (!/^[0-9a-fA-F]+$/.test(normalized)) return null;

  // #RRGGBB
  if (normalized.length === 6) return `#${normalized}`;

  // Excel/OOXML commonly stores colors as ARGB (AARRGGBB).
  if (normalized.length === 8) {
    const a = Number.parseInt(normalized.slice(0, 2), 16);
    const r = Number.parseInt(normalized.slice(2, 4), 16);
    const g = Number.parseInt(normalized.slice(4, 6), 16);
    const b = Number.parseInt(normalized.slice(6, 8), 16);
    if (![a, r, g, b].every((n) => Number.isFinite(n))) return null;

    if (a >= 255) {
      return `#${normalized.slice(2)}`;
    }

    const alpha = Math.max(0, Math.min(1, a / 255));
    const rounded = Math.round(alpha * 1000) / 1000;
    return `rgba(${r},${g},${b},${rounded})`;
  }

  return null;
}

function normalizeCssColor(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;

  // Accept plain CSS colors as-is.
  if (!/^[0-9a-fA-F#]+$/.test(trimmed)) return trimmed;

  return normalizeHexToCssColor(trimmed) ?? trimmed;
}

function isRichTextValue(value: unknown): value is RichTextValue {
  if (typeof value !== "object" || value == null) return false;
  const v = value as { text?: unknown; runs?: unknown };
  if (typeof v.text !== "string") return false;
  if (v.runs == null) return true;
  return Array.isArray(v.runs);
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
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

  /**
   * Per-sheet caches avoid `${sheetId}:${row},${col}` string allocations in the hot
   * `getCell` path. Keys are encoded as `row * 65536 + col` which is safe for Excel's
   * maxes (col <= 16_384; rows ~1M) and leaves ample headroom below MAX_SAFE_INTEGER.
   */
  private readonly sheetCaches = new Map<string, LruCache<number, CellData | null>>();
  private lastSheetId: string | null = null;
  private lastSheetCache: LruCache<number, CellData | null> | null = null;
  private readonly styleCache = new Map<number, CellStyle | undefined>();
  private readonly resolvedStyleCache = new WeakMap<object, CellStyle | undefined>();
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
    // Caches cover cell metadata + value formatting work. Keep each sheet bounded to
    // avoid memory blow-ups on huge scrolls.
  }

  private resolveStyle(styleId: unknown): CellStyle | undefined {
    const id = typeof styleId === "number" && Number.isInteger(styleId) && styleId >= 0 ? styleId : 0;
    if (this.styleCache.has(id)) return this.styleCache.get(id);

    const docStyle: DocStyle = this.options.document?.styleTable?.get?.(id) ?? {};
    const style = this.convertDocStyleToGridStyle(docStyle);
    this.styleCache.set(id, style);
    return style;
  }

  private resolveResolvedStyle(docStyle: unknown): CellStyle | undefined {
    if (!isPlainObject(docStyle)) return undefined;
    if (this.resolvedStyleCache.has(docStyle)) return this.resolvedStyleCache.get(docStyle);
    const style = this.convertDocStyleToGridStyle(docStyle);
    this.resolvedStyleCache.set(docStyle, style);
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
    const fillColor = normalizeCssColor(
      fill?.fgColor ??
        fill?.background ??
        fill?.bgColor ??
        (docStyle as any).backgroundColor ??
        (docStyle as any).background_color ??
        (docStyle as any).fillColor ??
        (docStyle as any).fill_color
    );
    if (fillColor) out.fill = fillColor;

    const font = isPlainObject(docStyle.font) ? docStyle.font : null;
    const bold =
      typeof font?.bold === "boolean"
        ? font.bold
        : typeof (docStyle as any).bold === "boolean"
          ? (docStyle as any).bold
          : undefined;
    if (bold === true) out.fontWeight = "700";

    const italic =
      typeof font?.italic === "boolean"
        ? font.italic
        : typeof (docStyle as any).italic === "boolean"
          ? (docStyle as any).italic
          : undefined;
    if (italic === true) out.fontStyle = "italic";

    const underlineRaw =
      typeof font?.underline === "boolean" || typeof font?.underline === "string"
        ? font.underline
        : typeof (docStyle as any).underline === "boolean" || typeof (docStyle as any).underline === "string"
          ? (docStyle as any).underline
          : undefined;
    if (underlineRaw === true) out.underline = true;
    if (typeof underlineRaw === "string" && underlineRaw !== "none") out.underline = true;

    const strike =
      typeof font?.strike === "boolean"
        ? font.strike
        : typeof (docStyle as any).strike === "boolean"
          ? (docStyle as any).strike
          : undefined;
    if (strike === true) out.strike = true;

    if (typeof font?.name === "string" && font.name.trim() !== "") out.fontFamily = font.name;
    if (typeof font?.size === "number" && Number.isFinite(font.size)) {
      out.fontSize = (font.size * 96) / 72;
    }
    const fontColor = normalizeCssColor(
      font?.color ??
        (docStyle as any).textColor ??
        (docStyle as any).text_color ??
        (docStyle as any).fontColor ??
        (docStyle as any).font_color
    );
    if (fontColor) out.color = fontColor;

    const rawNumberFormat = (docStyle as any).numberFormat ?? (docStyle as any).number_format;
    if (typeof rawNumberFormat === "string" && rawNumberFormat.trim() !== "") out.numberFormat = rawNumberFormat;

    const alignment = isPlainObject(docStyle.alignment) ? docStyle.alignment : null;
    const horizontal = alignment?.horizontal;
    if (horizontal === "center") out.textAlign = "center";
    else if (horizontal === "left") out.textAlign = "start";
    else if (horizontal === "right") out.textAlign = "end";
    // "general"/undefined: leave undefined so renderer can pick based on value type.
    if (alignment?.wrapText === true) out.wrapMode = "word";

    const vertical = alignment?.vertical;
    if (vertical === "top") out.verticalAlign = "top";
    else if (vertical === "center") out.verticalAlign = "middle";
    else if (vertical === "bottom") out.verticalAlign = "bottom";

    const rotationRaw =
      typeof (alignment as any)?.textRotation === "number" && Number.isFinite((alignment as any).textRotation)
        ? (alignment as any).textRotation
        : typeof (alignment as any)?.rotation === "number" && Number.isFinite((alignment as any).rotation)
          ? (alignment as any).rotation
          : undefined;

    if (rotationRaw != null) {
      // Excel's OOXML `textRotation` uses `255` as a sentinel for "vertical text"
      // (stacked letters). We approximate this in the grid by rotating 90 degrees.
      const normalized = rotationRaw === 255 ? 90 : rotationRaw;
      out.rotationDeg = clamp(normalized, -180, 180);
    }

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
        const color = normalizeCssColor(edge.color) ?? defaultBorderColor;
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
    this.sheetCaches.clear();
    this.lastSheetId = null;
    this.lastSheetCache = null;
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
      const cache = this.sheetCaches.get(sheetId);
      if (cache) {
        for (let r = gridRange.startRow; r < gridRange.endRow; r++) {
          const base = r * CACHE_KEY_COL_STRIDE;
          for (let c = gridRange.startCol; c < gridRange.endCol; c++) {
            cache.delete(base + c);
          }
        }
      }
    } else {
      // If the range is large, clear the caches to avoid spending too much time
      // iterating keys. This mirrors the legacy behavior where a large invalidation
      // cleared the provider-wide cache.
      this.sheetCaches.clear();
      this.lastSheetId = null;
      this.lastSheetCache = null;
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
    const cache = this.getSheetCache(sheetId);
    const key = row * CACHE_KEY_COL_STRIDE + col;
    const cached = cache.get(key);
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
      cache.set(key, cell);
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
      cache.set(key, null);
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
    const controller: any = this.options.document as any;
    const resolvedDocStyle: unknown =
      typeof controller.getCellFormat === "function"
        ? controller.getCellFormat(sheetId, { row: docRow, col: docCol })
        : this.options.document?.styleTable?.get?.(styleId) ?? {};
    const docStyle: DocStyle = isPlainObject(resolvedDocStyle) ? (resolvedDocStyle as DocStyle) : {};
    const styleAny = docStyle as any;
    const numberFormat =
      typeof styleAny?.numberFormat === "string"
        ? (styleAny.numberFormat as string)
        : typeof styleAny?.number_format === "string"
          ? (styleAny.number_format as string)
          : null;

    let style = typeof controller.getCellFormat === "function" ? this.resolveResolvedStyle(docStyle) : this.resolveStyle(styleId);

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
    cache.set(key, cell);
    return cell;
  }

  private getSheetCache(sheetId: string): LruCache<number, CellData | null> {
    if (this.lastSheetId === sheetId && this.lastSheetCache) return this.lastSheetCache;

    let cache = this.sheetCaches.get(sheetId);
    if (!cache) {
      cache = new LruCache<number, CellData | null>(SHEET_CACHE_MAX_SIZE);
      this.sheetCaches.set(sheetId, cache);
    }

    this.lastSheetId = sheetId;
    this.lastSheetCache = cache;
    return cache;
  }

  subscribe(listener: (update: CellProviderUpdate) => void): () => void {
    this.listeners.add(listener);

    if (!this.unsubscribeDoc) {
      // Coalesce document mutations into provider updates so the renderer can redraw
      // minimal dirty regions.
      this.unsubscribeDoc = this.options.document.on("change", (payload: any) => {
        const sheetId = this.options.getSheetId();
        const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
        const sheetViewDeltas = Array.isArray(payload?.sheetViewDeltas) ? payload.sheetViewDeltas : [];

        // New layered formatting deltas (row/col/sheet style maps) may arrive without per-cell deltas.
        const rowStyleDeltas = Array.isArray(payload?.rowStyleDeltas) ? payload.rowStyleDeltas : [];
        const colStyleDeltas = Array.isArray(payload?.colStyleDeltas) ? payload.colStyleDeltas : [];
        const sheetStyleDeltas = Array.isArray(payload?.sheetStyleDeltas) ? payload.sheetStyleDeltas : [];
        // DocumentController emits layered format updates as `formatDeltas`.
        const formatDeltas = Array.isArray(payload?.formatDeltas) ? payload.formatDeltas : [];

        const recalc = payload?.recalc === true;
        const hasFormatLayerDeltas =
          rowStyleDeltas.length > 0 || colStyleDeltas.length > 0 || sheetStyleDeltas.length > 0 || formatDeltas.length > 0;

        // No cell deltas + no formatting deltas: preserve the sheet-view optimization.
        if (deltas.length === 0 && !hasFormatLayerDeltas) {
          // Sheet view deltas (frozen panes, row/col sizes, etc.) do not affect cell contents.
          // Avoid evicting the provider cache in those cases; the renderer will be updated by
          // the view sync code (e.g. `syncFrozenPanes` / shared grid axis sync).
          if (sheetViewDeltas.length > 0 && !recalc) {
            return;
          }
          this.invalidateAll();
          return;
        }

        // Sheet-level style changes affect all cells.
        if (
          sheetStyleDeltas.some((delta: any) => {
            if (!delta) return false;
            const id = delta.sheetId;
            // If the delta doesn't specify a sheet, conservatively assume it impacts the visible sheet.
            if (id == null) return true;
            return String(id) === sheetId;
          }) ||
          formatDeltas.some((delta: any) => delta && String(delta.sheetId ?? "") === sheetId && delta.layer === "sheet")
        ) {
          this.invalidateAll();
          return;
        }

        const docRowCount = Math.max(0, this.options.rowCount - this.options.headerRows);
        const docColCount = Math.max(0, this.options.colCount - this.options.headerCols);
        const clampInt = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));

        const invalidateRanges: Array<{ startRow: number; endRow: number; startCol: number; endCol: number }> = [];

        const collectAxisSpan = (axisDeltas: any[], axis: "row" | "col"): { min: number; max: number } | null | "unknown" => {
          if (!axisDeltas || axisDeltas.length === 0) return null;
          let min = Infinity;
          let max = -Infinity;
          let sawAnyForSheet = false;
          for (const delta of axisDeltas) {
            if (!delta) continue;
            if (String(delta.sheetId ?? "") !== sheetId) continue;
            sawAnyForSheet = true;

            const indices: number[] = [];

            if (Array.isArray(delta.indices)) {
              for (const v of delta.indices) indices.push(Number(v));
            }
            if (Array.isArray(delta.rows) && axis === "row") {
              for (const v of delta.rows) indices.push(Number(v));
            }
            if (Array.isArray(delta.cols) && axis === "col") {
              for (const v of delta.cols) indices.push(Number(v));
            }

            const direct = delta[axis];
            if (direct != null) indices.push(Number(direct));

            const index = delta.index;
            if (index != null) indices.push(Number(index));

            const startKey = axis === "row" ? "startRow" : "startCol";
            const endKey = axis === "row" ? "endRow" : "endCol";
            const endExclusiveKey = axis === "row" ? "endRowExclusive" : "endColExclusive";

            const start = delta[startKey];
            const end = delta[endKey];
            const endExclusive = delta[endExclusiveKey];

            const startNum = Number(start);
            const endNum = Number(end);
            const endExclusiveNum = Number(endExclusive);
            if (Number.isInteger(startNum) && startNum >= 0) {
              if (Number.isInteger(endExclusiveNum) && endExclusiveNum > startNum) {
                // Exclusive end.
                min = Math.min(min, startNum);
                max = Math.max(max, endExclusiveNum - 1);
                continue;
              }
              if (Number.isInteger(endNum) && endNum >= startNum) {
                // Assume inclusive end.
                min = Math.min(min, startNum);
                max = Math.max(max, endNum);
                continue;
              }
            }

            for (const idx of indices) {
              if (!Number.isInteger(idx) || idx < 0) continue;
              min = Math.min(min, idx);
              max = Math.max(max, idx);
            }
          }

          if (!sawAnyForSheet) return null;
          if (min === Infinity || max === -Infinity) return "unknown";
          return { min, max };
        };

        // Cell-level deltas: keep existing minimal invalidation behavior.
        if (deltas.length > 0) {
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
            // If we can't determine the region (e.g. sheet mismatch), fall back.
            // This mirrors the prior behavior and ensures we don't miss cross-sheet formula dependencies.
            this.invalidateAll();
            return;
          }

          invalidateRanges.push({
            startRow: clampInt(minRow, 0, docRowCount),
            endRow: clampInt(maxRow + 1, 0, docRowCount),
            startCol: clampInt(minCol, 0, docColCount),
            endCol: clampInt(maxCol + 1, 0, docColCount)
          });
        }

        const rowSpan = collectAxisSpan(
          [...rowStyleDeltas, ...formatDeltas.filter((d: any) => d && d.layer === "row")],
          "row"
        );
        if (rowSpan === "unknown") {
          this.invalidateAll();
          return;
        }
        if (rowSpan) {
          invalidateRanges.push({
            startRow: clampInt(rowSpan.min, 0, docRowCount),
            endRow: clampInt(rowSpan.max + 1, 0, docRowCount),
            startCol: 0,
            endCol: docColCount
          });
        }

        const colSpan = collectAxisSpan(
          [...colStyleDeltas, ...formatDeltas.filter((d: any) => d && d.layer === "col")],
          "col"
        );
        if (colSpan === "unknown") {
          this.invalidateAll();
          return;
        }
        if (colSpan) {
          invalidateRanges.push({
            startRow: 0,
            endRow: docRowCount,
            startCol: clampInt(colSpan.min, 0, docColCount),
            endCol: clampInt(colSpan.max + 1, 0, docColCount)
          });
        }

        if (invalidateRanges.length === 0) {
          // Formatting/view deltas did not apply to this sheet. Only invalidate on recalc.
          if (recalc) {
            this.invalidateAll();
          }
          return;
        }

        // Emit provider updates for each invalidation span. This is conservative but lets the renderer
        // redraw without forcing a full-sheet invalidation for common row/col formatting operations.
        for (const range of invalidateRanges) {
          if (range.endRow <= range.startRow || range.endCol <= range.startCol) continue;
          this.invalidateDocCells(range);
        }
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
