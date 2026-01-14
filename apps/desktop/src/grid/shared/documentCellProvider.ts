import type { CellData, CellProvider, CellProviderUpdate, CellRange, CellRichText, CellStyle } from "@formula/grid/node";
import { DEFAULT_GRID_FONT_FAMILY, LruCache, toColumnName } from "@formula/grid/node";
import type { DocumentController } from "../../document/documentController.js";
import { applyStylePatch } from "../../formatting/styleTable.js";
import { resolveCssVar } from "../../theme/cssVars.js";
import { formatValueWithNumberFormat } from "../../formatting/numberFormat.ts";
import { getStyleNumberFormat } from "../../formatting/styleFieldAccess.js";
import { normalizeExcelColorToCss } from "../../shared/colors.js";
import { looksLikeExternalHyperlink } from "./looksLikeExternalHyperlink.js";

type RichTextValue = CellRichText;

type DocStyle = Record<string, any>;

const CACHE_KEY_COL_STRIDE = 65_536;
const DEFAULT_SHEET_CACHE_MAX_SIZE = 50_000;
// Excel stores alignment indent as an integer "level" (Increase Indent).
// We approximate each indent level as an 8px text indent at zoom=1.
const INDENT_STEP_PX = 8;
type ResolvedFormat = { style: CellStyle | undefined; numberFormat: string | null };
const DEFAULT_RESOLVED_FORMAT: ResolvedFormat = {
  style: undefined,
  numberFormat: null,
};
const DEFAULT_NUMERIC_ALIGNMENT_STYLE: CellStyle = { textAlign: "end" };
// Style ids come from DocumentController's StyleTable; in practice these stay well under 1M.
// Using a large stride allows encoding `(sheetDefaultStyleId, layerStyleId)` pairs into a single
// collision-free number for cache keys (stays below MAX_SAFE_INTEGER for ids < 1_048_576).
const STYLE_ID_PAIR_STRIDE = 1_048_576;
const DEFAULT_SINGLE_LAYER_FORMAT_CACHE_MAX_SIZE = 10_000;
const DEFAULT_RESOLVED_FORMAT_CACHE_MAX_SIZE = 10_000;

function isPlainObject(value: unknown): value is Record<string, any> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function hasOwn(obj: unknown, key: string): boolean {
  return Boolean(obj) && Object.prototype.hasOwnProperty.call(obj as object, key);
}

function normalizeCssColor(value: unknown): string | null {
  return normalizeExcelColorToCss(value) ?? null;
}

function isRichTextValue(value: unknown): value is RichTextValue {
  if (typeof value !== "object" || value == null) return false;
  const v = value as { text?: unknown; runs?: unknown };
  if (typeof v.text !== "string") return false;
  if (v.runs == null) return true;
  return Array.isArray(v.runs);
}

function normalizePositiveNumber(value: unknown): number | undefined {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) return undefined;
  return value;
}

function parseImageCellPayload(value: unknown): CellData["image"] | null {
  if (!isPlainObject(value)) return null;
  const obj: any = value;

  let payload: any = null;

  // formula-model `CellValue` envelope: `{ type: "image", value: {...} }`.
  if (typeof obj.type === "string" && obj.type.toLowerCase() === "image") {
    payload = isPlainObject(obj.value) ? obj.value : null;
  } else if (typeof obj.type === "string") {
    return null;
  } else {
    // Legacy / direct payload shapes.
    payload = obj;
  }

  if (!payload) return null;
  const imageIdRaw = payload.imageId ?? payload.image_id ?? payload.id;
  if (typeof imageIdRaw !== "string") return null;
  const imageId = imageIdRaw.trim();
  if (imageId === "") return null;

  const altTextRaw = payload.altText ?? payload.alt_text ?? payload.alt;
  let altText: string | undefined;
  if (typeof altTextRaw === "string") {
    const trimmed = altTextRaw.trim();
    if (trimmed !== "") altText = trimmed;
  }

  const width = normalizePositiveNumber(payload.width ?? payload.w);
  const height = normalizePositiveNumber(payload.height ?? payload.h);

  return { imageId, altText, width, height };
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

export class DocumentCellProvider implements CellProvider {
  private readonly headerStyle: CellStyle;
  private readonly rowHeaderStyle: CellStyle;
  private resolvedDefaultHyperlinkStyle: CellStyle | null = null;
  private resolvedLinkColor: string | null = null;
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
    getComputedValue: (cell: { row: number; col: number }) => unknown;
    getCommentMeta?: (row: number, col: number) => { resolved: boolean } | null;
    /**
     * Optional root element used when resolving theme CSS variables.
     *
     * This is useful when `--formula-grid-*` vars are overridden on a specific
     * grid root (e.g. split panes) instead of being applied globally.
     */
    cssVarRoot?: HTMLElement | null;
    /**
     * Max entries per sheet in the per-cell cache. Lower values reduce memory usage
     * but may increase document reads and formatting work during fast scrolling.
     */
    sheetCacheMaxSize?: number;
    /**
     * Max entries in the resolved multi-layer format LRU cache.
     */
    resolvedFormatCacheMaxSize?: number;
    /**
     * Max entries per single-layer resolved format cache (row/col/range-run/cell).
     */
    singleLayerFormatCacheMaxSize?: number;
  };
  private readonly sheetCacheMaxSize: number;
  private readonly resolvedFormatCacheMaxSize: number;
  private readonly singleLayerFormatCacheMaxSize: number;

  /**
   * Per-sheet caches avoid `${sheetId}:${row},${col}` string allocations in the hot
   * `getCell` path. Keys are encoded as `row * 65536 + col` which is safe for Excel's
   * maxes (col <= 16_384; rows ~1M) and leaves ample headroom below MAX_SAFE_INTEGER.
   */
  private readonly sheetCaches = new Map<string, LruCache<number, CellData | null>>();
  private lastSheetId: string | null = null;
  private lastSheetCache: LruCache<number, CellData | null> | null = null;
  private readonly coordScratch = { row: 0, col: 0 };
  private readonly styleCache = new Map<number, CellStyle | undefined>();
  private readonly sheetDefaultResolvedFormatCache = new Map<number, ResolvedFormat>();
  private readonly sheetColResolvedFormatCache: LruCache<number, ResolvedFormat>;
  private readonly sheetRowResolvedFormatCache: LruCache<number, ResolvedFormat>;
  private readonly sheetRunResolvedFormatCache: LruCache<number, ResolvedFormat>;
  private readonly sheetCellResolvedFormatCache: LruCache<number, ResolvedFormat>;
  private readonly numericAlignmentStyleCache = new WeakMap<CellStyle, CellStyle>();
  // Cache resolved layered formats by contributing style ids (sheet/col/row/range-run/cell). This avoids
  // re-merging OOXML-ish style objects for every cell when large regions share the same
  // formatting layers (e.g. column formatting).
  private readonly resolvedFormatCache: LruCache<string, ResolvedFormat>;
  private readonly listeners = new Set<(update: CellProviderUpdate) => void>();
  private unsubscribeDoc: (() => void) | null = null;
  private readonly mergedEpochBySheet = new Map<string, number>();
  private readonly mergedRangesBySheet = new Map<
    string,
    { epoch: number; ranges: Array<{ startRow: number; endRow: number; startCol: number; endCol: number }> }
  >();
  private disposed = false;

  /**
   * DocumentController lazily materializes sheets (calling `getCell()` / `getSheetView()` will
   * create missing sheet ids). The shared grid provider is frequently asked for cell contents
   * while the UI is transitioning between sheets (e.g. sheet deletion/undo) and should never
   * recreate a deleted sheet as a side-effect of rendering.
   *
   * We treat a sheet as "known missing" when the workbook already has *some* sheets, but the
   * requested `sheetId` is present in neither the underlying `model.sheets` map nor the
   * `sheetMeta` map. When the workbook is still empty (no materialized sheets yet), we treat
   * the sheet id as "unknown" and allow the normal lazy materialization behavior.
   */
  private isSheetKnownMissing(sheetId: string): boolean {
    const id = String(sheetId ?? "").trim();
    if (!id) return true;

    const docAny: any = this.options.document as any;
    const sheets: any = docAny?.model?.sheets;
    const sheetMeta: any = docAny?.sheetMeta;

    if (
      sheets &&
      typeof sheets.has === "function" &&
      typeof sheets.size === "number" &&
      sheetMeta &&
      typeof sheetMeta.has === "function" &&
      typeof sheetMeta.size === "number"
    ) {
      const workbookHasAnySheets = sheets.size > 0 || sheetMeta.size > 0;
      if (!workbookHasAnySheets) return false;
      return !sheets.has(id) && !sheetMeta.has(id);
    }

    // If we can't introspect the document internals (e.g. unit tests with a fake document),
    // conservatively assume the sheet exists so we preserve the previous behavior.
    return false;
  }

  constructor(options: {
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
    getComputedValue: (cell: { row: number; col: number }) => unknown;
    getCommentMeta?: (row: number, col: number) => { resolved: boolean } | null;
    cssVarRoot?: HTMLElement | null;
    sheetCacheMaxSize?: number;
    resolvedFormatCacheMaxSize?: number;
    singleLayerFormatCacheMaxSize?: number;
  }) {
    this.options = options;
    this.sheetCacheMaxSize = options.sheetCacheMaxSize ?? DEFAULT_SHEET_CACHE_MAX_SIZE;
    this.resolvedFormatCacheMaxSize = options.resolvedFormatCacheMaxSize ?? DEFAULT_RESOLVED_FORMAT_CACHE_MAX_SIZE;
    this.singleLayerFormatCacheMaxSize = options.singleLayerFormatCacheMaxSize ?? DEFAULT_SINGLE_LAYER_FORMAT_CACHE_MAX_SIZE;

    this.sheetColResolvedFormatCache = new LruCache<number, ResolvedFormat>(this.singleLayerFormatCacheMaxSize);
    this.sheetRowResolvedFormatCache = new LruCache<number, ResolvedFormat>(this.singleLayerFormatCacheMaxSize);
    this.sheetRunResolvedFormatCache = new LruCache<number, ResolvedFormat>(this.singleLayerFormatCacheMaxSize);
    this.sheetCellResolvedFormatCache = new LruCache<number, ResolvedFormat>(this.singleLayerFormatCacheMaxSize);
    this.resolvedFormatCache = new LruCache<string, ResolvedFormat>(this.resolvedFormatCacheMaxSize);
    // Caches cover cell metadata + value formatting work. Keep each sheet bounded to
    // avoid memory blow-ups on huge scrolls.

    const cssVarRoot = options.cssVarRoot ?? null;
    const headerFontFamily = cssVarRoot
      ? resolveCssVar("--font-sans", { root: cssVarRoot, fallback: DEFAULT_GRID_FONT_FAMILY })
      : resolveCssVar("--font-sans", { fallback: DEFAULT_GRID_FONT_FAMILY });
    this.headerStyle = {
      fontFamily: headerFontFamily,
      fontWeight: "600",
      textAlign: "center"
    };
    this.rowHeaderStyle = {
      fontFamily: headerFontFamily,
      fontWeight: "600",
      textAlign: "end"
    };
  }

  private resolveStyle(styleId: unknown): CellStyle | undefined {
    const id = typeof styleId === "number" && Number.isInteger(styleId) && styleId >= 0 ? styleId : 0;
    if (this.styleCache.has(id)) return this.styleCache.get(id);

    const docStyle: DocStyle = this.options.document?.styleTable?.get?.(id) ?? {};
    const style = this.convertDocStyleToGridStyle(docStyle);
    this.styleCache.set(id, style);
    return style;
  }

  private resolveResolvedFormat(sheetId: string, coord: { row: number; col: number }, cellStyleId: number): {
    style: CellStyle | undefined;
    numberFormat: string | null;
  } {
    const controller: any = this.options.document as any;

    const getNumberFormat = (docStyle: any): string | null => getStyleNumberFormat(docStyle);

    const styleTable: any = (this.options.document as any)?.styleTable;

    // Prefer stable style-id tuples when available so we can cache formatting results.
    // If the underlying style table is available we can avoid calling `getCellFormat()`
    // entirely; otherwise we still cache the resolved format by tuple and only call
    // `getCellFormat()` once per unique tuple.
    if (typeof controller.getCellFormatStyleIds === "function") {
      const ids = controller.getCellFormatStyleIds(sheetId, coord) as unknown;
      if (Array.isArray(ids) && ids.length >= 4) {
        const normalizeId = (value: unknown): number =>
          typeof value === "number" && Number.isInteger(value) && value >= 0 ? value : 0;

        const sheetDefaultStyleId = normalizeId(ids[0]);
        const rowStyleId = normalizeId(ids[1]);
        const colStyleId = normalizeId(ids[2]);
        const cellStyleId = normalizeId(ids[3]);
        const rangeRunStyleId = normalizeId(ids[4]);

        // Fast paths:
        // - Completely default formatting (`0` everywhere) is common for unformatted workbooks.
        // - Sheet-level default formatting (`sheetDefaultStyleId != 0`, others `0`) is common after
        //   applying global formatting (default font / number format).
        if (colStyleId === 0 && rowStyleId === 0 && rangeRunStyleId === 0 && cellStyleId === 0) {
          if (sheetDefaultStyleId === 0) return DEFAULT_RESOLVED_FORMAT;
          const cached = this.sheetDefaultResolvedFormatCache.get(sheetDefaultStyleId);
          if (cached) return cached;

          if (typeof styleTable?.get === "function") {
            const sheetStyle = styleTable.get(sheetDefaultStyleId);
            const out = { style: this.resolveStyle(sheetDefaultStyleId), numberFormat: getNumberFormat(sheetStyle) };
            this.sheetDefaultResolvedFormatCache.set(sheetDefaultStyleId, out);
            return out;
          }

          if (typeof controller.getCellFormat === "function") {
            const resolvedDocStyle: unknown = controller.getCellFormat(sheetId, coord);
            const docStyle: DocStyle = isPlainObject(resolvedDocStyle) ? (resolvedDocStyle as DocStyle) : {};
            const out = { style: this.convertDocStyleToGridStyle(docStyle), numberFormat: getNumberFormat(docStyle) };
            this.sheetDefaultResolvedFormatCache.set(sheetDefaultStyleId, out);
            return out;
          }
        }

        // Fast path for “single-layer” formatting: when only one formatting layer is active
        // (row/col/range-run/cell), we can cache by `(sheetDefaultStyleId, layerStyleId)` without
        // building a 5-tuple key string.
        //
        // Note: We only take this path when the cache key can be encoded as a safe integer.
        const resolveSingleLayer = (
          layerId: number,
          cache: LruCache<number, ResolvedFormat>
        ): ResolvedFormat | null => {
          if (sheetDefaultStyleId >= STYLE_ID_PAIR_STRIDE || layerId >= STYLE_ID_PAIR_STRIDE) return null;
          const key = sheetDefaultStyleId * STYLE_ID_PAIR_STRIDE + layerId;
          const cached = cache.get(key);
          if (cached !== undefined) return cached;

          if (typeof styleTable?.get === "function") {
            // When the sheet default is 0, the effective style is just the layer style.
            if (sheetDefaultStyleId === 0) {
              const layerStyle = styleTable.get(layerId);
              const out = { style: this.resolveStyle(layerId), numberFormat: getNumberFormat(layerStyle) };
              cache.set(key, out);
              return out;
            }

            const sheetStyle = styleTable.get(sheetDefaultStyleId);
            const layerStyle = styleTable.get(layerId);
            const merged = applyStylePatch(sheetStyle, layerStyle);
            const out = { style: this.convertDocStyleToGridStyle(merged), numberFormat: getNumberFormat(merged) };
            cache.set(key, out);
            return out;
          }

          if (typeof controller.getCellFormat === "function") {
            const resolvedDocStyle: unknown = controller.getCellFormat(sheetId, coord);
            const docStyle: DocStyle = isPlainObject(resolvedDocStyle) ? (resolvedDocStyle as DocStyle) : {};
            const out = { style: this.convertDocStyleToGridStyle(docStyle), numberFormat: getNumberFormat(docStyle) };
            cache.set(key, out);
            return out;
          }

          return null;
        };

        if (rowStyleId === 0 && rangeRunStyleId === 0 && cellStyleId === 0 && colStyleId !== 0) {
          const resolved = resolveSingleLayer(colStyleId, this.sheetColResolvedFormatCache);
          if (resolved) return resolved;
        } else if (colStyleId === 0 && rangeRunStyleId === 0 && cellStyleId === 0 && rowStyleId !== 0) {
          const resolved = resolveSingleLayer(rowStyleId, this.sheetRowResolvedFormatCache);
          if (resolved) return resolved;
        } else if (colStyleId === 0 && rowStyleId === 0 && cellStyleId === 0 && rangeRunStyleId !== 0) {
          const resolved = resolveSingleLayer(rangeRunStyleId, this.sheetRunResolvedFormatCache);
          if (resolved) return resolved;
        } else if (colStyleId === 0 && rowStyleId === 0 && rangeRunStyleId === 0 && cellStyleId !== 0) {
          const resolved = resolveSingleLayer(cellStyleId, this.sheetCellResolvedFormatCache);
          if (resolved) return resolved;
        }

        // Key order matches merge precedence `sheet < col < row < range-run < cell`.
        const key = `${sheetDefaultStyleId},${colStyleId},${rowStyleId},${rangeRunStyleId},${cellStyleId}`;
        const cached = this.resolvedFormatCache.get(key);
        if (cached !== undefined) return cached;

        if (typeof styleTable?.get === "function") {
          const sheetStyle = styleTable.get(sheetDefaultStyleId);
          const colStyle = styleTable.get(colStyleId);
          const rowStyle = styleTable.get(rowStyleId);
          const runStyle = styleTable.get(rangeRunStyleId);
          const cellStyle = styleTable.get(cellStyleId);

          // Precedence: sheet < col < row < range-run < cell.
          const sheetCol = applyStylePatch(sheetStyle, colStyle);
          const sheetColRow = applyStylePatch(sheetCol, rowStyle);
          const sheetColRowRun = applyStylePatch(sheetColRow, runStyle);
          const merged = applyStylePatch(sheetColRowRun, cellStyle);

          const style = this.convertDocStyleToGridStyle(merged);
          const numberFormat = getNumberFormat(merged);
          const out = { style, numberFormat };
          this.resolvedFormatCache.set(key, out);
          return out;
        }

        if (typeof controller.getCellFormat === "function") {
          const resolvedDocStyle: unknown = controller.getCellFormat(sheetId, coord);
          const docStyle: DocStyle = isPlainObject(resolvedDocStyle) ? (resolvedDocStyle as DocStyle) : {};
          const out = { style: this.convertDocStyleToGridStyle(docStyle), numberFormat: getNumberFormat(docStyle) };
          this.resolvedFormatCache.set(key, out);
          return out;
        }
      }
    }

    if (typeof controller.getCellFormat === "function") {
      const resolvedDocStyle: unknown = controller.getCellFormat(sheetId, coord);
      const docStyle: DocStyle = isPlainObject(resolvedDocStyle) ? (resolvedDocStyle as DocStyle) : {};
      return { style: this.convertDocStyleToGridStyle(docStyle), numberFormat: getNumberFormat(docStyle) };
    }

    // Legacy fallback: per-cell styleId only.
    const docStyle: DocStyle = this.options.document?.styleTable?.get?.(cellStyleId) ?? {};
    return { style: this.resolveStyle(cellStyleId), numberFormat: getNumberFormat(docStyle) };
  }

  private convertDocStyleToGridStyle(docStyle: unknown): CellStyle | undefined {
    if (!isPlainObject(docStyle)) return undefined;

    // Note: `@formula/grid` CellStyle is evolving; the shared-grid rendering pipeline reads
    // additional formatting primitives (borders, underline, etc.) off this object at runtime.
    // We intentionally build this as a plain object and cast at the end to avoid tight coupling
    // to the exact type shape.
    const out: any = {};

    // Fill colors appear in several legacy and modern shapes:
    // - Modern UI: `{ fill: { pattern: "solid", fgColor: ... } }`
    // - formula-model/XLSX import: `{ fill: { fg_color: ... } }`
    // - legacy/clipboard-ish snapshots: `{ backgroundColor: ... }`
    //
    // When the UI clears fill by setting `fill: null`, treat it as authoritative and do not
    // fall back to legacy `backgroundColor` keys.
    const fillRaw = hasOwn(docStyle, "fill") ? (docStyle as any).fill : undefined;
    const fill = isPlainObject(fillRaw) ? fillRaw : null;
    let fillColorInput: unknown = undefined;
    let fillOverride = false;

    if (hasOwn(docStyle, "fill")) {
      if (fillRaw == null) {
        fillOverride = true;
        fillColorInput = null;
      } else if (fill) {
        if (hasOwn(fill, "fgColor")) {
          fillOverride = true;
          fillColorInput = (fill as any).fgColor;
        } else if (hasOwn(fill, "fg_color")) {
          fillOverride = true;
          fillColorInput = (fill as any).fg_color;
        } else if (hasOwn(fill, "background")) {
          fillOverride = true;
          fillColorInput = (fill as any).background;
        } else if (hasOwn(fill, "bgColor")) {
          fillOverride = true;
          fillColorInput = (fill as any).bgColor;
        } else if (hasOwn(fill, "bg_color")) {
          fillOverride = true;
          fillColorInput = (fill as any).bg_color;
        }
      }
    }

    if (!fillOverride) {
      fillColorInput =
        fill?.fgColor ??
        fill?.fg_color ??
        fill?.background ??
        fill?.bgColor ??
        fill?.bg_color ??
        (docStyle as any).backgroundColor ??
        (docStyle as any).background_color ??
        (docStyle as any).fillColor ??
        (docStyle as any).fill_color;
    }

    const fillColor = normalizeCssColor(fillColorInput);
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

    // Baseline shift / font vertical position (Excel subscript/superscript).
    // OOXML uses `font.vertAlign` with values like "subscript" / "superscript".
    const vertAlignRaw =
      font && hasOwn(font, "vertAlign")
        ? (font as any).vertAlign
        : font && hasOwn(font, "vert_align")
          ? (font as any).vert_align
          : undefined;
    const vertAlign = typeof vertAlignRaw === "string" ? vertAlignRaw.toLowerCase() : null;
    if (vertAlign === "subscript" || vertAlign === "superscript") out.fontVariantPosition = vertAlign;

    const fontName =
      typeof font?.name === "string"
        ? font.name
        : typeof (docStyle as any).fontFamily === "string"
          ? (docStyle as any).fontFamily
          : typeof (docStyle as any).font_family === "string"
            ? (docStyle as any).font_family
            : typeof (docStyle as any).fontName === "string"
              ? (docStyle as any).fontName
              : typeof (docStyle as any).font_name === "string"
                ? (docStyle as any).font_name
                : null;
    if (fontName && fontName.trim() !== "") out.fontFamily = fontName;

    let fontSizePx: number | null = null;
    // Prefer the UI-style `font.size` (pt) when present so user edits override imported `size_100pt`.
    if (font && hasOwn(font, "size")) {
      const size = (font as any).size;
      if (typeof size === "number" && Number.isFinite(size)) {
        fontSizePx = (size * 96) / 72;
      }
      // If `size` exists but is not a finite number (e.g. explicitly cleared/null),
      // treat that as an override and do not fall back to imported values.
    } else if (font && hasOwn(font, "size_100pt")) {
      const size100 = (font as any).size_100pt;
      if (typeof size100 === "number" && Number.isFinite(size100)) {
        // formula-model / XLSX import serializes font sizes in 1/100th of a point.
        // Convert to CSS pixels assuming 96DPI.
        const pt = size100 / 100;
        fontSizePx = (pt * 96) / 72;
      }
    } else {
      const fontSizePt =
        typeof (docStyle as any).fontSize === "number"
          ? (docStyle as any).fontSize
          : typeof (docStyle as any).font_size === "number"
            ? (docStyle as any).font_size
            : null;
      if (fontSizePt != null && Number.isFinite(fontSizePt)) {
        fontSizePx = (fontSizePt * 96) / 72;
      }
    }
    if (fontSizePx != null) out.fontSize = fontSizePx;
    // Like other style attributes, font colors can come from different serialization shapes.
    // The UI clears back to "Automatic" by setting `font.color: null`, so treat the presence of
    // that key (even when null) as authoritative and do not fall back to legacy flat keys.
    let fontColorInput: unknown = undefined;
    if (font && hasOwn(font, "color")) fontColorInput = (font as any).color;
    else if (hasOwn(docStyle, "textColor")) fontColorInput = (docStyle as any).textColor;
    else if (hasOwn(docStyle, "text_color")) fontColorInput = (docStyle as any).text_color;
    else if (hasOwn(docStyle, "fontColor")) fontColorInput = (docStyle as any).fontColor;
    else if (hasOwn(docStyle, "font_color")) fontColorInput = (docStyle as any).font_color;
    else fontColorInput = font?.color;
    const fontColor = normalizeCssColor(fontColorInput);
    if (fontColor) out.color = fontColor;

    const numberFormat = getStyleNumberFormat(docStyle);
    if (numberFormat != null) out.numberFormat = numberFormat;

    const alignment = isPlainObject(docStyle.alignment) ? docStyle.alignment : null;
    // Horizontal alignment may exist in several shapes depending on provenance:
    // - UI patches: `alignment.horizontal`
    // - formula-model/XLSX import: `alignment.horizontal_alignment` / `horizontal_align`
    // - legacy/clipboard-ish flat keys: `horizontal_align` etc.
    //
    // When the UI explicitly sets `alignment.horizontal: null` to clear the imported alignment,
    // treat that as authoritative and do not fall back to snake_case/flat keys.
    let horizontalRaw: unknown = undefined;
    if (hasOwn(alignment, "horizontal")) horizontalRaw = (alignment as any).horizontal;
    else if (hasOwn(alignment, "horizontal_align")) horizontalRaw = (alignment as any).horizontal_align;
    else if (hasOwn(alignment, "horizontal_alignment")) horizontalRaw = (alignment as any).horizontal_alignment;
    else if (hasOwn(alignment, "horizontalAlign")) horizontalRaw = (alignment as any).horizontalAlign;
    else if (hasOwn(alignment, "horizontalAlignment")) horizontalRaw = (alignment as any).horizontalAlignment;
    else if (hasOwn(docStyle, "horizontalAlign")) horizontalRaw = (docStyle as any).horizontalAlign;
    else if (hasOwn(docStyle, "horizontal_align")) horizontalRaw = (docStyle as any).horizontal_align;
    else if (hasOwn(docStyle, "horizontalAlignment")) horizontalRaw = (docStyle as any).horizontalAlignment;
    else if (hasOwn(docStyle, "horizontal_alignment")) horizontalRaw = (docStyle as any).horizontal_alignment;

    const horizontal = typeof horizontalRaw === "string" ? horizontalRaw.toLowerCase() : null;
    if (horizontal === "center") out.textAlign = "center";
    else if (horizontal === "left") out.textAlign = "start";
    else if (horizontal === "right") out.textAlign = "end";
    else if (horizontal === "fill" || horizontal === "justify") {
      // @formula/grid only supports CanvasTextAlign today, but Excel's `fill`/`justify`
      // horizontal alignment values exist in formula-model serialization. Map them to
      // a deterministic fallback ("start") and preserve the semantic in a separate field
      // for renderers that opt into it.
      out.textAlign = "start";
      out.horizontalAlign = horizontal;
    }
    // "general"/undefined: leave undefined so renderer can pick based on value type.
    const wrapRaw = hasOwn(alignment, "wrapText")
      ? (alignment as any).wrapText
      : hasOwn(alignment, "wrap_text")
        ? (alignment as any).wrap_text
        : undefined;
    if (wrapRaw === true) out.wrapMode = "word";

    const indentRaw = (alignment as any)?.indent;
    const indentLevel =
      typeof indentRaw === "number"
        ? indentRaw
        : typeof indentRaw === "string" && indentRaw.trim() !== ""
          ? Number(indentRaw)
          : null;
    if (indentLevel != null && Number.isFinite(indentLevel)) {
      // Cap at a reasonable max to avoid pathological values creating huge padding.
      const textIndentPx = clamp(indentLevel, 0, 250) * INDENT_STEP_PX;
      if (textIndentPx > 0) out.textIndentPx = textIndentPx;
    }

    // Like horizontal alignment, vertical alignment may arrive in camelCase or snake_case forms.
    // Treat an explicit `alignment.vertical` (including null) as authoritative so callers can
    // clear imported `vertical_alignment` values.
    let verticalRaw: unknown = undefined;
    if (hasOwn(alignment, "vertical")) verticalRaw = (alignment as any).vertical;
    else if (hasOwn(alignment, "vertical_align")) verticalRaw = (alignment as any).vertical_align;
    else if (hasOwn(alignment, "vertical_alignment")) verticalRaw = (alignment as any).vertical_alignment;
    else if (hasOwn(alignment, "verticalAlign")) verticalRaw = (alignment as any).verticalAlign;
    else if (hasOwn(alignment, "verticalAlignment")) verticalRaw = (alignment as any).verticalAlignment;
    else if (hasOwn(docStyle, "verticalAlign")) verticalRaw = (docStyle as any).verticalAlign;
    else if (hasOwn(docStyle, "vertical_align")) verticalRaw = (docStyle as any).vertical_align;
    else if (hasOwn(docStyle, "verticalAlignment")) verticalRaw = (docStyle as any).verticalAlignment;
    else if (hasOwn(docStyle, "vertical_alignment")) verticalRaw = (docStyle as any).vertical_alignment;
    else verticalRaw = alignment?.vertical;

    const vertical = typeof verticalRaw === "string" ? verticalRaw.toLowerCase() : null;
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
      const cssVarRoot = this.options.cssVarRoot ?? null;
      const defaultBorderColor = cssVarRoot
        ? resolveCssVar("--text-primary", { root: cssVarRoot, fallback: "CanvasText" })
        : resolveCssVar("--text-primary", { fallback: "CanvasText" });
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
      for (const _key in borders) {
        out.borders = borders;
        break;
      }

      const normalizeDiagonalBorderColor = (color: string): string => {
        const match = /^#([0-9a-f]{6})$/i.exec(color);
        if (!match) return color;
        const hex = match[1];
        const r = Number.parseInt(hex.slice(0, 2), 16);
        const g = Number.parseInt(hex.slice(2, 4), 16);
        const b = Number.parseInt(hex.slice(4, 6), 16);
        if (![r, g, b].every((n) => Number.isFinite(n))) return color;
        return `rgba(${r},${g},${b},1)`;
      };

      const mapDiagonalEdge = (edge: any) => {
        if (!isPlainObject(edge)) return undefined;
        const mapped = mapExcelBorderStyle(edge.style);
        if (!mapped) return undefined;
        const normalized = normalizeCssColor(edge.color);
        const color = normalized ? normalizeDiagonalBorderColor(normalized) : defaultBorderColor;
        return { width: mapped.width, style: mapped.style, color };
      };

      const diagonal = mapDiagonalEdge(border.diagonal);
      const diagonalUp = (border as any).diagonalUp === true || (border as any).diagonal_up === true;
      const diagonalDown = (border as any).diagonalDown === true || (border as any).diagonal_down === true;
      if (diagonal && (diagonalUp || diagonalDown)) {
        const diagonalBorders: any = {};
        if (diagonalUp) diagonalBorders.up = diagonal;
        if (diagonalDown) diagonalBorders.down = diagonal;
        for (const _key in diagonalBorders) {
          out.diagonalBorders = diagonalBorders;
          break;
        }
      }
    }

    for (const _key in out) return out as CellStyle;
    return undefined;
  }

  private resolveLinkColor(): string {
    if (this.resolvedLinkColor != null) return this.resolvedLinkColor;
    // Canvas renderers cannot consume raw `var(--token)` values. Additionally, system
    // color keywords (e.g. `LinkText`) are not always accepted by canvas color parsers,
    // so best-effort normalize the token into a computed `rgb(...)` string via a
    // hidden DOM probe when available.
    //
    // We intentionally use a system color fallback so unit tests / non-DOM
    // environments remain readable.
    const cssVarRoot = this.options.cssVarRoot ?? null;
    let resolved = cssVarRoot
      ? resolveCssVar("--formula-grid-link", { root: cssVarRoot, fallback: "LinkText" })
      : resolveCssVar("--formula-grid-link", { fallback: "LinkText" });

    const root = cssVarRoot ?? globalThis?.document?.documentElement ?? null;
    if (root && typeof getComputedStyle === "function") {
      const probe = root.ownerDocument.createElement("div");
      probe.style.position = "absolute";
      probe.style.width = "0";
      probe.style.height = "0";
      probe.style.overflow = "hidden";
      probe.style.pointerEvents = "none";
      probe.style.visibility = "hidden";
      probe.style.setProperty("contain", "strict");

      root.appendChild(probe);
      try {
        probe.style.color = "";
        probe.style.color = resolved;

        // If the assignment was rejected (invalid value), keep the raw token
        // (it may still be valid in other environments).
        if (probe.style.color) {
          const computed = getComputedStyle(probe).color;
          const normalized = computed?.trim();
          if (normalized) resolved = normalized;
        }
      } finally {
        probe.remove();
      }
    }

    this.resolvedLinkColor = resolved;
    return resolved;
  }

  private resolveDefaultHyperlinkStyle(): CellStyle {
    if (this.resolvedDefaultHyperlinkStyle != null) return this.resolvedDefaultHyperlinkStyle;
    const style: any = {
      color: this.resolveLinkColor(),
      underline: true,
    };
    this.resolvedDefaultHyperlinkStyle = style as CellStyle;
    return style as CellStyle;
  }

  invalidateAll(): void {
    if (this.disposed) return;
    this.sheetCaches.clear();
    this.lastSheetId = null;
    this.lastSheetCache = null;
    // Some style conversions (e.g. default border colors) resolve theme CSS vars
    // into concrete canvas colors, so flush style caches on full invalidation to
    // ensure theme switches re-render correctly.
    this.styleCache.clear();
    this.sheetDefaultResolvedFormatCache.clear();
    this.sheetColResolvedFormatCache.clear();
    this.sheetRowResolvedFormatCache.clear();
    this.sheetRunResolvedFormatCache.clear();
    this.sheetCellResolvedFormatCache.clear();
    this.resolvedFormatCache.clear();
    this.resolvedDefaultHyperlinkStyle = null;
    this.resolvedLinkColor = null;
    this.mergedRangesBySheet.clear();
    this.mergedEpochBySheet.clear();
    for (const listener of this.listeners) listener({ type: "invalidateAll" });
  }

  invalidateDocCells(range: { startRow: number; endRow: number; startCol: number; endCol: number }): boolean {
    if (this.disposed) return false;
    const { headerRows, headerCols } = this.options;
    const gridRange: CellRange = {
      startRow: range.startRow + headerRows,
      endRow: range.endRow + headerRows,
      startCol: range.startCol + headerCols,
      endCol: range.endCol + headerCols
    };

    // Best-effort cache eviction for the affected region.
    const height = Math.max(0, gridRange.endRow - gridRange.startRow);
    const width = Math.max(0, gridRange.endCol - gridRange.startCol);
    const cellCount = height * width;
    const sheetId = this.options.getSheetId();
    const cache = this.sheetCaches.get(sheetId);

    // Strategy:
    // - Direct-evict per-cell cache keys for invalidations up to `maxDirectEvictions`.
    // - For larger invalidations, scan the bounded LRU cache and evict only matching entries.
    // - For *huge* invalidations (large relative to our bounded cache size), fall back to `invalidateAll()`
    //   rather than scanning. This keeps worst-case invalidation work predictable (Map.clear + a single
    //   update) even when a user formats a massive rectangle.
    //
    // Scanning stays cheap because the provider cache is capped, and it avoids dropping unrelated
    // cached cells when a user formats a large rectangle far away from the current viewport.
    const maxDirectEvictions = 50_000;
    const shouldEvictDirectly = cellCount <= maxDirectEvictions;

    if (cache) {
      if (shouldEvictDirectly) {
        for (let r = gridRange.startRow; r < gridRange.endRow; r++) {
          const base = r * CACHE_KEY_COL_STRIDE;
          for (let c = gridRange.startCol; c < gridRange.endCol; c++) {
            cache.delete(base + c);
          }
        }
      } else {
        // If the invalidation covers an enormous number of cells relative to what we can cache,
        // it's usually not worth scanning the entire LRU to preserve unrelated entries.
        // Clearing caches is faster and gives more predictable latency.
        // Keep the threshold > `maxDirectEvictions` so medium-sized invalidations (where preserving
        // unrelated cached cells is still valuable) continue to take the scan+selective-evict path.
        const hugeInvalidationThreshold = Math.max(
          // Keep medium/large invalidations scanning the LRU (instead of dropping all caches),
          // but cap worst-case invalidation work when the invalidated region is large relative to
          // what we can even cache.
          maxDirectEvictions * 2,
          Math.floor(this.sheetCacheMaxSize * 0.5)
        );
        if (cellCount >= hugeInvalidationThreshold) {
          this.invalidateAll();
          return true;
        }
        for (const key of cache.keys()) {
          const row = Math.floor(key / CACHE_KEY_COL_STRIDE);
          const col = key - row * CACHE_KEY_COL_STRIDE;
          if (row < gridRange.startRow || row >= gridRange.endRow) continue;
          if (col < gridRange.startCol || col >= gridRange.endCol) continue;
          cache.delete(key);
        }
      }
    } else if (this.lastSheetId === sheetId) {
      // Should be rare (we normally keep a cache per active sheet), but keep the last-sheet
      // fast path consistent if something cleared the map entry.
      this.lastSheetId = null;
      this.lastSheetCache = null;
    }

    for (const listener of this.listeners) listener({ type: "cells", range: gridRange });
    return false;
  }

  prefetch(_range: CellRange): void {
    if (this.disposed) return;
    // NOTE: `prefetch` is primarily a hint for *async* providers to begin fetching
    // cell contents ahead of time.
    //
    // The desktop DocumentCellProvider is synchronous (it reads from the local
    // DocumentController and formats values on-demand). Warming the cache here is
    // redundant because CanvasGridRenderer will immediately call `getCell()` again
    // for visible cells during the same scroll+render cycle, doubling work on the
    // scroll critical path.
    //
    // Cache entries are still populated through normal `getCell()` calls during
    // rendering.
  }

  getCell(row: number, col: number): CellData | null {
    if (this.disposed) return null;
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

    // Avoid resurrecting deleted sheets: only render cells when the sheet is known to exist.
    if (this.isSheetKnownMissing(sheetId)) {
      cache.set(key, null);
      return null;
    }

    const docRow = row - headerRows;
    const docCol = col - headerCols;

    const coord = this.coordScratch;
    coord.row = docRow;
    coord.col = docCol;

    const state = this.options.document.getCell(sheetId, coord) as {
      value: unknown;
      formula: string | null;
      styleId?: number;
    };
    if (!state) {
      cache.set(key, null);
      return null;
    }

    let value: CellData["value"] = null;
    let richText: RichTextValue | undefined;
    let image: CellData["image"] | undefined;

    const applyScalarValue = (raw: unknown) => {
      if (raw == null) {
        value = null;
        return;
      }
      if (typeof raw === "string" || typeof raw === "number" || typeof raw === "boolean") {
        value = raw;
        return;
      }
      value = String(raw);
    };

    const applyImageValue = (raw: unknown): boolean => {
      const parsed = parseImageCellPayload(raw);
      if (!parsed) return false;
      image = parsed;
      value = parsed.altText ?? "[Image]";
      return true;
    };

    if (state.formula != null) {
      if (this.options.showFormulas()) {
        value = state.formula;
      } else {
        // Most formula cells should display the calc engine's computed value. However, imported XLSX
        // snapshots can include a cached rich-value image payload for IMAGE() / "Place in Cell"
        // pictures. Prefer rendering that cached in-cell image when present, falling back to the
        // computed value for all other formulas.
        if (!applyImageValue(state.value)) {
          const computed = this.options.getComputedValue(coord);
          if (!applyImageValue(computed)) {
            applyScalarValue(computed);
          }
        }
      }
    } else if (state.value != null) {
      if (applyImageValue(state.value)) {
        // Image payloads are rendered via `cell.image` (scalar `value` is for accessibility).
      } else if (isRichTextValue(state.value)) {
        richText = state.value;
        value = richText.text;
      } else {
        applyScalarValue(state.value);
      }
    }

    const styleId = typeof state.styleId === "number" ? state.styleId : 0;
    const { style, numberFormat } = this.resolveResolvedFormat(sheetId, coord, styleId);
    let resolvedStyle = style;

    if (typeof value === "number" && numberFormat !== null) {
      value = formatValueWithNumberFormat(value, numberFormat);
      // Preserve spreadsheet-like default alignment for numeric values even though we
      // render them as formatted strings (CanvasGridRenderer defaults to left-aligning strings).
      if (resolvedStyle?.textAlign === undefined) {
        if (!resolvedStyle) {
          resolvedStyle = DEFAULT_NUMERIC_ALIGNMENT_STYLE;
        } else {
          let aligned = this.numericAlignmentStyleCache.get(resolvedStyle);
          if (!aligned) {
            aligned = { ...resolvedStyle, textAlign: "end" };
            this.numericAlignmentStyleCache.set(resolvedStyle, aligned);
          }
          resolvedStyle = aligned;
        }
      }
    }

    // Apply default hyperlink styling for URL-like values so users can visually
    // identify Ctrl/Cmd+clickable links.
    //
    // This is intentionally theme-token driven (`--link` / `--formula-grid-link`)
    // so links remain readable across light/dark/high-contrast themes.
    if (richText == null && typeof value === "string" && looksLikeExternalHyperlink(value)) {
      const needsUnderline = resolvedStyle?.underline === undefined;
      const needsColor = resolvedStyle?.color === undefined;
      if (needsUnderline || needsColor) {
        if (!resolvedStyle && needsUnderline && needsColor) {
          resolvedStyle = this.resolveDefaultHyperlinkStyle();
        } else {
          const linkColor = needsColor ? this.resolveLinkColor() : null;
          if (!resolvedStyle) {
            const style: any = {};
            if (needsColor && linkColor) style.color = linkColor;
            if (needsUnderline) style.underline = true;
            resolvedStyle = style as CellStyle;
          } else {
            const next: any = { ...resolvedStyle };
            if (needsColor && linkColor) next.color = linkColor;
            if (needsUnderline) next.underline = true;
            resolvedStyle = next as CellStyle;
          }
        }
      }
    }

    const metaProvider = this.options.getCommentMeta;
    const meta = metaProvider ? metaProvider(docRow, docCol) : null;

    // Only attach optional metadata when present to avoid per-cell allocations
    // and keep CellData objects lean for fast scrolling.
    const cell: CellData = { row, col, value, style: resolvedStyle };
    if (richText) cell.richText = richText;
    if (image) cell.image = image;
    if (meta) cell.comment = meta;
    cache.set(key, cell);
    return cell;
  }

  private getSheetCache(sheetId: string): LruCache<number, CellData | null> {
    if (this.lastSheetId === sheetId && this.lastSheetCache) return this.lastSheetCache;

    let cache = this.sheetCaches.get(sheetId);
    if (!cache) {
      cache = new LruCache<number, CellData | null>(this.sheetCacheMaxSize);
      this.sheetCaches.set(sheetId, cache);
    }

    this.lastSheetId = sheetId;
    this.lastSheetCache = cache;
    return cache;
  }

  private mergedRangesEqual(left: unknown, right: unknown): boolean {
    const la = Array.isArray(left) ? (left as any[]) : [];
    const ra = Array.isArray(right) ? (right as any[]) : [];
    if (la.length !== ra.length) return false;
    for (let i = 0; i < la.length; i += 1) {
      const l = la[i];
      const r = ra[i];
      if (!l || !r) return false;
      if (l.startRow !== r.startRow) return false;
      if (l.endRow !== r.endRow) return false;
      if (l.startCol !== r.startCol) return false;
      if (l.endCol !== r.endCol) return false;
    }
    return true;
  }

  private bumpMergedEpoch(sheetId: string): void {
    this.mergedEpochBySheet.set(sheetId, (this.mergedEpochBySheet.get(sheetId) ?? 0) + 1);
    this.mergedRangesBySheet.delete(sheetId);
  }

  private mergedRanges(sheetId: string): Array<{ startRow: number; endRow: number; startCol: number; endCol: number }> {
    const epoch = this.mergedEpochBySheet.get(sheetId) ?? 0;
    const cached = this.mergedRangesBySheet.get(sheetId);
    if (cached && cached.epoch === epoch) return cached.ranges;

    // Avoid resurrecting deleted sheets by calling `DocumentController.getMergedRanges()` which
    // internally materializes sheets via `getSheetView()`.
    if (this.isSheetKnownMissing(sheetId)) {
      const ranges: Array<{ startRow: number; endRow: number; startCol: number; endCol: number }> = [];
      this.mergedRangesBySheet.set(sheetId, { epoch, ranges });
      return ranges;
    }

    const anyDoc: any = this.options.document as any;
    const raw = typeof anyDoc.getMergedRanges === "function" ? anyDoc.getMergedRanges(sheetId) : [];
    const ranges: Array<{ startRow: number; endRow: number; startCol: number; endCol: number }> = [];
    if (Array.isArray(raw)) {
      for (const entry of raw) {
        if (!entry) continue;
        const startRow = Number((entry as any).startRow);
        const endRow = Number((entry as any).endRow);
        const startCol = Number((entry as any).startCol);
        const endCol = Number((entry as any).endCol);
        if (!Number.isInteger(startRow) || startRow < 0) continue;
        if (!Number.isInteger(endRow) || endRow < 0) continue;
        if (!Number.isInteger(startCol) || startCol < 0) continue;
        if (!Number.isInteger(endCol) || endCol < 0) continue;
        ranges.push({ startRow, endRow, startCol, endCol });
      }
    }
    this.mergedRangesBySheet.set(sheetId, { epoch, ranges });
    return ranges;
  }

  getMergedRangeAt(row: number, col: number): CellRange | null {
    const { rowCount, colCount, headerRows, headerCols } = this.options;
    if (row < 0 || col < 0 || row >= rowCount || col >= colCount) return null;
    // Never treat header cells as merged.
    if (row < headerRows || col < headerCols) return null;
    const sheetId = this.options.getSheetId();
    const docRow = row - headerRows;
    const docCol = col - headerCols;
    if (docRow < 0 || docCol < 0) return null;

    for (const range of this.mergedRanges(sheetId)) {
      if (docRow < range.startRow || docRow > range.endRow) continue;
      if (docCol < range.startCol || docCol > range.endCol) continue;
      return {
        startRow: range.startRow + headerRows,
        endRow: range.endRow + headerRows + 1,
        startCol: range.startCol + headerCols,
        endCol: range.endCol + headerCols + 1,
      };
    }
    return null;
  }

  getMergedRangesInRange(range: CellRange): CellRange[] {
    const { headerRows, headerCols } = this.options;
    const sheetId = this.options.getSheetId();

    const docStartRow = range.startRow - headerRows;
    const docEndRowExclusive = range.endRow - headerRows;
    const docStartCol = range.startCol - headerCols;
    const docEndColExclusive = range.endCol - headerCols;

    // If the requested range falls entirely within header rows/cols, there are no merges to return.
    if (docEndRowExclusive <= 0 || docEndColExclusive <= 0) return [];

    const startRow = Math.max(0, docStartRow);
    const endRow = docEndRowExclusive - 1;
    const startCol = Math.max(0, docStartCol);
    const endCol = docEndColExclusive - 1;
    if (endRow < startRow || endCol < startCol) return [];

    const regions = this.mergedRanges(sheetId);
    if (!regions || regions.length === 0) return [];

    const out: CellRange[] = [];
    for (const r of regions) {
      if (!r) continue;
      // Inclusive intersection checks (our merged-region storage is inclusive).
      if (r.endRow < startRow || r.startRow > endRow) continue;
      if (r.endCol < startCol || r.startCol > endCol) continue;
      out.push({
        startRow: r.startRow + headerRows,
        endRow: r.endRow + headerRows + 1,
        startCol: r.startCol + headerCols,
        endCol: r.endCol + headerCols + 1,
      });
    }
    return out;
  }

  getCacheStats(): {
    sheetCache: { size: number; max: number };
    resolvedFormatCache: { size: number; max: number };
    sheetColResolvedFormatCache: { size: number; max: number };
    sheetRowResolvedFormatCache: { size: number; max: number };
    sheetRunResolvedFormatCache: { size: number; max: number };
    sheetCellResolvedFormatCache: { size: number; max: number };
  } {
    const sheetId = this.options.getSheetId();
    const sheetCache = this.sheetCaches.get(sheetId);

    return {
      sheetCache: { size: sheetCache?.size ?? 0, max: this.sheetCacheMaxSize },
      resolvedFormatCache: { size: this.resolvedFormatCache.size, max: this.resolvedFormatCacheMaxSize },
      sheetColResolvedFormatCache: { size: this.sheetColResolvedFormatCache.size, max: this.singleLayerFormatCacheMaxSize },
      sheetRowResolvedFormatCache: { size: this.sheetRowResolvedFormatCache.size, max: this.singleLayerFormatCacheMaxSize },
      sheetRunResolvedFormatCache: { size: this.sheetRunResolvedFormatCache.size, max: this.singleLayerFormatCacheMaxSize },
      sheetCellResolvedFormatCache: { size: this.sheetCellResolvedFormatCache.size, max: this.singleLayerFormatCacheMaxSize }
    };
  }

  subscribe(listener: (update: CellProviderUpdate) => void): () => void {
    if (this.disposed) return () => {};
    this.listeners.add(listener);

    if (!this.unsubscribeDoc) {
      // Coalesce document mutations into provider updates so the renderer can redraw
      // minimal dirty regions.
      this.unsubscribeDoc = this.options.document.on("change", (payload: any) => {
        const sheetId = this.options.getSheetId();
        const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
        const sheetViewDeltas = Array.isArray(payload?.sheetViewDeltas) ? payload.sheetViewDeltas : [];
        const imageDeltas: any[] = Array.isArray(payload?.imageDeltas)
          ? payload.imageDeltas
          : Array.isArray(payload?.imagesDeltas)
            ? payload.imagesDeltas
            : [];
        const hasImageDeltas = imageDeltas.length > 0;
        let mergedRangesChangedForVisibleSheet = false;
        if (sheetViewDeltas.length > 0) {
          for (const delta of sheetViewDeltas) {
            const id = typeof delta?.sheetId === "string" ? delta.sheetId : null;
            if (!id) continue;
            const before = (delta as any)?.before;
            const after = (delta as any)?.after;
            // Some callers emit legacy sheetView delta shapes (e.g. `{ sheetId, frozenRows }`).
            // Only treat deltas that have explicit before/after objects as eligible for merged-range updates.
            if (!before || !after) continue;
            const beforeRanges = (before as any)?.mergedRanges ?? (before as any)?.mergedCells;
            const afterRanges = (after as any)?.mergedRanges ?? (after as any)?.mergedCells;
            if (this.mergedRangesEqual(beforeRanges, afterRanges)) continue;
            this.bumpMergedEpoch(id);
            if (id === sheetId) mergedRangesChangedForVisibleSheet = true;
          }
        }

        // New layered formatting deltas (row/col/sheet style maps) may arrive without per-cell deltas.
        const rowStyleDeltas = Array.isArray(payload?.rowStyleDeltas) ? payload.rowStyleDeltas : [];
        const colStyleDeltas = Array.isArray(payload?.colStyleDeltas) ? payload.colStyleDeltas : [];
        const sheetStyleDeltas = Array.isArray(payload?.sheetStyleDeltas) ? payload.sheetStyleDeltas : [];
        // DocumentController may emit layered format updates either as `formatDeltas` (sheet/row/col)
        // or via the compressed rectangular range-run layer (`rangeRunDeltas`).
        const formatDeltas = Array.isArray(payload?.formatDeltas) ? payload.formatDeltas : [];
        const rangeRunDeltas = Array.isArray(payload?.rangeRunDeltas) ? payload.rangeRunDeltas : [];

        const recalc = payload?.recalc === true;
        const hasFormatLayerDeltas =
          rowStyleDeltas.length > 0 ||
          colStyleDeltas.length > 0 ||
          sheetStyleDeltas.length > 0 ||
          formatDeltas.length > 0 ||
          rangeRunDeltas.length > 0;

        const mergedRangesDirtyRange = (() => {
          if (sheetViewDeltas.length === 0) return null;
          let minRow = Infinity;
          let maxRow = -Infinity;
          let minCol = Infinity;
          let maxCol = -Infinity;

          const add = (range: any) => {
            const startRow = Number(range?.startRow);
            const endRow = Number(range?.endRow);
            const startCol = Number(range?.startCol);
            const endCol = Number(range?.endCol);
            if (!Number.isInteger(startRow) || startRow < 0) return;
            if (!Number.isInteger(endRow) || endRow < startRow) return;
            if (!Number.isInteger(startCol) || startCol < 0) return;
            if (!Number.isInteger(endCol) || endCol < startCol) return;
            minRow = Math.min(minRow, startRow);
            maxRow = Math.max(maxRow, endRow);
            minCol = Math.min(minCol, startCol);
            maxCol = Math.max(maxCol, endCol);
          };

          const buildMap = (raw: any) => {
            const map = new Map<string, any>();
            if (!Array.isArray(raw)) return map;
            for (const r of raw) {
              const startRow = Number(r?.startRow);
              const endRow = Number(r?.endRow);
              const startCol = Number(r?.startCol);
              const endCol = Number(r?.endCol);
              if (!Number.isInteger(startRow) || startRow < 0) continue;
              if (!Number.isInteger(endRow) || endRow < startRow) continue;
              if (!Number.isInteger(startCol) || startCol < 0) continue;
              if (!Number.isInteger(endCol) || endCol < startCol) continue;
              const key = `${startRow},${endRow},${startCol},${endCol}`;
              map.set(key, { startRow, endRow, startCol, endCol });
            }
            return map;
          };

          for (const delta of sheetViewDeltas) {
            if (!delta) continue;
            if (String(delta.sheetId ?? "") !== sheetId) continue;
            const beforeView: any = (delta as any).before;
            const afterView: any = (delta as any).after;
            const beforeMap = buildMap(beforeView?.mergedRanges ?? beforeView?.mergedCells);
            const afterMap = buildMap(afterView?.mergedRanges ?? afterView?.mergedCells);

            if (beforeMap.size === afterMap.size && beforeMap.size > 0) {
              let same = true;
              for (const key of beforeMap.keys()) {
                if (!afterMap.has(key)) {
                  same = false;
                  break;
                }
              }
              if (same) continue;
            } else if (beforeMap.size === 0 && afterMap.size === 0) {
              continue;
            }

            for (const [key, r] of beforeMap) {
              if (!afterMap.has(key)) add(r);
            }
            for (const [key, r] of afterMap) {
              if (!beforeMap.has(key)) add(r);
            }
          }

          if (minRow === Infinity || minCol === Infinity) return null;
          return { startRow: minRow, endRow: maxRow + 1, startCol: minCol, endCol: maxCol + 1 };
        })();

        // No cell deltas + no formatting deltas: preserve the sheet-view optimization.
        if (deltas.length === 0 && !hasFormatLayerDeltas) {
          // Image bytes updates should cause a redraw (to refresh in-cell image placeholders),
          // but they do not affect any cell values, formulas, or formatting layers. Avoid
          // evicting the provider caches for those cases; the renderer can re-read the same
          // CellData while resolving updated image bytes.
          if (hasImageDeltas && sheetViewDeltas.length === 0 && !recalc) {
            for (const listener of this.listeners) listener({ type: "invalidateAll" });
            return;
          }

          // Sheet view deltas (frozen panes, row/col sizes, etc.) do not affect cell contents.
          // Avoid evicting the provider cache in those cases; the renderer will be updated by
          // the view sync code (e.g. `syncFrozenPanes` / shared grid axis sync).
          if (sheetViewDeltas.length > 0 && !recalc) {
            // Merged-cell regions *do* affect rendering (text layout + gridline suppression),
            // but they arrive via sheetViewDeltas in the current document model. Ask the
            // renderer to redraw without flushing cell caches.
            if (mergedRangesChangedForVisibleSheet) {
              for (const listener of this.listeners) listener({ type: "invalidateAll" });
            }
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

        if (rangeRunDeltas.length > 0) {
          let minRow = Infinity;
          let maxRow = -Infinity;
          let minCol = Infinity;
          let maxCol = -Infinity;
          let saw = false;

          for (const delta of rangeRunDeltas) {
            if (!delta) continue;
            if (String(delta.sheetId ?? "") !== sheetId) continue;

            const col = Number((delta as any).col);
            if (!Number.isInteger(col) || col < 0) continue;

            const startRow = Number((delta as any).startRow);
            const endRowExclusive = Number((delta as any).endRowExclusive);
            if (!Number.isInteger(startRow) || startRow < 0) continue;
            if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow) continue;

            saw = true;
            minRow = Math.min(minRow, startRow);
            maxRow = Math.max(maxRow, endRowExclusive - 1);
            minCol = Math.min(minCol, col);
            maxCol = Math.max(maxCol, col);
          }

          if (!saw) {
            // If we can't determine the affected region, fall back.
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

        if (mergedRangesDirtyRange) {
          invalidateRanges.push({
            startRow: clampInt(mergedRangesDirtyRange.startRow, 0, docRowCount),
            endRow: clampInt(mergedRangesDirtyRange.endRow, 0, docRowCount),
            startCol: clampInt(mergedRangesDirtyRange.startCol, 0, docColCount),
            endCol: clampInt(mergedRangesDirtyRange.endCol, 0, docColCount)
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
          if (this.invalidateDocCells(range)) return;
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

  /**
   * Best-effort teardown for tests/hot-reload.
   *
   * Clears per-sheet caches and unsubscribes from DocumentController change events so a destroyed
   * grid view does not retain large formatting/value caches if the provider instance remains
   * referenced.
   */
  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    try {
      this.unsubscribeDoc?.();
    } catch {
      // ignore
    }
    this.unsubscribeDoc = null;
    this.listeners.clear();
    this.sheetCaches.clear();
    this.lastSheetId = null;
    this.lastSheetCache = null;
    this.styleCache.clear();
    this.sheetDefaultResolvedFormatCache.clear();
    this.sheetColResolvedFormatCache.clear();
    this.sheetRowResolvedFormatCache.clear();
    this.sheetRunResolvedFormatCache.clear();
    this.sheetCellResolvedFormatCache.clear();
    this.resolvedFormatCache.clear();
    this.resolvedDefaultHyperlinkStyle = null;
    this.resolvedLinkColor = null;
    this.mergedRangesBySheet.clear();
    this.mergedEpochBySheet.clear();
  }
}
