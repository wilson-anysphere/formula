import type { CellChange } from "./protocol.ts";
import { toA1 } from "./backend/a1.ts";
import { normalizeFormulaTextOpt } from "./backend/formula.ts";
import { docColWidthPxToExcelChars } from "./excelColumnWidth.ts";

export type EngineCellScalar = number | string | boolean | null;

export type EngineSheetJson = {
  /**
   * Sparse cell map keyed by A1 address.
   */
  cells: Record<string, EngineCellScalar>;
  /**
   * Optional UI sheet view metadata (DocumentController base units, zoom=1).
   *
   * These keys are forwarded through the engine hydration path so functions like
   * `CELL("width")` can consult worksheet metadata once engine-side wiring is in place.
   */
  colWidths?: Record<string, number>;
  rowHeights?: Record<string, number>;
  /**
   * Optional logical worksheet dimensions (row count).
   *
   * When set, this controls how whole-column/row references like `A:A` / `1:1`
   * are expanded by the WASM engine.
   */
  rowCount?: number;
  /**
   * Optional logical worksheet dimensions (column count).
   */
  colCount?: number;
};

export type EngineWorkbookJson = {
  /**
   * Optional workbook formula locale identifier (e.g. `"en-US"`, `"de-DE"`).
   *
   * When provided, `crates/formula-wasm` will configure the workbook locale prior to importing
   * formulas from JSON so localized formula inputs can be canonicalized correctly.
   */
  localeId?: string;
  /**
   * Optional workbook legacy Windows text codepage (e.g. `932` for Shift-JIS).
   *
   * This is used by Excel's DBCS (`*B`) text functions (e.g. `LENB`) and by `ASC`/`DBCS`.
   * When omitted, the WASM engine defaults to `1252` (en-US).
   */
  textCodepage?: number;
  /**
   * Optional sheet tab order (array of sheet identifiers).
   *
   * When present, `crates/formula-wasm` should create sheets in this order so sheet-indexed
   * semantics (notably 3D references like `Sheet1:Sheet3!A1` and worksheet functions like
   * `SHEET()`) match the UI's sheet ordering.
   *
    * The identifiers in this list must match the keys in `sheets` (i.e. the engine-facing sheet id).
    *
    * In the desktop app, this is the DocumentController stable `sheetId`.
    */
  sheetOrder?: string[];
  /**
   * Worksheet map keyed by sheet identifier.
   *
   * In the desktop app, this should use the DocumentController stable sheet id (`sheetId`).
   *
   * User-facing sheet tab names are synchronized separately via `EngineSyncTarget.setSheetDisplayName`,
   * allowing sheet rename events to update worksheet-info functions (e.g. `CELL("filename")`)
   * without rebuilding the entire engine workbook.
   */
  sheets: Record<string, EngineSheetJson>;
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

export type DocumentSheetMetaState = {
  name: string;
  // Preserve additional metadata for forward compatibility; engine sync currently only cares
  // about `name`.
  visibility?: unknown;
  tabColor?: unknown;
};

export type DocumentSheetMetaDelta = {
  sheetId: string;
  before: DocumentSheetMetaState | null;
  after: DocumentSheetMetaState | null;
};

export interface EngineSyncTarget {
  loadWorkbookFromJson: (json: string) => Promise<void> | void;
  setCell: (address: string, value: EngineCellScalar, sheet?: string) => Promise<void> | void;
  setCells?: (
    updates: Array<{ address: string; value: EngineCellScalar; sheet?: string }>,
  ) => Promise<void> | void;
  /**
   * Rename a worksheet and rewrite formulas that reference it (Excel-like).
   *
   * This is optional because older WASM builds may not expose the API.
   */
  renameSheet?: (oldName: string, newName: string) => Promise<boolean> | boolean;
  recalculate: (sheet?: string) => Promise<CellChange[]> | CellChange[];
  /**
   * Update a sheet's user-visible display (tab) name without changing its stable id/key.
   *
   * When provided, this is used to keep the engine's sheet-name-emitting functions (e.g.
   * `CELL("address")`) in sync with DocumentController sheet renames.
   */
  setSheetDisplayName?: (sheetId: string, name: string) => Promise<void> | void;
  /**
   * Update workbook-level file metadata used by Excel-compatible functions like `CELL("filename")`
   * and `INFO("directory")`.
   *
   * This is optional because older WASM builds may not expose the API.
   */
  setWorkbookFileMetadata?: (directory: string | null, filename: string | null) => Promise<void> | void;
  /**
   * Update a worksheet's column width metadata (OOXML `col/@width` units).
   *
   * The engine treats this as sheet-scoped metadata that can affect worksheet information
   * functions like `CELL("width")`.
   *
   * - `sheet` is the worksheet id/name (e.g. `"Sheet1"`).
   * - `col` is 0-based (A=0).
   * - `widthChars` is expressed in Excel character units (1.0 = width of one '0' digit),
   *   matching OOXML.
   * - `null` clears the override back to the sheet default.
   */
  setColWidth?: (sheet: string, col: number, widthChars: number | null) => Promise<void> | void;
  /**
   * Optional row/col/sheet formatting layer metadata APIs.
   *
   * These are additive and may not be implemented by all engine targets.
   * Sync helpers must treat them as best-effort.
   */
  internStyle?: (styleObj: unknown) => Promise<number> | number;
  /**
   * Replace the range-run formatting runs for a column (DocumentController `formatRunsByCol`).
   *
   * Runs are expressed as half-open row intervals `[startRow, endRowExclusive)`.
   */
  setColFormatRuns?: (
    sheet: string,
    col: number,
    runs: Array<{ startRow: number; endRowExclusive: number; styleId: number }>,
  ) => Promise<void> | void;
  setCellStyleId?: (sheet: string, address: string, styleId: number) => Promise<void> | void;
  setRowStyleId?: (sheet: string, row: number, styleId: number | null) => Promise<void> | void;
  setColStyleId?: (sheet: string, col: number, styleId: number | null) => Promise<void> | void;
  setSheetDefaultStyleId?: (sheet: string, styleId: number | null) => Promise<void> | void;
  /**
   * Update a sheet's compressed range-run formatting layer (`DocumentController`'s `formatRunsByCol`).
   *
   * Runs are half-open row intervals `[startRow, endRowExclusive)` and reference an engine style id.
   *
   * This is an additive API and may not be implemented by all engine targets.
   */
  setFormatRunsByCol?: (
    sheet: string,
    col: number,
    runs: Array<{ startRow: number; endRowExclusive: number; styleId: number }>,
  ) => Promise<void> | void;
  /**
   * Set (or clear) a per-column width override for a sheet.
   *
   * `widthChars` is expressed in Excel "character" units (OOXML `col/@width`), not pixels.
   */
  setColWidthChars?: (sheet: string, col: number, widthChars: number | null) => Promise<void> | void;
}

export type EngineApplyDocumentChangeOptions = {
  /**
   * Whether to call `engine.recalculate()` after applying the change payload.
   *
   * Defaults to the DocumentController-provided `payload.recalc` flag when available
   * (`true`/`false`), with a backwards-compatible fallback of `true` when omitted.
   */
  recalculate?: boolean;
  /**
   * Resolve a DocumentController `styleId` into a style object suitable for `internStyle`.
   *
   * If not provided, style deltas are ignored (backwards compatible).
   */
  getStyleById?: (styleId: number) => unknown;
  /**
   * Optional sheet-id resolver.
   *
   * DocumentController uses stable sheet ids internally. Some engine implementations may instead
   * address sheets by their user-facing display names. When provided, this function maps a
   * DocumentController `sheetId` to the sheet identifier used by the engine API surface.
   */
  sheetIdToSheet?: (sheetId: string) => string | null | undefined;
};

type StyleTableLike = { get(styleId: number): unknown };

type EngineStyleSyncContext = {
  /**
   * Reference to the document's style table, used to resolve `DocumentCellState.styleId` into the
   * style payload passed to `EngineSyncTarget.internStyle`.
   */
  styleTable: StyleTableLike;
  /**
   * Cache mapping `DocumentController` style ids → engine style ids.
   */
  docStyleIdToEngineStyleId: Map<number, number | Promise<number>>;
};

// WeakMap ensures we don't leak engine instances in long-lived apps.
const STYLE_SYNC_BY_ENGINE = new WeakMap<EngineSyncTarget, EngineStyleSyncContext>();

function getDocumentSheetIds(doc: any): string[] {
  const ids: string[] = [];
  const seen = new Set<string>();

  const push = (raw: unknown) => {
    const id = typeof raw === "string" ? raw.trim() : "";
    if (!id) return;
    if (seen.has(id)) return;
    seen.add(id);
    ids.push(id);
  };

  // Prefer the sheet ordering returned by `DocumentController.getSheetIds()` (tab order).
  const sheetIds: unknown =
    typeof doc?.getSheetIds === "function" ? (doc.getSheetIds() as string[]) : [];
  if (Array.isArray(sheetIds)) {
    for (const id of sheetIds) push(id);
  }

  // Append any sheets that exist only in metadata (defensive / backwards compatibility).
  const meta: unknown = doc?.sheetMeta;
  if (meta instanceof Map) {
    for (const id of meta.keys()) push(id);
  }

  return ids.length > 0 ? ids : ["Sheet1"];
}

function resolveEngineSheetNameForDocumentSheetId(doc: any, sheetId: string): string {
  const rawId = typeof sheetId === "string" ? sheetId.trim() : "";
  if (!rawId) return "Sheet1";
  // The WASM engine workbook is addressed by stable DocumentController sheet ids. Keep this as an
  // identity mapping so sheet rename events can be applied incrementally by updating display-name
  // metadata (see `setSheetDisplayName`).
  return rawId;
}

function resolveDocumentSheetDisplayName(doc: any, sheetId: string): string {
  const rawId = typeof sheetId === "string" ? sheetId.trim() : "";
  if (!rawId) return "Sheet1";

  const meta = typeof doc?.getSheetMeta === "function" ? doc.getSheetMeta(rawId) : null;
  const name = typeof meta?.name === "string" ? meta.name.trim() : "";
  return name || rawId;
}

function getOrCreateStyleSyncContext(
  engine: EngineSyncTarget,
  options: { getStyleById?: (styleId: number) => unknown } | undefined,
): EngineStyleSyncContext | null {
  const existing = STYLE_SYNC_BY_ENGINE.get(engine);
  if (existing) return existing;
  const getStyleById = options?.getStyleById;
  if (typeof getStyleById !== "function") return null;
  const ctx: EngineStyleSyncContext = { styleTable: { get: getStyleById }, docStyleIdToEngineStyleId: new Map() };
  STYLE_SYNC_BY_ENGINE.set(engine, ctx);
  return ctx;
}

async function resolveEngineStyleIdForDocStyleId(
  engine: EngineSyncTarget,
  ctx: EngineStyleSyncContext,
  docStyleId: number,
): Promise<number> {
  const cached = ctx.docStyleIdToEngineStyleId.get(docStyleId);
  if (cached != null) {
    return await cached;
  }

  if (!engine.internStyle) {
    throw new Error("resolveEngineStyleIdForDocStyleId: engine.internStyle is required for non-zero styles");
  }

  const style = ctx.styleTable.get(docStyleId);
  const promise = Promise.resolve(engine.internStyle(style));
  ctx.docStyleIdToEngineStyleId.set(docStyleId, promise);

  try {
    const engineStyleId = await promise;
    if (ctx.docStyleIdToEngineStyleId.get(docStyleId) === promise) {
      ctx.docStyleIdToEngineStyleId.set(docStyleId, engineStyleId);
    }
    return engineStyleId;
  } catch (err) {
    if (ctx.docStyleIdToEngineStyleId.get(docStyleId) === promise) {
      ctx.docStyleIdToEngineStyleId.delete(docStyleId);
    }
    throw err;
  }
}

export type DocumentControllerChangePayload = {
  deltas: readonly DocumentCellDelta[];
  sheetViewDeltas: readonly unknown[];
  formatDeltas: readonly unknown[];
  rowStyleDeltas: readonly unknown[];
  colStyleDeltas: readonly unknown[];
  sheetStyleDeltas: readonly unknown[];
  rangeRunDeltas: readonly unknown[];
  sheetMetaDeltas: readonly unknown[];
  sheetOrderDelta: unknown | null;
  source?: string;
  recalc?: boolean;
};

function parseRowColKey(key: string): { row: number; col: number } | null {
  const [rowStr, colStr] = key.split(",");
  const row = Number(rowStr);
  const col = Number(colStr);
  if (!Number.isInteger(row) || row < 0) return null;
  if (!Number.isInteger(col) || col < 0) return null;
  return { row, col };
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
  if (typeof cell.formula === "string") {
    const normalized = normalizeFormulaTextOpt(cell.formula);
    if (normalized != null) return normalized;
  }
  return coerceDocumentValueToScalar(cell.value);
}

function isAxisValueEqual(a: number | null, b: number | null): boolean {
  if (a === b) return true;
  if (a == null || b == null) return false;
  return Math.abs(a - b) <= 1e-6;
}

/**
 * Export the current DocumentController workbook into the JSON format consumed by
 * `crates/formula-wasm` (`WasmWorkbook.fromJson`).
 *
 * Note: empty/cleared cells are omitted from the JSON entirely (sparse semantics).
 */
export function exportDocumentToEngineWorkbookJson(
  doc: any,
  options: { localeId?: string | null } = {},
): EngineWorkbookJson {
  return exportDocumentToEngineWorkbookJsonWithOptions(doc, options);
}

function exportDocumentToEngineWorkbookJsonWithOptions(
  doc: any,
  options: { localeId?: string | null } = {},
): EngineWorkbookJson {
  const sheets: Record<string, EngineSheetJson> = {};
  const sheetOrder: string[] = [];
  const seenSheetsInOrder = new Set<string>();

  const ids = getDocumentSheetIds(doc);

  for (const sheetId of ids) {
    const engineSheetId = resolveEngineSheetNameForDocumentSheetId(doc, sheetId);
    if (!seenSheetsInOrder.has(engineSheetId)) {
      seenSheetsInOrder.add(engineSheetId);
      sheetOrder.push(engineSheetId);
    }
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

    const view = typeof doc?.getSheetView === "function" ? doc.getSheetView(sheetId) : sheet?.view;

    const sanitizeAxisMap = (raw: unknown): Record<string, number> | undefined => {
      if (!raw || typeof raw !== "object" || Array.isArray(raw)) return undefined;
      const out: Record<string, number> = {};
      for (const [key, value] of Object.entries(raw as Record<string, unknown>)) {
        const idx = Number(key);
        if (!Number.isInteger(idx) || idx < 0) continue;
        const size = Number(value);
        if (!Number.isFinite(size) || size <= 0) continue;
        out[String(idx)] = size;
      }
      return Object.keys(out).length > 0 ? out : undefined;
    };

    const colWidths = sanitizeAxisMap(view?.colWidths);
    const rowHeights = sanitizeAxisMap(view?.rowHeights);

    sheets[engineSheetId] = {
      cells,
      ...(colWidths ? { colWidths } : {}),
      ...(rowHeights ? { rowHeights } : {}),
    };
  }

  const localeIdRaw = options.localeId;
  const localeId = typeof localeIdRaw === "string" && localeIdRaw.trim() !== "" ? localeIdRaw.trim() : undefined;

  return { ...(localeId ? { localeId } : {}), sheetOrder, sheets };
}

export type EngineHydrateFromDocumentOptions = {
  /**
   * Optional formula locale id to embed in the exported workbook JSON (e.g. `"de-DE"`).
   *
   * When provided, the WASM workbook loader will apply locale-specific parsing rules while
   * importing formulas (argument separators, decimal commas, localized function names, etc).
   */
  localeId?: string | null;
  /**
   * Workbook-level file metadata used by Excel-compatible functions like `CELL("filename")`
   * and `INFO("directory")`.
   *
   * When provided, we apply it after `loadWorkbookFromJson` (which replaces the workbook state)
   * but before the initial `recalculate()` so any dependent formulas compute with the latest
   * metadata.
   */
  workbookFileMetadata?: { directory: string | null; filename: string | null } | null;
};

export async function engineHydrateFromDocument(
  engine: EngineSyncTarget,
  doc: any,
  options: EngineHydrateFromDocumentOptions = {},
): Promise<CellChange[]> {
  const workbookJson = exportDocumentToEngineWorkbookJsonWithOptions(doc, { localeId: options.localeId });
  await engine.loadWorkbookFromJson(JSON.stringify(workbookJson));

  // Hydrate sheet display (tab) names separately from stable sheet ids.
  //
  // DocumentController uses stable sheet ids for addressing/persistence, but users see (and Excel
  // functions emit) the display name. Keep the engine informed so `CELL("filename")`,
  // `CELL("address")`, and runtime name resolution match Excel semantics even after renames.
  if (typeof engine.setSheetDisplayName === "function") {
    const ids = getDocumentSheetIds(doc);
    for (const sheetId of ids) {
      const sheetKey = resolveEngineSheetNameForDocumentSheetId(doc, sheetId);
      const displayName = resolveDocumentSheetDisplayName(doc, sheetId);
      await engine.setSheetDisplayName(sheetKey, displayName);
    }
  }

  // Sync sheet view column widths (DocumentController stores base px; engine uses Excel char units).
  //
  // This ensures `CELL("width")` metadata is correct immediately after hydration (e.g. when
  // opening a document that already contains persisted column width overrides).
  const setColWidthChars =
    typeof engine.setColWidthChars === "function"
      ? engine.setColWidthChars.bind(engine)
      : typeof engine.setColWidth === "function"
        ? engine.setColWidth.bind(engine)
        : null;

  if (setColWidthChars && typeof doc?.getSheetView === "function") {
    const ids = getDocumentSheetIds(doc);

    for (const sheetId of ids) {
      const engineSheetId = resolveEngineSheetNameForDocumentSheetId(doc, sheetId);
      const view = doc.getSheetView(sheetId) as { colWidths?: Record<string, number> } | null;
      const colWidths = view?.colWidths ?? null;
      if (!colWidths) continue;

      for (const [key, rawWidth] of Object.entries(colWidths)) {
        const col = Number(key);
        if (!Number.isInteger(col) || col < 0 || col >= 16_384) continue;
        const widthPx = Number(rawWidth);
        if (!Number.isFinite(widthPx) || widthPx <= 0) continue;
        const widthChars = docColWidthPxToExcelChars(widthPx);
        // Call via `.call(engine, ...)` to preserve `this` bindings for class-based engine targets.
        await setColWidthChars.call(engine as any, engineSheetId, col, widthChars);
      }
    }
  }

  const styleTable = doc?.styleTable as StyleTableLike | null;
  if (styleTable && typeof styleTable.get === "function") {
    const ctx: EngineStyleSyncContext = { styleTable, docStyleIdToEngineStyleId: new Map() };
    STYLE_SYNC_BY_ENGINE.set(engine, ctx);

    const ids = getDocumentSheetIds(doc);

    // Sync sheet/row/col formatting layers when the engine exposes the optional hooks.
    if (engine.internStyle) {
      for (const sheetId of ids) {
        const engineSheetId = resolveEngineSheetNameForDocumentSheetId(doc, sheetId);
        // Ensure sheet model is materialized (DocumentController is lazily sheet-creating).
        doc?.model?.getCell?.(sheetId, 0, 0);
        const sheet = doc?.model?.sheets?.get?.(sheetId);
        if (!sheet) continue;

        if (engine.setSheetDefaultStyleId) {
          const docStyleId = typeof sheet?.defaultStyleId === "number" ? sheet.defaultStyleId : 0;
          if (Number.isInteger(docStyleId) && docStyleId !== 0) {
            const engineStyleId = await resolveEngineStyleIdForDocStyleId(engine, ctx, docStyleId);
            await engine.setSheetDefaultStyleId(engineSheetId, engineStyleId);
          }
        }

        if (engine.setRowStyleId && sheet?.rowStyleIds?.entries) {
          for (const [row, rawDocStyleId] of sheet.rowStyleIds.entries() as Iterable<[number, number]>) {
            if (!Number.isInteger(row) || row < 0) continue;
            const docStyleId = typeof rawDocStyleId === "number" ? rawDocStyleId : 0;
            if (!Number.isInteger(docStyleId) || docStyleId === 0) continue;
            const engineStyleId = await resolveEngineStyleIdForDocStyleId(engine, ctx, docStyleId);
            await engine.setRowStyleId(engineSheetId, row, engineStyleId);
          }
        }

        if (engine.setColStyleId && sheet?.colStyleIds?.entries) {
          for (const [col, rawDocStyleId] of sheet.colStyleIds.entries() as Iterable<[number, number]>) {
            if (!Number.isInteger(col) || col < 0) continue;
            const docStyleId = typeof rawDocStyleId === "number" ? rawDocStyleId : 0;
            if (!Number.isInteger(docStyleId) || docStyleId === 0) continue;
            const engineStyleId = await resolveEngineStyleIdForDocStyleId(engine, ctx, docStyleId);
            await engine.setColStyleId(engineSheetId, col, engineStyleId);
          }
        }

        // Sync compressed range-run formatting (`SheetModel.formatRunsByCol`) when the engine
        // exposes the optional hook. This ensures large range formatting operations are visible
        // to Excel-compatible functions like `CELL("prefix")` / `CELL("protect")` even when the
        // DocumentController did not materialize per-cell style ids.
        const setFormatRunsByCol =
          typeof engine.setFormatRunsByCol === "function"
            ? engine.setFormatRunsByCol
            : typeof engine.setColFormatRuns === "function"
              ? engine.setColFormatRuns
              : null;

        if (setFormatRunsByCol && sheet?.formatRunsByCol?.entries) {
          for (const [col, rawRuns] of sheet.formatRunsByCol.entries() as Iterable<[number, any[]]>) {
            if (!Number.isInteger(col) || col < 0 || col >= 16_384) continue;
            if (!Array.isArray(rawRuns) || rawRuns.length === 0) continue;

            const runs: Array<{ startRow: number; endRowExclusive: number; styleId: number }> = [];
            for (const run of rawRuns) {
              const startRow = Number((run as any)?.startRow);
              const endRowExclusive = Number((run as any)?.endRowExclusive);
              const docStyleId = Number((run as any)?.styleId);

              if (!Number.isInteger(startRow) || startRow < 0 || startRow >= 1_048_576) continue;
              if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow || endRowExclusive > 1_048_576) continue;
              if (!Number.isInteger(docStyleId) || docStyleId <= 0) continue;

              const engineStyleId = await resolveEngineStyleIdForDocStyleId(engine, ctx, docStyleId);
              runs.push({ startRow, endRowExclusive, styleId: engineStyleId });
            }

            if (runs.length > 0) {
              // Preserve method binding for class-based engine implementations (e.g. tests).
              await setFormatRunsByCol.call(engine as any, engineSheetId, col, runs);
            }
          }
        }
      }
    }

    // Only attempt to sync per-cell style ids if the engine exposes the optional formatting hooks.
    if (engine.setCellStyleId && engine.internStyle) {
      for (const sheetId of ids) {
        const engineSheetId = resolveEngineSheetNameForDocumentSheetId(doc, sheetId);
        const sheet = doc?.model?.sheets?.get?.(sheetId);
        if (!sheet?.cells?.entries) continue;

        // Note: we intentionally include formatting-only cells (no value/formula) so CELL() can
        // observe formatting/protection metadata even on empty cells.
        const styleUpdates: Array<Promise<void> | void> = [];
        for (const [key, cell] of sheet.cells.entries() as Iterable<[string, DocumentCellState]>) {
          const docStyleId = cell?.styleId ?? 0;
          if (!docStyleId) continue;

          const coord = parseRowColKey(key);
          if (!coord) continue;
          const address = toA1(coord.row, coord.col);

          const engineStyleId = await resolveEngineStyleIdForDocStyleId(engine, ctx, docStyleId);
          styleUpdates.push(engine.setCellStyleId(engineSheetId, address, engineStyleId));
        }
        // Apply per-sheet to keep memory bounded for large workbooks.
        if (styleUpdates.length > 0) await Promise.all(styleUpdates);
      }
    }

  }

  const metadata = options.workbookFileMetadata ?? null;
  if (metadata && typeof engine.setWorkbookFileMetadata === "function") {
    await engine.setWorkbookFileMetadata(metadata.directory ?? null, metadata.filename ?? null);
  }

  return await engine.recalculate();
}

export async function engineApplyDeltas(
  engine: EngineSyncTarget,
  deltas: readonly DocumentCellDelta[],
  options: { recalculate?: boolean; sheetIdToSheet?: (sheetId: string) => string | null | undefined } = {},
): Promise<CellChange[]> {
  const shouldRecalculate = options.recalculate ?? true;
  const updates: Array<{ address: string; value: EngineCellScalar; sheet?: string }> = [];
  const styleUpdates: Array<{ address: string; docStyleId: number; sheet?: string }> = [];
  let didApply = false;

  const resolveSheet = (sheetId: string | undefined): string | undefined => {
    if (!sheetId) return sheetId;
    const resolved = options.sheetIdToSheet?.(sheetId);
    const trimmed = typeof resolved === "string" ? resolved.trim() : "";
    return trimmed ? trimmed : sheetId;
  };

  for (const delta of deltas) {
    const beforeInput = cellStateToEngineInput(delta.before);
    const afterInput = cellStateToEngineInput(delta.after);

    const address = toA1(delta.row, delta.col);

    // Track value/formula edits (including rich-text run edits that change plain text).
    if (beforeInput !== afterInput) {
      updates.push({ address, value: afterInput, sheet: resolveSheet(delta.sheetId) });
    }

    // Track formatting-only edits when the engine supports style metadata.
    if (engine.setCellStyleId && delta.before.styleId !== delta.after.styleId) {
      styleUpdates.push({ address, docStyleId: delta.after.styleId, sheet: resolveSheet(delta.sheetId) });
    }
  }

  if (updates.length > 0) {
    if (engine.setCells) {
      await engine.setCells(updates);
    } else {
      await Promise.all(updates.map((u) => engine.setCell(u.address, u.value, u.sheet)));
    }
    didApply = true;
  }

  if (styleUpdates.length > 0 && engine.setCellStyleId) {
    const ctx = STYLE_SYNC_BY_ENGINE.get(engine);

    const setStylePromises: Array<Promise<void> | void> = [];
    for (const update of styleUpdates) {
      const sheet = update.sheet ?? "Sheet1";
      // Clearing back to the default style does not require interning.
      if (update.docStyleId === 0) {
        setStylePromises.push(engine.setCellStyleId(sheet, update.address, 0));
        continue;
      }

      if (!engine.internStyle || !ctx) {
        // Backwards compatibility: if we can't resolve the style object, ignore formatting-only deltas.
        continue;
      }

      const engineStyleId = await resolveEngineStyleIdForDocStyleId(engine, ctx, update.docStyleId);
      setStylePromises.push(engine.setCellStyleId(sheet, update.address, engineStyleId));
    }

    if (setStylePromises.length > 0) {
      await Promise.all(setStylePromises);
      didApply = true;
    }
  }

  if (!didApply) return [];
  if (!shouldRecalculate) return [];
  return await engine.recalculate();
}

export async function engineApplyDocumentChange(
  engine: EngineSyncTarget,
  changePayload: unknown,
  options: EngineApplyDocumentChangeOptions = {},
): Promise<CellChange[]> {
  // `DocumentController` emits a JSON-ish payload; keep parsing tolerant since callers may
  // pass through other event sources in tests.
  const payload = changePayload as any;
  const resolveSheet = (sheetId: string): string => {
    const resolved = typeof options.sheetIdToSheet === "function" ? options.sheetIdToSheet(sheetId) : null;
    const trimmed = typeof resolved === "string" ? resolved.trim() : "";
    return trimmed ? trimmed : sheetId;
  };

  // Handle "reset boundary" payloads by forcing callers to re-hydrate. We intentionally
  // don't attempt to mirror sheet structure changes incrementally here.
  if (payload?.source === "applyState") {
    return [];
  }

  const sheetMetaDeltas: readonly DocumentSheetMetaDelta[] = Array.isArray(payload?.sheetMetaDeltas)
    ? (payload.sheetMetaDeltas as DocumentSheetMetaDelta[])
    : [];
  const deltas: readonly DocumentCellDelta[] = Array.isArray(payload?.deltas) ? payload.deltas : [];
  const formatDeltas: unknown[] = Array.isArray(payload?.formatDeltas) ? payload.formatDeltas : [];
  const rowStyleDeltas: Array<{ sheetId: string; row: number; afterStyleId: number }> = Array.isArray(payload?.rowStyleDeltas)
    ? [...payload.rowStyleDeltas]
    : [];
  const colStyleDeltas: Array<{ sheetId: string; col: number; afterStyleId: number }> = Array.isArray(payload?.colStyleDeltas)
    ? [...payload.colStyleDeltas]
    : [];
  const sheetStyleDeltas: Array<{ sheetId: string; afterStyleId: number }> = Array.isArray(payload?.sheetStyleDeltas)
    ? [...payload.sheetStyleDeltas]
    : [];
  const sheetViewDeltas: Array<{ sheetId: string; before: any; after: any }> = Array.isArray(payload?.sheetViewDeltas)
    ? payload.sheetViewDeltas
    : [];
  const rangeRunDeltas: unknown[] = Array.isArray(payload?.rangeRunDeltas) ? payload.rangeRunDeltas : [];

  // Backwards compatibility: older DocumentController payload shapes only include `formatDeltas`.
  // Prefer the explicit delta streams when present, but derive them from `formatDeltas` if needed.
  if (
    formatDeltas.length > 0 &&
    rowStyleDeltas.length === 0 &&
    colStyleDeltas.length === 0 &&
    sheetStyleDeltas.length === 0
  ) {
    for (const delta of formatDeltas) {
      const sheetId = typeof (delta as any)?.sheetId === "string" ? (delta as any).sheetId : "";
      if (!sheetId) continue;
      const layer = typeof (delta as any)?.layer === "string" ? (delta as any).layer : "";
      const afterStyleId = typeof (delta as any)?.afterStyleId === "number" ? (delta as any).afterStyleId : 0;
      if (!Number.isInteger(afterStyleId) || afterStyleId < 0) continue;

      if (layer === "row") {
        const row = Number((delta as any)?.index);
        if (!Number.isInteger(row) || row < 0) continue;
        rowStyleDeltas.push({ sheetId, row, afterStyleId });
      } else if (layer === "col") {
        const col = Number((delta as any)?.index);
        if (!Number.isInteger(col) || col < 0) continue;
        colStyleDeltas.push({ sheetId, col, afterStyleId });
      } else if (layer === "sheet") {
        sheetStyleDeltas.push({ sheetId, afterStyleId });
      }
    }
  }

  // If the caller supplied `getStyleById`, seed a style sync context so both
  // `engineApplyDeltas` (cell styles) and the row/col/sheet helpers can resolve
  // style objects without needing an initial `engineHydrateFromDocument` call.
  const ctx = getOrCreateStyleSyncContext(engine, options);
  const canResolveNonZeroStyles = Boolean(engine.internStyle && ctx);

  // Apply sheet metadata renames before any other deltas so `sheetIdToSheet` resolution stays
  // coherent for the remainder of the payload (cell edits, view metadata, etc).
  let didRenameAnySheets = false;
  let didApplyAnySheetDisplayNames = false;
  if (sheetMetaDeltas.length > 0) {
    for (const delta of sheetMetaDeltas) {
      if (!delta) continue;
      const sheetId = typeof delta.sheetId === "string" ? delta.sheetId.trim() : "";
      if (!sheetId) continue;
      // We only support sheet renames here (not add/delete).
      if (delta.before == null || delta.after == null) continue;

      const beforeNameRaw = typeof delta.before?.name === "string" ? delta.before.name : "";
      const afterNameRaw = typeof delta.after?.name === "string" ? delta.after.name : "";
      const oldName = beforeNameRaw.trim() || sheetId;
      const newName = afterNameRaw.trim() || sheetId;
      if (!oldName || !newName || oldName === newName) continue;

      // Prefer the stable-id display name API when available (new engine behavior).
      if (typeof engine.setSheetDisplayName === "function") {
        await engine.setSheetDisplayName(sheetId, newName);
        didApplyAnySheetDisplayNames = true;
        continue;
      }

      if (typeof engine.renameSheet !== "function") {
        throw new Error(
          `engineApplyDocumentChange: sheet rename detected (${JSON.stringify(oldName)} -> ${JSON.stringify(newName)}) but neither engine.setSheetDisplayName nor engine.renameSheet is available; rehydrate the engine`,
        );
      }

      const ok = await engine.renameSheet(oldName, newName);
      if (!ok) {
        throw new Error(
          `engineApplyDocumentChange: failed to rename sheet (${JSON.stringify(oldName)} -> ${JSON.stringify(newName)})`,
        );
      }
      didRenameAnySheets = true;
    }
  }

  const didApplyCellInputs = deltas.some((d) => cellStateToEngineInput(d.before) !== cellStateToEngineInput(d.after));
  const didApplyCellStyles =
    Boolean(engine.setCellStyleId) &&
    deltas.some((d) => d.before.styleId !== d.after.styleId && (d.after.styleId === 0 || canResolveNonZeroStyles));

  // Apply cell deltas first (value/formula + per-cell style ids), but defer recalculation so we
  // only recalc once per DocumentController change payload.
  if (deltas.length > 0) {
    await engineApplyDeltas(engine, deltas, { recalculate: false, sheetIdToSheet: options.sheetIdToSheet });
  }

  let didApplyAnyLayerStyles = false;

  if (rowStyleDeltas.length > 0 && typeof engine.setRowStyleId === "function") {
    for (const d of rowStyleDeltas) {
      const docStyleId = typeof d?.afterStyleId === "number" ? d.afterStyleId : 0;
      if (!Number.isInteger(docStyleId) || docStyleId < 0) continue;
      if (docStyleId !== 0 && !canResolveNonZeroStyles) continue;

      const engineStyleId =
        docStyleId === 0 ? null : await resolveEngineStyleIdForDocStyleId(engine, ctx!, docStyleId);
      const sheet = resolveSheet(d.sheetId);
      await engine.setRowStyleId(sheet, d.row, engineStyleId);
      didApplyAnyLayerStyles = true;
    }
  }

  if (colStyleDeltas.length > 0 && typeof engine.setColStyleId === "function") {
    for (const d of colStyleDeltas) {
      const docStyleId = typeof d?.afterStyleId === "number" ? d.afterStyleId : 0;
      if (!Number.isInteger(docStyleId) || docStyleId < 0) continue;
      if (docStyleId !== 0 && !canResolveNonZeroStyles) continue;

      const engineStyleId =
        docStyleId === 0 ? null : await resolveEngineStyleIdForDocStyleId(engine, ctx!, docStyleId);
      const sheet = resolveSheet(d.sheetId);
      await engine.setColStyleId(sheet, d.col, engineStyleId);
      didApplyAnyLayerStyles = true;
    }
  }

  if (sheetStyleDeltas.length > 0 && typeof engine.setSheetDefaultStyleId === "function") {
    for (const d of sheetStyleDeltas) {
      const docStyleId = typeof d?.afterStyleId === "number" ? d.afterStyleId : 0;
      if (!Number.isInteger(docStyleId) || docStyleId < 0) continue;
      if (docStyleId !== 0 && !canResolveNonZeroStyles) continue;

      const engineStyleId =
        docStyleId === 0 ? null : await resolveEngineStyleIdForDocStyleId(engine, ctx!, docStyleId);
      const sheet = resolveSheet(d.sheetId);
      await engine.setSheetDefaultStyleId(sheet, engineStyleId);
      didApplyAnyLayerStyles = true;
    }
  }

  // Apply range-run formatting deltas (compressed per-column formatting layer).
  //
  // These are produced by DocumentController when formatting very large rectangles so it can avoid
  // enumerating every cell. Without syncing these to the engine, Excel-compatible metadata functions
  // like `CELL("prefix")` cannot observe formatting correctly for cells that do not have explicit
  // per-cell style ids.
  let didApplyAnyRangeRuns = false;
  const setFormatRunsByCol =
    typeof engine.setFormatRunsByCol === "function"
      ? engine.setFormatRunsByCol.bind(engine)
      : typeof engine.setColFormatRuns === "function"
        ? engine.setColFormatRuns.bind(engine)
        : null;
  if (setFormatRunsByCol && rangeRunDeltas.length > 0) {
    for (const delta of rangeRunDeltas) {
      const sheetId = typeof (delta as any)?.sheetId === "string" ? (delta as any).sheetId : "";
      if (!sheetId) continue;
      const col = Number((delta as any)?.col);
      if (!Number.isInteger(col) || col < 0 || col >= 16_384) continue;

      const sheet = resolveSheet(sheetId);

      const afterRuns: unknown[] = Array.isArray((delta as any)?.afterRuns) ? (delta as any).afterRuns : [];
      if (afterRuns.length === 0) {
        // Clearing all range-run formatting for a column does not require style resolution.
        await setFormatRunsByCol(sheet, col, []);
        didApplyAnyRangeRuns = true;
        continue;
      }

      if (!canResolveNonZeroStyles) {
        // Backwards compatibility: if we can't resolve style objects, ignore non-empty run deltas.
        continue;
      }

      const runs: Array<{ startRow: number; endRowExclusive: number; styleId: number }> = [];
      for (const run of afterRuns) {
        const startRow = Number((run as any)?.startRow);
        const endRowExclusive = Number((run as any)?.endRowExclusive);
        const docStyleId = Number((run as any)?.styleId);

        if (!Number.isInteger(startRow) || startRow < 0 || startRow >= 1_048_576) continue;
        if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow || endRowExclusive > 1_048_576) continue;
        if (!Number.isInteger(docStyleId) || docStyleId < 0) continue;

        // Skip default-style runs; the engine should treat absence of a run as default formatting.
        if (docStyleId === 0) continue;

        const engineStyleId = await resolveEngineStyleIdForDocStyleId(engine, ctx!, docStyleId);
        if (!Number.isInteger(engineStyleId) || engineStyleId <= 0) continue;
        runs.push({ startRow, endRowExclusive, styleId: engineStyleId });
      }

      // Always send the full column's run list for the delta. An empty list clears the column.
      await setFormatRunsByCol(sheet, col, runs);
      didApplyAnyRangeRuns = true;
    }
  }

  // Apply sheet view metadata deltas (currently: column widths) into the engine model.
  //
  // DocumentController stores widths in base CSS px (zoom=1). The formula engine model stores
  // widths in Excel character units (OOXML `col/@width`).
  let didApplyAnyColWidths = false;
  const setColWidthChars =
    typeof engine.setColWidthChars === "function"
      ? engine.setColWidthChars.bind(engine)
      : typeof engine.setColWidth === "function"
        ? engine.setColWidth.bind(engine)
        : null;

  if (setColWidthChars && sheetViewDeltas.length > 0) {
    for (const delta of sheetViewDeltas) {
      const sheetId = typeof delta?.sheetId === "string" ? delta.sheetId : "";
      if (!sheetId) continue;
      const sheet = resolveSheet(sheetId);

      const beforeColWidths = delta?.before?.colWidths ?? null;
      const afterColWidths = delta?.after?.colWidths ?? null;
      if (!beforeColWidths && !afterColWidths) continue;

      const keys = new Set<string>();
      for (const key of Object.keys(beforeColWidths ?? {})) keys.add(key);
      for (const key of Object.keys(afterColWidths ?? {})) keys.add(key);

      for (const key of keys) {
        const col = Number(key);
        // The WASM engine is Excel-bounded (16,384 cols). DocumentController view state can
        // contain arbitrary indices (e.g. external/collab inputs), so clamp.
        if (!Number.isInteger(col) || col < 0 || col >= 16_384) continue;

        const rawBefore = beforeColWidths?.[key];
        const rawAfter = afterColWidths?.[key];
        const before = rawBefore != null && Number.isFinite(rawBefore) && rawBefore > 0 ? rawBefore : null;
        const after = rawAfter != null && Number.isFinite(rawAfter) && rawAfter > 0 ? rawAfter : null;
        if (isAxisValueEqual(before, after)) continue;

        didApplyAnyColWidths = true;
        const widthChars = after == null ? null : docColWidthPxToExcelChars(after);
        // Call via `.call(engine, ...)` to preserve `this` bindings for class-based engine targets.
        await setColWidthChars.call(engine as any, sheet, col, widthChars);
      }
    }
  }

  const recalcFlag = typeof payload?.recalc === "boolean" ? (payload.recalc as boolean) : undefined;
  let shouldRecalculate = options.recalculate ?? recalcFlag ?? true;

  // DocumentController emits `recalc: false` for many metadata-only edits (formatting, view deltas,
  // sheet meta renames). Volatile worksheet-info functions like `CELL()`/`INFO()` must observe the
  // updated metadata after such edits, so force a recalculation tick unless the caller explicitly
  // disabled it.
  //
  // Note: We only force recalculation when the metadata changes were actually applied to the engine
  // (i.e. the engine exposes the relevant optional hooks). This avoids recalculating in targets that
  // ignore formatting/view metadata entirely (backwards compatibility).
  const didApplyAnyFormattingMetadata =
    didApplyCellStyles || didApplyAnyLayerStyles || didApplyAnyRangeRuns || didApplyAnyColWidths;
  const didApplyAnyMetadataDeltas = didApplyAnyFormattingMetadata || didApplyAnySheetDisplayNames || didRenameAnySheets;

  if (didApplyAnyMetadataDeltas && options.recalculate !== false) {
    shouldRecalculate = true;
  }

  // Sheet renames can affect worksheet information functions like `CELL("filename")` and
  // `CELL("address")`, but DocumentController emits them with `recalc: false`. Override so any
  // dependent formulas observe the updated tab name.
  if (didRenameAnySheets && options.recalculate !== false) {
    shouldRecalculate = true;
  }

  // Range-run formatting can affect worksheet information functions like `CELL("format")`, but
  // DocumentController may emit formatting deltas with `recalc: false`. Override so
  // formatting-dependent formulas update when large formatted rectangles change.
  if (didApplyAnyRangeRuns && options.recalculate !== false) {
    shouldRecalculate = true;
  }
  if (!shouldRecalculate) return [];

  const didApplyAnyUpdates =
    didApplyCellInputs || didApplyCellStyles || didApplyAnyLayerStyles || didApplyAnyRangeRuns || didApplyAnyColWidths;
  if (!didApplyAnyUpdates && !didApplyAnyMetadataDeltas && recalcFlag !== true && options.recalculate !== true) return [];
  return await engine.recalculate();
}
