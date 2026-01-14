import type { CellChange } from "./protocol.ts";
import { toA1 } from "./backend/a1.ts";
import { normalizeFormulaTextOpt } from "./backend/formula.ts";

export type EngineCellScalar = number | string | boolean | null;

export type EngineSheetJson = {
  /**
   * Sparse cell map keyed by A1 address.
   */
  cells: Record<string, EngineCellScalar>;
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

export interface EngineSyncTarget {
  loadWorkbookFromJson: (json: string) => Promise<void> | void;
  setCell: (address: string, value: EngineCellScalar, sheet?: string) => Promise<void> | void;
  setCells?: (
    updates: Array<{ address: string; value: EngineCellScalar; sheet?: string }>,
  ) => Promise<void> | void;
  recalculate: (sheet?: string) => Promise<CellChange[]> | CellChange[];
  /**
   * Optional formatting metadata interning + application methods.
   */
  internStyle?: (styleObj: unknown) => Promise<number> | number;
  setCellStyleId?: (address: string, styleId: number, sheet?: string) => Promise<void> | void;
  setRowStyleId?: (row: number, styleId: number, sheet?: string) => Promise<void> | void;
  setColStyleId?: (col: number, styleId: number, sheet?: string) => Promise<void> | void;
  setSheetDefaultStyleId?: (styleId: number, sheet?: string) => Promise<void> | void;
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
};

type StyleIdCache = Map<number, Promise<number>>;

// Per-engine cache mapping DocumentController style ids → engine style ids.
//
// We use a WeakMap so engine clients can be GC'd without explicit disposal.
const styleIdCacheByEngine = new WeakMap<object, StyleIdCache>();

function resetEngineStyleCache(engine: EngineSyncTarget): void {
  // Clear any cached document→engine style ids when the engine is re-hydrated.
  // Workbook loads typically reset internal style tables, so stale ids would be unsafe.
  styleIdCacheByEngine.delete(engine as object);
}

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
  docStyleIdToEngineStyleId: Map<number, number>;
};

// WeakMap ensures we don't leak engine instances in long-lived apps.
const STYLE_SYNC_BY_ENGINE = new WeakMap<EngineSyncTarget, EngineStyleSyncContext>();

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

/**
 * Export the current DocumentController workbook into the JSON format consumed by
 * `crates/formula-wasm` (`WasmWorkbook.fromJson`).
 *
 * Note: empty/cleared cells are omitted from the JSON entirely (sparse semantics).
 */
export function exportDocumentToEngineWorkbookJson(doc: any): EngineWorkbookJson {
  const sheets: Record<string, EngineSheetJson> = {};

  const sheetIds: string[] =
    typeof doc?.getSheetIds === "function" ? (doc.getSheetIds() as string[]) : [];
  const ids = sheetIds.length > 0 ? sheetIds : ["Sheet1"];

  for (const sheetId of ids) {
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

    sheets[sheetId] = { cells };
  }

  return { sheets };
}

export async function engineHydrateFromDocument(engine: EngineSyncTarget, doc: any): Promise<CellChange[]> {
  const workbookJson = exportDocumentToEngineWorkbookJson(doc);
  resetEngineStyleCache(engine);
  await engine.loadWorkbookFromJson(JSON.stringify(workbookJson));

  const styleTable = doc?.styleTable as StyleTableLike | null;
  if (styleTable && typeof styleTable.get === "function") {
    const ctx: EngineStyleSyncContext = { styleTable, docStyleIdToEngineStyleId: new Map() };
    STYLE_SYNC_BY_ENGINE.set(engine, ctx);

    // Only attempt to sync per-cell style ids if the engine exposes the optional formatting hooks.
    if (engine.setCellStyleId && engine.internStyle) {
      const sheetIds: string[] =
        typeof doc?.getSheetIds === "function" ? (doc.getSheetIds() as string[]) : [];
      const ids = sheetIds.length > 0 ? sheetIds : ["Sheet1"];

      for (const sheetId of ids) {
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

          let engineStyleId = ctx.docStyleIdToEngineStyleId.get(docStyleId);
          if (engineStyleId == null) {
            const style = styleTable.get(docStyleId);
            engineStyleId = await engine.internStyle(style);
            ctx.docStyleIdToEngineStyleId.set(docStyleId, engineStyleId);
          }
          styleUpdates.push(engine.setCellStyleId(address, engineStyleId, sheetId));
        }
        // Apply per-sheet to keep memory bounded for large workbooks.
        if (styleUpdates.length > 0) await Promise.all(styleUpdates);
      }
    }
  }

  return await engine.recalculate();
}

export async function engineApplyDeltas(
  engine: EngineSyncTarget,
  deltas: readonly DocumentCellDelta[],
  options: { recalculate?: boolean } = {},
): Promise<CellChange[]> {
  const shouldRecalculate = options.recalculate ?? true;
  const updates: Array<{ address: string; value: EngineCellScalar; sheet?: string }> = [];
  const styleUpdates: Array<{ address: string; docStyleId: number; sheet?: string }> = [];
  let didApply = false;

  for (const delta of deltas) {
    const beforeInput = cellStateToEngineInput(delta.before);
    const afterInput = cellStateToEngineInput(delta.after);

    const address = toA1(delta.row, delta.col);

    // Track value/formula edits (including rich-text run edits that change plain text).
    if (beforeInput !== afterInput) {
      updates.push({ address, value: afterInput, sheet: delta.sheetId });
    }

    // Track formatting-only edits when the engine supports style metadata.
    if (engine.setCellStyleId && delta.before.styleId !== delta.after.styleId) {
      styleUpdates.push({ address, docStyleId: delta.after.styleId, sheet: delta.sheetId });
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
    const styleTable = ctx?.styleTable;
    const cache = ctx?.docStyleIdToEngineStyleId;

    const setStylePromises: Array<Promise<void> | void> = [];
    for (const update of styleUpdates) {
      // Clearing back to the default style does not require interning.
      if (update.docStyleId === 0) {
        setStylePromises.push(engine.setCellStyleId(update.address, 0, update.sheet));
        continue;
      }

      if (!engine.internStyle || !styleTable || !cache) {
        // Backwards compatibility: if we can't resolve the style object, ignore formatting-only deltas.
        continue;
      }

      let engineStyleId = cache.get(update.docStyleId);
      if (engineStyleId == null) {
        const style = styleTable.get(update.docStyleId);
        engineStyleId = await engine.internStyle(style);
        cache.set(update.docStyleId, engineStyleId);
      }
      setStylePromises.push(engine.setCellStyleId(update.address, engineStyleId, update.sheet));
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

  // Handle "reset boundary" payloads by forcing callers to re-hydrate. We intentionally
  // don't attempt to mirror sheet structure changes incrementally here.
  if (payload?.source === "applyState") {
    return [];
  }

  const deltas: readonly DocumentCellDelta[] = Array.isArray(payload?.deltas) ? payload.deltas : [];
  const rowStyleDeltas: Array<{ sheetId: string; row: number; afterStyleId: number }> = Array.isArray(payload?.rowStyleDeltas)
    ? payload.rowStyleDeltas
    : [];
  const colStyleDeltas: Array<{ sheetId: string; col: number; afterStyleId: number }> = Array.isArray(payload?.colStyleDeltas)
    ? payload.colStyleDeltas
    : [];
  const sheetStyleDeltas: Array<{ sheetId: string; afterStyleId: number }> = Array.isArray(payload?.sheetStyleDeltas)
    ? payload.sheetStyleDeltas
    : [];

  const hasRowStyleDeltas = rowStyleDeltas.length > 0;
  const hasColStyleDeltas = colStyleDeltas.length > 0;
  const hasSheetStyleDeltas = sheetStyleDeltas.length > 0;

  // Apply cell content deltas (same logic as `engineApplyDeltas`, but we delay recalc until
  // after formatting metadata is applied so we only recalc once per DocumentController event).
  const updates: Array<{ address: string; value: EngineCellScalar; sheet?: string }> = [];

  for (const delta of deltas) {
    const beforeInput = cellStateToEngineInput(delta.before);
    const afterInput = cellStateToEngineInput(delta.after);

    // Ignore formatting-only edits and rich-text run edits that don't change the plain input.
    if (beforeInput === afterInput) continue;

    const address = toA1(delta.row, delta.col);
    updates.push({ address, value: afterInput, sheet: delta.sheetId });
  }

  const didApplyCellInputs = updates.length > 0;
  if (didApplyCellInputs) {
    if (engine.setCells) {
      await engine.setCells(updates);
    } else {
      await Promise.all(updates.map((u) => engine.setCell(u.address, u.value, u.sheet)));
    }
  }

  // Apply row/col/sheet style deltas (best-effort).
  const getStyleById = options.getStyleById;
  const internStyle = engine.internStyle;

  const didAttemptStyleSync =
    Boolean(getStyleById) && typeof internStyle === "function" && (hasRowStyleDeltas || hasColStyleDeltas || hasSheetStyleDeltas);

  let didApplyAnyStyles = false;

  if (didAttemptStyleSync) {
    let cache = styleIdCacheByEngine.get(engine as object);
    if (!cache) {
      cache = new Map();
      // Seed the default style: document `0` always maps to engine `0`.
      cache.set(0, Promise.resolve(0));
      styleIdCacheByEngine.set(engine as object, cache);
    }

    const resolveEngineStyleId = async (docStyleId: unknown): Promise<number> => {
      const id = typeof docStyleId === "number" ? docStyleId : 0;
      if (!Number.isInteger(id) || id <= 0) return 0;

      const existing = cache.get(id);
      if (existing) return await existing;

      const styleObj = getStyleById!(id);
      const promise = Promise.resolve(internStyle!(styleObj));
      cache.set(id, promise);
      return await promise;
    };

    if (hasRowStyleDeltas && typeof engine.setRowStyleId === "function") {
      for (const d of rowStyleDeltas) {
        const styleId = await resolveEngineStyleId(d.afterStyleId);
        await engine.setRowStyleId(d.row, styleId, d.sheetId);
        didApplyAnyStyles = true;
      }
    }

    if (hasColStyleDeltas && typeof engine.setColStyleId === "function") {
      for (const d of colStyleDeltas) {
        const styleId = await resolveEngineStyleId(d.afterStyleId);
        await engine.setColStyleId(d.col, styleId, d.sheetId);
        didApplyAnyStyles = true;
      }
    }

    if (hasSheetStyleDeltas && typeof engine.setSheetDefaultStyleId === "function") {
      for (const d of sheetStyleDeltas) {
        const styleId = await resolveEngineStyleId(d.afterStyleId);
        await engine.setSheetDefaultStyleId(styleId, d.sheetId);
        didApplyAnyStyles = true;
      }
    }
  }

  const didApplyAnyUpdates = didApplyCellInputs || didApplyAnyStyles;

  const recalcFlag = typeof payload?.recalc === "boolean" ? (payload.recalc as boolean) : undefined;
  const shouldRecalculate =
    typeof options.recalculate === "boolean" ? options.recalculate : didApplyAnyUpdates ? recalcFlag !== false : recalcFlag === true;

  if (!shouldRecalculate) return [];
  return await engine.recalculate();
}
