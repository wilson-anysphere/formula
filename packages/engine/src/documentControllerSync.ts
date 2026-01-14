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
  /**
   * Intern a formatting/protection style object into the engine and return its engine-specific id.
   *
   * Optional: engines that don't implement formatting metadata can omit this method.
   */
  internStyle?: (style: unknown) => Promise<number> | number;
  /**
   * Set the engine style id for a single cell.
   *
   * Optional: engines that don't implement formatting metadata can omit this method.
   */
  setCellStyleId?: (sheet: string, address: string, styleId: number) => Promise<void> | void;
  recalculate: (sheet?: string) => Promise<CellChange[]> | CellChange[];
  /**
   * Optional row/col/sheet formatting layer metadata APIs.
   *
   * These are additive and may not be implemented by all engine targets.
   * Sync helpers must treat them as best-effort.
   */
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
  await engine.loadWorkbookFromJson(JSON.stringify(workbookJson));

  const styleTable = doc?.styleTable as StyleTableLike | null;
  if (styleTable && typeof styleTable.get === "function") {
    const ctx: EngineStyleSyncContext = { styleTable, docStyleIdToEngineStyleId: new Map() };
    STYLE_SYNC_BY_ENGINE.set(engine, ctx);

    const sheetIds: string[] =
      typeof doc?.getSheetIds === "function" ? (doc.getSheetIds() as string[]) : [];
    const ids = sheetIds.length > 0 ? sheetIds : ["Sheet1"];

    // Sync sheet/row/col formatting layers when the engine exposes the optional hooks.
    if (engine.internStyle) {
      for (const sheetId of ids) {
        // Ensure sheet model is materialized (DocumentController is lazily sheet-creating).
        doc?.model?.getCell?.(sheetId, 0, 0);
        const sheet = doc?.model?.sheets?.get?.(sheetId);
        if (!sheet) continue;

        if (engine.setSheetDefaultStyleId) {
          const docStyleId = typeof sheet?.defaultStyleId === "number" ? sheet.defaultStyleId : 0;
          if (Number.isInteger(docStyleId) && docStyleId !== 0) {
            const engineStyleId = await resolveEngineStyleIdForDocStyleId(engine, ctx, docStyleId);
            await engine.setSheetDefaultStyleId(engineStyleId, sheetId);
          }
        }

        if (engine.setRowStyleId && sheet?.rowStyleIds?.entries) {
          for (const [row, rawDocStyleId] of sheet.rowStyleIds.entries() as Iterable<[number, number]>) {
            if (!Number.isInteger(row) || row < 0) continue;
            const docStyleId = typeof rawDocStyleId === "number" ? rawDocStyleId : 0;
            if (!Number.isInteger(docStyleId) || docStyleId === 0) continue;
            const engineStyleId = await resolveEngineStyleIdForDocStyleId(engine, ctx, docStyleId);
            await engine.setRowStyleId(row, engineStyleId, sheetId);
          }
        }

        if (engine.setColStyleId && sheet?.colStyleIds?.entries) {
          for (const [col, rawDocStyleId] of sheet.colStyleIds.entries() as Iterable<[number, number]>) {
            if (!Number.isInteger(col) || col < 0) continue;
            const docStyleId = typeof rawDocStyleId === "number" ? rawDocStyleId : 0;
            if (!Number.isInteger(docStyleId) || docStyleId === 0) continue;
            const engineStyleId = await resolveEngineStyleIdForDocStyleId(engine, ctx, docStyleId);
            await engine.setColStyleId(col, engineStyleId, sheetId);
          }
        }
      }
    }

    // Only attempt to sync per-cell style ids if the engine exposes the optional formatting hooks.
    if (engine.setCellStyleId && engine.internStyle) {
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

          const engineStyleId = await resolveEngineStyleIdForDocStyleId(engine, ctx, docStyleId);
          styleUpdates.push(engine.setCellStyleId(sheetId, address, engineStyleId));
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

  // If the caller supplied `getStyleById`, seed a style sync context so both
  // `engineApplyDeltas` (cell styles) and the row/col/sheet helpers can resolve
  // style objects without needing an initial `engineHydrateFromDocument` call.
  const ctx = getOrCreateStyleSyncContext(engine, options);
  const canResolveNonZeroStyles = Boolean(engine.internStyle && ctx);

  const didApplyCellInputs = deltas.some((d) => cellStateToEngineInput(d.before) !== cellStateToEngineInput(d.after));
  const didApplyCellStyles =
    Boolean(engine.setCellStyleId) &&
    deltas.some((d) => d.before.styleId !== d.after.styleId && (d.after.styleId === 0 || canResolveNonZeroStyles));

  // Apply cell deltas first (value/formula + per-cell style ids), but defer recalculation so we
  // only recalc once per DocumentController change payload.
  if (deltas.length > 0) {
    await engineApplyDeltas(engine, deltas, { recalculate: false });
  }

  let didApplyAnyLayerStyles = false;

  if (rowStyleDeltas.length > 0 && typeof engine.setRowStyleId === "function") {
    for (const d of rowStyleDeltas) {
      const docStyleId = typeof d?.afterStyleId === "number" ? d.afterStyleId : 0;
      if (!Number.isInteger(docStyleId) || docStyleId < 0) continue;
      if (docStyleId !== 0 && !canResolveNonZeroStyles) continue;

      const engineStyleId = docStyleId === 0 ? 0 : await resolveEngineStyleIdForDocStyleId(engine, ctx!, docStyleId);
      await engine.setRowStyleId(d.row, engineStyleId, d.sheetId);
      didApplyAnyLayerStyles = true;
    }
  }

  if (colStyleDeltas.length > 0 && typeof engine.setColStyleId === "function") {
    for (const d of colStyleDeltas) {
      const docStyleId = typeof d?.afterStyleId === "number" ? d.afterStyleId : 0;
      if (!Number.isInteger(docStyleId) || docStyleId < 0) continue;
      if (docStyleId !== 0 && !canResolveNonZeroStyles) continue;

      const engineStyleId = docStyleId === 0 ? 0 : await resolveEngineStyleIdForDocStyleId(engine, ctx!, docStyleId);
      await engine.setColStyleId(d.col, engineStyleId, d.sheetId);
      didApplyAnyLayerStyles = true;
    }
  }

  if (sheetStyleDeltas.length > 0 && typeof engine.setSheetDefaultStyleId === "function") {
    for (const d of sheetStyleDeltas) {
      const docStyleId = typeof d?.afterStyleId === "number" ? d.afterStyleId : 0;
      if (!Number.isInteger(docStyleId) || docStyleId < 0) continue;
      if (docStyleId !== 0 && !canResolveNonZeroStyles) continue;

      const engineStyleId = docStyleId === 0 ? 0 : await resolveEngineStyleIdForDocStyleId(engine, ctx!, docStyleId);
      await engine.setSheetDefaultStyleId(engineStyleId, d.sheetId);
      didApplyAnyLayerStyles = true;
    }
  }

  const recalcFlag = typeof payload?.recalc === "boolean" ? (payload.recalc as boolean) : undefined;
  const shouldRecalculate = options.recalculate ?? recalcFlag ?? true;
  if (!shouldRecalculate) return [];

  const didApplyAnyUpdates = didApplyCellInputs || didApplyCellStyles || didApplyAnyLayerStyles;
  if (!didApplyAnyUpdates && recalcFlag !== true && options.recalculate !== true) return [];
  return await engine.recalculate();
}
