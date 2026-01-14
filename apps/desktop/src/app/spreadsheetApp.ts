import { CellEditorOverlay } from "../editor/cellEditorOverlay";
import { FormulaBarTabCompletionController } from "../ai/completion/formulaBarTabCompletion.js";
import { FormulaBarView } from "../formula-bar/FormulaBarView";
import type { RangeAddress as A1RangeAddress } from "../spreadsheet/a1.js";
import { Outline, groupDetailRange, isHidden } from "../grid/outline/outline.js";
import { parseA1Range } from "../charts/a1.js";
import { emuToPx } from "../charts/overlay.js";
import { chartAnchorToDrawingAnchor } from "../charts/chartAnchorToDrawingAnchor";
import { ChartCanvasStoreAdapter } from "../charts/chartCanvasStoreAdapter";
import { chartIdToDrawingId as chartStoreIdToDrawingId, chartRecordToDrawingObject, drawingAnchorToChartAnchor } from "../charts/chartDrawingAdapter";
import { ChartStore, type ChartRecord } from "../charts/chartStore";
import { ChartRendererAdapter, type ChartStore as ChartRendererStore } from "../charts/chartRendererAdapter";
import type { ChartModel } from "../charts/renderChart";
import { FormulaChartModelStore } from "../charts/formulaChartModelStore";
import { FALLBACK_CHART_THEME, type ChartTheme } from "../charts/theme";
import { buildHitTestIndex, drawingObjectToViewportRect, hitTestDrawings, type HitTestIndex } from "../drawings/hitTest";
import {
  DrawingInteractionController,
  resizeAnchor,
  shiftAnchor,
  type DrawingInteractionCallbacks,
} from "../drawings/interaction.js";
import {
  cursorForRotationHandle,
  cursorForResizeHandle,
  cursorForResizeHandleWithTransform,
  getResizeHandleCenters,
  hitTestRotationHandle,
  hitTestResizeHandle,
  RESIZE_HANDLE_SIZE_PX,
  type ResizeHandle,
} from "../drawings/selectionHandles";
import {
  DrawingOverlay,
  anchorToRectPx,
  effectiveScrollForAnchor,
  pxToEmu,
  type GridGeometry as DrawingGridGeometry,
  type Viewport as DrawingViewport,
} from "../drawings/overlay";
import { createDrawingObjectId, type Anchor as DrawingAnchor, type DrawingObject, type ImageEntry, type ImageStore } from "../drawings/types";
import { convertDocumentSheetDrawingsToUiDrawingObjects, convertModelWorksheetDrawingsToUiDrawingObjects } from "../drawings/modelAdapters";
import { duplicateSelected as duplicateDrawingSelected } from "../drawings/commands";
import { decodeBase64ToBytes as decodeClipboardImageBase64ToBytes, insertImageFromFile } from "../drawings/insertImage";
import { pickLocalImageFiles } from "../drawings/pickLocalImageFiles.js";
import { MAX_INSERT_IMAGE_BYTES } from "../drawings/insertImageLimits.js";
import { IndexedDbImageStore } from "../drawings/persistence/indexedDbImageStore";
import { applyPlainTextEdit } from "../grid/text/rich-text/edit.js";
import { renderRichText } from "../grid/text/rich-text/render.js";
import {
  createClipboardProvider,
  clipboardFormatToDocStyle,
  CLIPBOARD_LIMITS,
  getCellGridFromRange,
  parseClipboardContentToCellGrid,
  serializeCellGridToClipboardPayload,
} from "../clipboard/index.js";
import { reconcileClipboardCopyContextForPaste } from "./clipboardPasteContext";
import { cellToA1, rangeToA1 } from "../selection/a1";
import { computeCurrentRegionRange } from "../selection/currentRegion";
import { navigateSelectionByKey } from "../selection/navigation";
import { cellInRange } from "../selection/range";
import { SelectionRenderer } from "../selection/renderer";
import type { CellCoord, GridLimits, Range, SelectionState } from "../selection/types";
import { AuditingOverlayRenderer } from "../grid/auditing-overlays/AuditingOverlayRenderer";
import { computeAuditingOverlays, type AuditingMode, type AuditingOverlays } from "../grid/auditing-overlays/overlays";
import { resolveCssVar } from "../theme/cssVars.js";
import { t, tWithVars } from "../i18n/index.js";
import {
  DEFAULT_GRID_LIMITS,
  addCellToSelection,
  createSelection,
  buildSelection,
  extendSelectionToCell,
  selectAll,
  selectColumns,
  selectRows,
  setActiveCell
} from "../selection/selection";
import { DEFAULT_DESKTOP_LOAD_MAX_COLS, DEFAULT_DESKTOP_LOAD_MAX_ROWS } from "../workbook/load/clampUsedRange.js";
import { DocumentController } from "../document/documentController.js";
import { MockEngine } from "../document/engine.js";
import { isRedoKeyboardEvent, isUndoKeyboardEvent } from "../document/shortcuts.js";
import { showToast, showQuickPick } from "../extensions/ui.js";
import { applyNumberFormatPreset, toggleBold, toggleItalic, toggleStrikethrough, toggleUnderline } from "../formatting/toolbar.js";
import {
  DEFAULT_FORMATTING_APPLY_CELL_LIMIT,
  evaluateFormattingSelectionSize,
  normalizeSelectionRange,
} from "../formatting/selectionSizeGuard.js";
import { formatValueWithNumberFormat } from "../formatting/numberFormat.ts";
import { dateToExcelSerial } from "../shared/valueParsing.js";
import { createDesktopDlpContext } from "../dlp/desktopDlp.js";
import { enforceClipboardCopy } from "../dlp/enforceClipboardCopy.js";
import { DlpViolationError } from "../../../../packages/security/dlp/src/errors.js";
import {
  createEngineClient,
  engineApplyDocumentChange,
  engineHydrateFromDocument,
  type EditOp,
  type EditResult,
  type EngineClient
} from "@formula/engine";
import { createUndoService, type UndoService } from "@formula/collab-undo";
import { drawCommentIndicator } from "../comments/CommentIndicator";
import { evaluateFormula, type SpreadsheetValue } from "../spreadsheet/evaluateFormula";
import { AiCellFunctionEngine } from "../spreadsheet/AiCellFunctionEngine.js";
import { DocumentWorkbookAdapter } from "../search/documentWorkbookAdapter.js";
import type { SheetNameResolver } from "../sheet/sheetNameResolver";
import { formatSheetNameForA1 } from "../sheet/formatSheetNameForA1.js";
import { parseGoTo, splitSheetQualifier } from "../../../../packages/search/index.js";
import type { CreateChartResult, CreateChartSpec } from "../../../../packages/ai-tools/src/spreadsheet/api.js";
import { colToName as colToNameA1, fromA1 as fromA1A1 } from "@formula/spreadsheet-frontend/a1";
import { extractFormulaReferences, shiftA1References, toggleA1AbsoluteAtCursor } from "@formula/spreadsheet-frontend";
import { createSchemaProviderFromSearchWorkbook } from "../ai/context/searchWorkbookSchemaProvider.js";
import type { WorkbookContextBuildStats } from "../ai/context/WorkbookContextBuilder.js";
import { InlineEditController, type InlineEditLLMClient } from "../ai/inline-edit/inlineEditController";
import type { AIAuditStore } from "../../../../packages/ai-audit/src/store.js";
import { DEFAULT_GRID_FONT_FAMILY, DEFAULT_GRID_MONOSPACE_FONT_FAMILY, clampZoom } from "@formula/grid";
import type {
  CanvasGridImageResolver,
  CellRange as GridCellRange,
  GridAxisSizeChange,
  GridPresence,
  GridViewportState
} from "@formula/grid";
import { wheelDeltaToPixels } from "@formula/grid";
import { resolveDesktopGridMode, type DesktopGridMode } from "../grid/shared/desktopGridMode.js";
import { DocumentCellProvider } from "../grid/shared/documentCellProvider.js";
import { DesktopSharedGrid, type DesktopSharedGridCallbacks } from "../grid/shared/desktopSharedGrid.js";
import { DesktopImageStore } from "../images/imageStore.js";
import { openExternalHyperlink } from "../hyperlinks/openExternal.js";
import { getTauriDialogOpenOrNull } from "../tauri/api";
import * as nativeDialogs from "../tauri/nativeDialogs.js";
import { shellOpen } from "../tauri/shellOpen.js";
import {
  applyFillCommitToDocumentController,
  applyFillCommitToDocumentControllerWithFormulaRewrite,
  computeFillEditsForDocumentControllerWithFormulaRewrite,
} from "../fill/applyFillCommit";
import type { CellRange as FillEngineRange, FillMode as FillHandleMode } from "@formula/fill-engine";
import { bindSheetViewToCollabSession, type SheetViewBinder } from "../collab/sheetViewBinder";
import { bindImageBytesToCollabSession, type ImageBytesBinder } from "../collab/imageBytesBinder";
import { resolveDevCollabEncryptionFromSearch } from "../collab/devEncryption.js";
import { CollabEncryptionKeyStore } from "../collab/encryptionKeyStore";
import {
  createEncryptedRangeManagerForSession,
  createEncryptionPolicyFromDoc,
  type EncryptedRangeManager
} from "@formula/collab-encrypted-ranges";
import { loadCollabConnectionForWorkbook, saveCollabConnectionForWorkbook } from "../sharing/collabConnectionStore.js";
import { reservedRootGuardUiMessage, subscribeToReservedRootGuardDisconnect } from "../panels/collabReservedRootGuard.js";
import { loadCollabToken, storeCollabToken } from "../sharing/collabTokenStore.js";
import { showCollabEditRejectedToast } from "../collab/editRejectionToast";
import { ImageBitmapCache } from "../drawings/imageBitmapCache";
import { applyTransformVector, inverseTransformVector } from "../drawings/transform";

import * as Y from "yjs";
import { CommentManager, bindDocToStorage, createCommentManagerForDoc, getCommentsRoot } from "@formula/collab-comments";
import type { Comment, CommentAuthor } from "@formula/collab-comments";
import { bindCollabSessionToDocumentController, createCollabSession, makeCellKey, type CollabSession } from "@formula/collab-session";
import { IndexedDbCollabPersistence } from "@formula/collab-persistence/indexeddb";
import { tryDeriveCollabSessionPermissionsFromJwtToken } from "../collab/jwt";
import { getCollabUserIdentity, overrideCollabUserIdentityId, type CollabUserIdentity } from "../collab/userIdentity";

import { PresenceRenderer } from "../grid/presence-renderer/presenceRenderer.js";
import { ConflictUiController } from "../collab/conflicts-ui/conflict-ui-controller.js";
import { StructuralConflictUiController } from "../collab/conflicts-ui/structural-conflict-ui-controller.js";
import {
  CollaboratorsListUiController,
  type CollaboratorListEntry,
} from "../collab/presence-ui/collaborators-list-ui-controller.js";
type FormulaBarCommit = Parameters<ConstructorParameters<typeof FormulaBarView>[1]["onCommit"]>[1];

type NameBoxDropdownProvider = NonNullable<
  NonNullable<ConstructorParameters<typeof FormulaBarView>[2]>["nameBoxDropdownProvider"]
>;

type EngineCellRef = { sheetId?: string; sheet?: string; row?: number; col?: number; address?: string; value?: unknown };
type AuditingCacheEntry = {
  precedents: string[];
  dependents: string[];
  precedentsError: string | null;
  dependentsError: string | null;
};
const MAX_KEYBOARD_FORMATTING_CELLS = DEFAULT_FORMATTING_APPLY_CELL_LIMIT;
// Copying a large rectangle requires allocating a per-cell clipboard payload (TSV/HTML/RTF)
// and (for internal pastes) a per-cell snapshot of effective formats. Keep this bounded so
// Excel-scale sheet limits don't allow accidental multi-million-cell allocations.
const MAX_CLIPBOARD_CELLS = 200_000;
// Fill-handle + fill shortcut operations materialize per-cell edits (and, in some cases, a
// source snapshot) in JS. Keep this bounded so users can't accidentally generate millions
// of edits on Excel-scale sheets.
const MAX_FILL_CELLS = 200_000;
// Excel-style date/time insertion shortcuts (Ctrl+; / Ctrl+Shift+;) can target the full
// selection. Cap enumeration so accidental large selections don't allocate huge 2D arrays.
const MAX_DATE_TIME_INSERT_CELLS = 10_000;
// Chart rendering is synchronous and (today) materializes the full series ranges into JS arrays.
// Keep charts bounded so a large A1 range doesn't allocate millions of values on every render.
const MAX_CHART_DATA_CELLS = 100_000;
// Formula-bar range preview tooltip should never enumerate massive ranges.
// Keep reads bounded to avoid keystroke-latency regressions when editing formulas.
const MAX_FORMULA_RANGE_PREVIEW_CELLS = 100;
const FORMULA_RANGE_PREVIEW_SAMPLE_ROWS = 3;
const FORMULA_RANGE_PREVIEW_SAMPLE_COLS = 3;
let formulaRangePreviewTooltipIdCounter = 0;
function nextFormulaRangePreviewTooltipId(): string {
  formulaRangePreviewTooltipIdCounter += 1;
  return `formula-range-preview-tooltip-${formulaRangePreviewTooltipIdCounter}`;
}
// Encode (row, col) into a single numeric key for allocation-free lookups.
// `16_384` matches Excel's maximum column count, so the mapping is collision-free for Excel-sized sheets.
const COMMENT_COORD_COL_STRIDE = 16_384;
const COMPUTED_COORD_COL_STRIDE = COMMENT_COORD_COL_STRIDE;
// Encode (row, col) into a single numeric key for per-call formula memoization / cycle detection.
// Use a large stride so even when the UI is configured with a small `maxCols` (e.g. tests),
// formulas referencing larger column indexes don't collide in memoization keys.
const EVAL_COORD_COL_STRIDE = 1_048_576; // 2^20 (also Excel's max row count)
// Plain A1 address (with optional `$` absolute markers) without sheet qualification.
const A1_CELL_REF_RE = /^\$?[A-Za-z]+\$?[1-9][0-9]*$/;
const AI_FUNCTION_CALL_RE = /\bAI(?:\.(?:EXTRACT|CLASSIFY|TRANSLATE))?\s*\(/i;

function isThenable(value: unknown): value is PromiseLike<unknown> {
  return typeof (value as { then?: unknown } | null)?.then === "function";
}

function isInteger(value: unknown): value is number {
  return typeof value === "number" && Number.isInteger(value);
}

async function mapWithConcurrencyLimit<T, R>(
  items: readonly T[],
  limit: number,
  fn: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const count = items.length;
  if (count === 0) return [];
  const max = Math.max(1, Math.trunc(limit));
  const results = new Array<R>(count);
  let nextIndex = 0;

  const worker = async () => {
    while (true) {
      const idx = nextIndex;
      nextIndex += 1;
      if (idx >= count) return;
      results[idx] = await fn(items[idx], idx);
    }
  };

  await Promise.all(Array.from({ length: Math.min(max, count) }, () => worker()));
  return results;
}

function looksLikeExternalHyperlink(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed) return false;
  // Avoid treating arbitrary "foo:bar" values as URLs; require either a scheme
  // separator (`://`) or a `mailto:` prefix.
  if (/^mailto:/i.test(trimmed)) return true;
  return /^[a-zA-Z][a-zA-Z0-9+.-]*:\/\//.test(trimmed);
}

function decodeBase64ToBytes(base64: string): Uint8Array {
  if (typeof atob !== "function") return new Uint8Array();
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function readPngDimensions(bytes: Uint8Array): { width: number; height: number } | null {
  // PNG signature (8 bytes) + IHDR chunk header (8 bytes) + width/height (8 bytes).
  if (!(bytes instanceof Uint8Array) || bytes.byteLength < 24) return null;

  if (
    bytes[0] !== 0x89 ||
    bytes[1] !== 0x50 ||
    bytes[2] !== 0x4e ||
    bytes[3] !== 0x47 ||
    bytes[4] !== 0x0d ||
    bytes[5] !== 0x0a ||
    bytes[6] !== 0x1a ||
    bytes[7] !== 0x0a
  ) {
    return null;
  }

  // The first chunk should be IHDR: length (4) + type (4) + data...
  if (bytes[12] !== 0x49 || bytes[13] !== 0x48 || bytes[14] !== 0x44 || bytes[15] !== 0x52) {
    return null;
  }

  try {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const width = view.getUint32(16, false);
    const height = view.getUint32(20, false);
    if (width === 0 || height === 0) return null;
    return { width, height };
  } catch {
    return null;
  }
}
function inferMimeTypeFromId(id: string, bytes?: Uint8Array): string {
  const ext = String(id ?? "").split(".").pop()?.toLowerCase();
  switch (ext) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "webp":
      return "image/webp";
    case "svg":
      return "image/svg+xml";
    default:
      break;
  }

  if (bytes && bytes.length >= 4) {
    // PNG
    if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return "image/png";
    // JPEG
    if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return "image/jpeg";
    // GIF
    if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x38) return "image/gif";
    // BMP
    if (bytes[0] === 0x42 && bytes[1] === 0x4d) return "image/bmp";
    // WebP: "RIFF"...."WEBP"
    if (
      bytes.length >= 12 &&
      bytes[0] === 0x52 &&
      bytes[1] === 0x49 &&
      bytes[2] === 0x46 &&
      bytes[3] === 0x46 &&
      bytes[8] === 0x57 &&
      bytes[9] === 0x45 &&
      bytes[10] === 0x42 &&
      bytes[11] === 0x50
    ) {
      return "image/webp";
    }
  }

  return "application/octet-stream";
}

function normalizeImageEntry(id: string, raw: unknown): ImageEntry | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const record = raw as any;
  const bytes: unknown = record.bytes;
  if (!(bytes instanceof Uint8Array)) return undefined;
  const entryId = typeof record.id === "string" && record.id.trim() !== "" ? record.id : id;
  const mimeTypeRaw: unknown = record.mimeType ?? record.contentType ?? record.content_type;
  const mimeType =
    typeof mimeTypeRaw === "string" && mimeTypeRaw.trim() !== "" ? mimeTypeRaw.trim() : inferMimeTypeFromId(entryId, bytes);
  return { id: entryId, bytes, mimeType };
}

function lookupImageEntry(id: string, images: unknown): ImageEntry | undefined {
  if (!images) return undefined;

  if (images instanceof Map) {
    return normalizeImageEntry(id, (images as Map<string, unknown>).get(id));
  }

  if (typeof images === "object") {
    return normalizeImageEntry(id, (images as Record<string, unknown>)[id]);
  }

  return undefined;
}

/**
 * ImageStore adapter that reads workbook images from DocumentController (when available).
 *
 * This avoids per-frame copying by returning the stored Uint8Array bytes directly and relies
 * on DrawingOverlay's ImageBitmapCache to dedupe decoding work.
 */
class DocumentImageStore implements ImageStore {
  private readonly fallback = new Map<string, ImageEntry>();

  constructor(
    private readonly document: DocumentController,
    private readonly persisted: IndexedDbImageStore,
    private readonly options: { mode?: "user" | "external"; source?: string } = {},
  ) {}

  /**
   * Best-effort teardown for tests/hot-reload.
   *
   * Clears in-memory image bytes caches so a disposed SpreadsheetApp does not
   * retain large image payloads via the ImageStore even if the app object is
   * still referenced somewhere.
   */
  dispose(): void {
    this.fallback.clear();
    try {
      this.persisted.clearMemory();
    } catch {
      // ignore
    }
  }

  get(id: string): ImageEntry | undefined {
    const imageId = String(id ?? "");
    if (!imageId) return undefined;

    const doc = this.document as any;

    // Prefer the internal map to avoid cloning bytes on every read.
    const entry = lookupImageEntry(imageId, doc.images);
    if (entry) return entry;

    // Fallback: direct getter (may clone bytes).
    if (typeof doc.getImage === "function") {
      try {
        const direct = normalizeImageEntry(imageId, doc.getImage(imageId));
        if (direct) return direct;
      } catch {
        // ignore
      }
    }

    return this.fallback.get(imageId);
  }

  set(entry: ImageEntry): void {
    if (!entry || typeof entry.id !== "string") return;

    // Best-effort persistence: never block callers on IndexedDB availability.
    this.persisted.set(entry);

    // External sync path (collab hydration): apply without creating undo history.
    const doc = this.document as any;
    if (this.options.mode === "external" && typeof doc.applyExternalImageDeltas === "function") {
      try {
        const imageId = entry.id;
        const existing = doc.images?.get?.(imageId) ?? null;
        const before =
          existing && existing.bytes instanceof Uint8Array
            ? // Preserve whether the stored entry explicitly had a mimeType field.
              ("mimeType" in existing ? { bytes: existing.bytes, mimeType: existing.mimeType ?? null } : { bytes: existing.bytes })
            : null;
        doc.applyExternalImageDeltas(
          [
            {
              imageId,
              before,
              after: { bytes: entry.bytes, mimeType: entry.mimeType },
            },
          ],
          { source: this.options.source ?? "collab" },
        );
      } catch {
        // ignore
      }
    }

    // Keep a session-local cache so synchronous `get()` calls can resolve quickly without
    // reading from IndexedDB.
    this.fallback.set(entry.id, entry);
  }

  async getAsync(id: string): Promise<ImageEntry | undefined> {
    const imageId = String(id ?? "");
    if (!imageId) return undefined;

    // Fast-path: already present in memory (DocumentController or fallback).
    const existing = this.get(imageId);
    if (existing) return existing;

    const loaded = await this.persisted.getAsync(imageId);
    if (!loaded) return undefined;

    // Cache in-memory so subsequent sync `get()` calls can resolve quickly without
    // awaiting IndexedDB again. We intentionally do *not* mutate the DocumentController
    // image store here (that would be undoable / affect snapshots).
    this.fallback.set(imageId, loaded);
    return loaded;
  }

  async setAsync(entry: ImageEntry): Promise<void> {
    await this.persisted.setAsync(entry);
  }

  async garbageCollectAsync(keep: Iterable<string>): Promise<void> {
    const keepSet = new Set(Array.from(keep, (id) => String(id)));
    // Remove unused records from the persistent store.
    await this.persisted.garbageCollectAsync(keepSet);
    // Also drop any now-unreferenced in-memory cached entries.
    for (const id of Array.from(this.fallback.keys())) {
      if (!keepSet.has(id)) this.fallback.delete(id);
    }
  }
}

/**
 * Tracks async engine work so tests can deterministically await "engine idle".
 *
 * This intentionally *does not* assume a specific engine implementation; it just
 * aggregates promises representing:
 *  - batched input syncing (e.g. worker `setCells`)
 *  - recalculation
 *  - applying computed CellChange[] into the app's computed-value cache
 */
class IdleTracker {
  private readonly pending = new Set<Promise<unknown>>();
  private waiters: Array<() => void> = [];
  private error: unknown = null;
  private resolveScheduled = false;

  track(value: PromiseLike<unknown>): void {
    const promise = Promise.resolve(value);
    // Prevent unhandled rejection noise if no one is currently awaiting `whenIdle`.
    promise.catch((err) => {
      if (this.error == null) this.error = err;
    });
    this.pending.add(promise);
    promise.finally(() => {
      this.pending.delete(promise);
      this.maybeResolve();
    });
  }

  whenIdle(): Promise<void> {
    if (this.pending.size === 0) {
      return this.error == null ? Promise.resolve() : Promise.reject(this.error);
    }

    return new Promise((resolve, reject) => {
      this.waiters.push(() => {
        if (this.error == null) resolve();
        else reject(this.error);
      });
    });
  }

  private maybeResolve(): void {
    if (this.pending.size !== 0) return;
    if (this.resolveScheduled) return;
    this.resolveScheduled = true;

    // Wait one microtask turn before resolving to allow follow-up work to be
    // tracked (e.g. `.then(() => track(...))` attached after the original
    // promise was tracked).
    queueMicrotask(() => {
      this.resolveScheduled = false;
      if (this.pending.size !== 0) return;
      const waiters = this.waiters;
      this.waiters = [];
      for (const resolve of waiters) resolve();
    });
  }
}

type EngineIntegration = {
  applyChanges?(changes: unknown): unknown;
  recalculate?(): unknown;
  syncNow?(): unknown;
  beginBatch?(): unknown;
  endBatch?(): unknown;
};

class IdleTrackingEngine {
  recalcCount = 0;

  constructor(
    private readonly inner: EngineIntegration,
    private readonly idle: IdleTracker,
    private readonly onInputsChanged: (changes: unknown) => void,
    private readonly onComputedChanges: (changes: unknown) => void | Promise<void>
  ) {}

  applyChanges(changes: unknown): void {
    this.onInputsChanged(changes);
    const result = this.inner.applyChanges?.(changes);
    if (isThenable(result)) {
      this.idle.track(result);
    }
  }

  recalculate(): void {
    this.recalcCount += 1;
    const result = this.inner.recalculate?.();

    if (isThenable(result)) {
      const tracked = Promise.resolve(result).then((changes) => this.onComputedChanges(changes));
      this.idle.track(tracked);
      return;
    }

    // Some engines may return the changes synchronously.
    if (result !== undefined) {
      const applied = this.onComputedChanges(result);
      if (isThenable(applied)) this.idle.track(applied);
    }
  }

  syncNow(): void {
    const result = this.inner.syncNow?.();
    if (isThenable(result)) {
      const tracked = Promise.resolve(result).then((changes) => this.onComputedChanges(changes));
      this.idle.track(tracked);
      return;
    }

    if (result !== undefined) {
      const applied = this.onComputedChanges(result);
      if (isThenable(applied)) this.idle.track(applied);
    }
  }

  beginBatch(): void {
    const result = this.inner.beginBatch?.();
    if (isThenable(result)) this.idle.track(result);
  }

  endBatch(): void {
    const result = this.inner.endBatch?.();
    if (isThenable(result)) this.idle.track(result);
  }
}

type ChartSeriesDef = {
  name?: string | null;
  categories?: string | null;
  values?: string | null;
  xValues?: string | null;
  yValues?: string | null;
};

type ChartDef = {
  chartType: { kind: string; name?: string };
  title?: string;
  series: ChartSeriesDef[];
  anchor: {
    kind: string;
    fromCol?: number;
    fromRow?: number;
    fromColOffEmu?: number;
    fromRowOffEmu?: number;
    toCol?: number;
    toRow?: number;
    toColOffEmu?: number;
    toRowOffEmu?: number;
    xEmu?: number;
    yEmu?: number;
    cxEmu?: number;
    cyEmu?: number;
  };
};

type DragState =
  | { pointerId: number; mode: "normal" }
  | { pointerId: number; mode: "formula" }
  | {
      pointerId: number;
      mode: "fill";
      sourceRange: Range;
      targetRange: Range;
      endCell: CellCoord;
      fillMode: FillHandleMode;
      activeRangeIndex: number;
    };

type DrawingGestureState =
  | {
      pointerId: number;
      mode: "drag";
      objectId: number;
      startSheetX: number;
      startSheetY: number;
      startAnchor: DrawingAnchor;
    }
  | {
      pointerId: number;
      mode: "resize";
      objectId: number;
      handle: ResizeHandle;
      startSheetX: number;
      startSheetY: number;
      startAnchor: DrawingAnchor;
      startWidthPx: number;
      startHeightPx: number;
      transform?: DrawingObject["transform"];
      /** Only set for image objects; used when Shift is held during resize. */
      aspectRatio: number | null;
    };

function lockAspectRatioResize(args: {
  handle: ResizeHandle;
  dx: number;
  dy: number;
  startWidthPx: number;
  startHeightPx: number;
  aspectRatio: number;
  minSizePx: number;
}): { dx: number; dy: number } {
  const { handle, startWidthPx, startHeightPx } = args;
  let { dx, dy, aspectRatio } = args;

  // Only lock corner-handle resizes (edge handles remain unconstrained).
  if (handle === "n" || handle === "e" || handle === "s" || handle === "w") return { dx, dy };
  if (!Number.isFinite(aspectRatio) || aspectRatio <= 0) return { dx, dy };
  if (!Number.isFinite(startWidthPx) || !Number.isFinite(startHeightPx)) return { dx, dy };
  if (startWidthPx <= 0 || startHeightPx <= 0) return { dx, dy };

  const sx = handle === "ne" || handle === "se" ? 1 : -1;
  const sy = handle === "sw" || handle === "se" ? 1 : -1;

  const proposedWidth = startWidthPx + sx * dx;
  const proposedHeight = startHeightPx + sy * dy;
  const scaleW = proposedWidth / startWidthPx;
  const scaleH = proposedHeight / startHeightPx;

  const widthDriven = Math.abs(scaleW - 1) >= Math.abs(scaleH - 1);

  const minScale = Math.max(
    startWidthPx > args.minSizePx ? args.minSizePx / startWidthPx : 0,
    startHeightPx > args.minSizePx ? args.minSizePx / startHeightPx : 0,
  );

  const clampScale = (s: number): number => {
    if (!Number.isFinite(s)) return 1;
    return Math.max(s, minScale, 0);
  };

  if (widthDriven) {
    const scale = clampScale(scaleW);
    const nextWidth = startWidthPx * scale;
    const nextHeight = nextWidth / aspectRatio;
    return { dx: (nextWidth - startWidthPx) * sx, dy: (nextHeight - startHeightPx) * sy };
  }

  const scale = clampScale(scaleH);
  const nextHeight = startHeightPx * scale;
  const nextWidth = nextHeight * aspectRatio;
  return { dx: (nextWidth - startWidthPx) * sx, dy: (nextHeight - startHeightPx) * sy };
}

export interface SpreadsheetAppStatusElements {
  activeCell: HTMLElement;
  selectionRange: HTMLElement;
  activeValue: HTMLElement;
  /**
   * Optional status-bar element that is shown when the current collaboration session
   * is read-only (viewer/commenter).
   */
  readOnlyIndicator?: HTMLElement;
  /**
   * Optional status-bar element for Excel-like quick stats (sum of numeric values in selection).
   */
  selectionSum?: HTMLElement;
  /**
   * Optional status-bar element for Excel-like quick stats (average of numeric values in selection).
   */
  selectionAverage?: HTMLElement;
  /**
   * Optional status-bar element for Excel-like quick stats (count of non-empty cells in selection).
   */
  selectionCount?: HTMLElement;
}

export type UndoRedoState = {
  canUndo: boolean;
  canRedo: boolean;
  undoLabel: string | null;
  redoLabel: string | null;
};

export interface SpreadsheetSelectionSummary {
  /**
   * Sum of numeric values in the current selection.
   *
   * - Uses computed values for formula cells.
   * - Ignores non-numeric values (text, booleans, errors).
   * - `null` when there are no numeric values selected.
   */
  sum: number | null;
  /**
   * Average of numeric values in the current selection.
   *
   * - Uses computed values for formula cells.
   * - Ignores non-numeric values (text, booleans, errors).
   * - `null` when there are no numeric values selected.
   */
  average: number | null;
  /**
   * "Count" as shown by default in Excel's status bar: number of *non-empty* cells.
   *
   * This counts cells with a value or a formula, and ignores format-only cells
   * (styleId-only entries).
   */
  count: number;
  /**
   * Number of numeric values in the selection (Excel "Numerical Count").
   */
  numericCount: number;
  /**
   * Alias for `count` (explicit name for consumers that also expose `numericCount`).
   */
  countNonEmpty: number;
}

export type SpreadsheetAppCollabOptions = {
  wsUrl: string;
  docId: string;
  token?: string;
  /**
   * Legacy option used by older desktop collab wiring to toggle local durability.
   *
   * @deprecated Use `persistenceEnabled` (or the `collabPersistence` URL param) instead.
   */
  offlineEnabled?: boolean;
  /**
   * Local durability (offline-first) for the collaborative Yjs document.
   *
   * Defaults to enabled; can be disabled via `?collabPersistence=0` for debugging.
   */
  persistenceEnabled?: boolean;
  user: CollabUserIdentity;
  disableBc?: boolean;
};

function resolveCollabPersistenceEnabled(params: URLSearchParams): boolean {
  const raw = params.get("collabPersistence");
  if (raw != null) {
    const normalized = raw.trim().toLowerCase();
    if (normalized === "0" || normalized === "false" || normalized === "no" || normalized === "off") return false;
    if (normalized === "1" || normalized === "true" || normalized === "yes" || normalized === "on") return true;
  }

  // Backwards compatibility: `collabOffline=0` historically disabled local durability.
  // We now use `CollabSession.persistence`; treat it as an alias.
  const offlineRaw = params.get("collabOffline");
  if (offlineRaw != null) {
    const normalized = String(offlineRaw).trim().toLowerCase();
    if (normalized === "0" || normalized === "false" || normalized === "no" || normalized === "off") return false;
  }

  return true;
}

function resolveUseCanvasCharts(search: string = typeof window !== "undefined" ? window.location.search : ""): boolean {
  try {
    const params = new URLSearchParams(search);
    const raw = params.get("canvasCharts") ?? params.get("useCanvasCharts") ?? params.get("canvas_charts");
    if (raw != null) {
      const normalized = raw.trim().toLowerCase();
      if (normalized === "1" || normalized === "true" || normalized === "yes" || normalized === "on") return true;
      if (normalized === "0" || normalized === "false" || normalized === "no" || normalized === "off") return false;
    }
  } catch {
    // Ignore invalid URLSearchParams input.
  }

  // Vite exposes env vars via `import.meta.env`, but tests may also set Node-style env.
  const metaEnv = (import.meta as any)?.env as Record<string, unknown> | undefined;
  const viteValue = metaEnv?.VITE_CANVAS_CHARTS ?? metaEnv?.VITE_USE_CANVAS_CHARTS;
  if (typeof viteValue === "string" && viteValue.trim() !== "") {
    const normalized = viteValue.trim().toLowerCase();
    if (normalized === "1" || normalized === "true") return true;
    if (normalized === "0" || normalized === "false") return false;
  }
  if (typeof viteValue === "boolean") return viteValue;

  const nodeEnv = (globalThis as any)?.process?.env as Record<string, unknown> | undefined;
  const nodeValue = nodeEnv?.CANVAS_CHARTS ?? nodeEnv?.USE_CANVAS_CHARTS;
  if (typeof nodeValue === "string" && nodeValue.trim() !== "") {
    const normalized = nodeValue.trim().toLowerCase();
    if (normalized === "1" || normalized === "true") return true;
    if (normalized === "0" || normalized === "false") return false;
  }
  if (typeof nodeValue === "boolean") return nodeValue;

  return false;
}

function resolveCollabOptionsFromUrl(): SpreadsheetAppCollabOptions | null {
  if (typeof window === "undefined") return null;
  try {
    const url = new URL(window.location.href);
    const params = url.searchParams;
    const enabled = params.get("collab");
    if (enabled !== "1" && enabled !== "true") return null;

    const docId = params.get("collabDocId") ?? params.get("docId") ?? "";
    const wsUrl = params.get("collabWsUrl") ?? params.get("wsUrl") ?? "";
    if (!docId || !wsUrl) return null;

    // Tokens are accepted from either query params (legacy) or the URL hash (recommended),
    // but they should never remain in the address bar. Stash in session storage and scrub.
    const hashParams = new URLSearchParams(url.hash.startsWith("#") ? url.hash.slice(1) : url.hash);
    const tokenFromUrl =
      params.get("collabToken") ?? params.get("token") ?? hashParams.get("collabToken") ?? hashParams.get("token") ?? undefined;
    const tokenFromUrlTrimmed = typeof tokenFromUrl === "string" && tokenFromUrl.trim() !== "" ? tokenFromUrl : undefined;
    if (tokenFromUrlTrimmed) {
      storeCollabToken({ wsUrl, docId, token: tokenFromUrlTrimmed });
      // Best-effort: remove tokens from query/hash so we don't leak secrets into screenshots,
      // error reports, or browser history.
      try {
        params.delete("collabToken");
        params.delete("token");
        hashParams.delete("collabToken");
        hashParams.delete("token");
        url.hash = hashParams.toString();
        history.replaceState(null, "", url.toString());
      } catch {
        // ignore history errors
      }
    }

    const token = tokenFromUrlTrimmed ?? loadCollabToken({ wsUrl, docId }) ?? undefined;
    const identity = getCollabUserIdentity({ search: window.location.search });
    const disableBcRaw = params.get("collabDisableBc") ?? params.get("disableBc");
    const disableBc = disableBcRaw === "1" || disableBcRaw === "true" || disableBcRaw === "yes";
    const persistenceEnabled = resolveCollabPersistenceEnabled(params);
    return {
      wsUrl,
      docId,
      token,
      disableBc,
      offlineEnabled: persistenceEnabled,
      persistenceEnabled,
      user: identity,
    };
  } catch {
    return null;
  }
}

function resolveCollabOptionsFromStoredConnection(workbookId: string | undefined): SpreadsheetAppCollabOptions | null {
  if (typeof window === "undefined") return null;
  const workbookKey = String(workbookId ?? "").trim();
  if (!workbookKey) return null;

  const stored = loadCollabConnectionForWorkbook({ workbookKey });
  if (!stored) return null;

  // Tokens are persisted across app restarts on desktop via an OS-keychain-backed
  // encrypted store (see `collabTokenStore`), and cached into session-scoped
  // storage during startup.
  const token = loadCollabToken({ wsUrl: stored.wsUrl, docId: stored.docId });
  if (!token) return null;

  const persistenceEnabled = resolveCollabPersistenceEnabled(new URL(window.location.href).searchParams);
  return {
    wsUrl: stored.wsUrl,
    docId: stored.docId,
    token,
    offlineEnabled: persistenceEnabled,
    persistenceEnabled,
    user: getCollabUserIdentity({ search: window.location.search }),
  };
}

function resolveDrawingsDemoEnabledFromUrl(): boolean {
  if (typeof window === "undefined") return false;
  try {
    const params = new URL(window.location.href).searchParams;
    const raw = params.get("drawings") ?? params.get("drawing") ?? params.get("overlayDrawings");
    if (!raw) return false;
    const normalized = raw.trim().toLowerCase();
    return normalized === "1" || normalized === "true" || normalized === "yes" || normalized === "on";
  } catch {
    return false;
  }
}

export class SpreadsheetApp {
  private sheetId = "Sheet1";
  private readonly idle = new IdleTracker();
  private readonly computedValuesByCoord = new Map<string, Map<number, SpreadsheetValue>>();
  private computedValuesVersion = 0;
  private lastComputedValuesSheetId: string | null = null;
  private lastComputedValuesSheetCache: Map<number, SpreadsheetValue> | null = null;
  private uiReady = false;
  private readonly sheetNameResolver: SheetNameResolver | null;
  private readonly gridMode: DesktopGridMode;
  private readonly useCanvasCharts: boolean;
  /**
   * When enabled, comments are keyed by a sheet-qualified cell ref (`${sheetId}!A1`).
   *
   * This is required for collaboration (multi-sheet) to avoid cross-sheet collisions, but we
   * intentionally keep legacy A1-only comment refs in non-collab mode for back-compat with
   * existing persisted comment docs + tests.
   */
  private readonly collabMode: boolean;
  private readonly engine = new IdleTrackingEngine(
    new MockEngine(),
    this.idle,
    (changes) => this.invalidateComputedValues(changes),
    (changes) => this.applyComputedChanges(changes)
  );
  private readonly document = new DocumentController({ engine: this.engine });
  private readonly imageStore = new DesktopImageStore();
  private readonly sharedGridImageResolver: CanvasGridImageResolver = async (imageId) => this.document.getImageBlob(imageId);
  /**
   * In collaborative mode, keyboard undo/redo must use Yjs UndoManager semantics
   * (see `@formula/collab-undo`) so we never overwrite newer remote edits.
   *
   * This service is expected to be created by the CollabSession↔DocumentController
   * binder integration layer and injected into the app when collaboration is active.
   */
  private collabUndoService: UndoService | null = null;
  private readonly searchWorkbook: DocumentWorkbookAdapter;
  private readonly aiCellFunctions: AiCellFunctionEngine;
  private limits: GridLimits;
  private sharedGrid: DesktopSharedGrid | null = null;
  private sharedProvider: DocumentCellProvider | null = null;
  private readonly commentMetaByCoord = new Map<number, { resolved: boolean }>();
  private readonly commentPreviewByCoord = new Map<number, string>();
  private readonly commentThreadsByCellRef = new Map<string, Comment[]>();
  private sharedGridSelectionSyncInProgress = false;
  private sharedGridZoom = 1;

  private readonly workbookImageBitmaps = new ImageBitmapCache();
  private activeSheetBackgroundImageId: string | null = null;
  private activeSheetBackgroundBitmap: ImageBitmap | null = null;
  private activeSheetBackgroundLoadToken = 0;
  private activeSheetBackgroundAbort: AbortController | null = null;

  private wasmEngine: EngineClient | null = null;
  private wasmSyncSuspended = false;
  private wasmUnsubscribe: (() => void) | null = null;
  private wasmSyncPromise: Promise<void> = Promise.resolve();
  private auditingUnsubscribe: (() => void) | null = null;
  private externalRepaintUnsubscribe: (() => void) | null = null;
  private drawingsUnsubscribe: (() => void) | null = null;

  private gridCanvas: HTMLCanvasElement;
  private chartCanvas: HTMLCanvasElement;
  private drawingCanvas: HTMLCanvasElement;
  private readonly drawingGeom: DrawingGridGeometry;
  private drawingOverlay: DrawingOverlay;
  private readonly drawingChartRenderer: ChartRendererAdapter;
  private readonly drawingsDemoEnabled: boolean = resolveDrawingsDemoEnabledFromUrl();
  private drawingInteractionController: DrawingInteractionController | null = null;
  private drawingInteractionCallbacks: DrawingInteractionCallbacks | null = null;
  private readonly drawingImages: ImageStore;
  private drawingObjectsCache: { sheetId: string; objects: DrawingObject[]; source: unknown } | null = null;
  /**
   * Cached drawing objects for the active sheet.
   *
   * This avoids recomputing/allocating draw-object lists on pointermove hover paths.
   * `renderDrawings()` refreshes this list when the underlying drawing layer changes.
   */
  private drawingObjects: DrawingObject[] = [];
  private drawingHitTestIndex: HitTestIndex | null = null;
  private drawingHitTestIndexObjects: readonly DrawingObject[] | null = null;
  private selectedDrawingId: number | null = null;
  private readonly formulaChartModelStore = new FormulaChartModelStore();
  private nextDrawingImageId = 1;
  private insertImageInput: HTMLInputElement | null = null;
  private drawingViewportMemo:
    | {
        width: number;
        height: number;
        dpr: number;
      }
    | null = null;
  private referenceCanvas: HTMLCanvasElement;
  private auditingCanvas: HTMLCanvasElement;
  private presenceCanvas: HTMLCanvasElement | null = null;
  private selectionCanvas: HTMLCanvasElement;
  private gridCtx: CanvasRenderingContext2D;
  private chartCtx: CanvasRenderingContext2D;
  private referenceCtx: CanvasRenderingContext2D;
  private auditingCtx: CanvasRenderingContext2D;
  private presenceCtx: CanvasRenderingContext2D | null = null;
  private selectionCtx: CanvasRenderingContext2D;
  private presenceRenderer: PresenceRenderer | null = null;
  private remotePresences: GridPresence[] = [];

  private auditingRenderer = new AuditingOverlayRenderer();
  private auditingMode: "off" | AuditingMode = "off";
  private auditingTransitive = false;
  private auditingHighlights: AuditingOverlays = { precedents: new Set(), dependents: new Set() };
  private auditingErrors: { precedents: string | null; dependents: string | null } = { precedents: null, dependents: null };
  private readonly auditingCache = new Map<string, AuditingCacheEntry>();
  private auditingLegend!: HTMLDivElement;
  private auditingUpdateScheduled = false;
  private auditingNeedsUpdateAfterDrag = false;
  private auditingRequestId = 0;
  private auditingIdlePromise: Promise<void> = Promise.resolve();
  private auditingLastCellKey: string | null = null;
  private auditingWasRendered = false;
  private auditingNeedsClear = false;

  private readonly outlinesBySheet = new Map<string, Outline>();
  private getOutlineForSheet(sheetId: string): Outline {
    const key = String(sheetId ?? "");
    let outline = this.outlinesBySheet.get(key);
    if (!outline) {
      outline = new Outline();
      this.outlinesBySheet.set(key, outline);
    }
    return outline;
  }

  private outlineLayer: HTMLDivElement;
  private readonly outlineButtons = new Map<string, HTMLButtonElement>();

  private dpr = 1;
  private width = 0;
  private height = 0;
  private rootLeft = 0;
  private rootTop = 0;
  private rootPosLastMeasuredAtMs = 0;

  // Scroll offsets in CSS pixels relative to the sheet data origin (A1 at 0,0).
  private scrollX = 0;
  private scrollY = 0;
  private readonly cellWidth = 100;
  private readonly cellHeight = 24;
  private readonly rowHeaderWidth = 48;
  private readonly colHeaderHeight = 24;

  // Precomputed "visual" (i.e. not hidden) row/col indices.
  // Rebuilt only when outline visibility changes.
  private rowIndexByVisual: number[] = [];
  private colIndexByVisual: number[] = [];

  // Current viewport window (virtualized render range).
  private visibleRows: number[] = [];
  private visibleCols: number[] = [];
  private visibleRowStart = 0;
  private visibleColStart = 0;

  private viewportMappingState:
    | {
        scrollX: number;
        scrollY: number;
        viewportWidth: number;
        viewportHeight: number;
        rowCount: number;
        colCount: number;
      }
    | null = null;

  // Frozen panes for the active sheet. `frozenRows`/`frozenCols` are sheet
  // indices (0-based counts); `frozenWidth`/`frozenHeight` are viewport pixels
  // for the visible (i.e. not hidden) frozen rows/cols.
  private frozenRows = 0;
  private frozenCols = 0;
  private frozenVisibleRows: number[] = [];
  private frozenVisibleCols: number[] = [];
  private frozenWidth = 0;
  private frozenHeight = 0;

  // Maps from sheet row/col -> visual index (excluding hidden rows/cols).
  // These allow O(1) cell->pixel conversions for any in-bounds cell.
  private rowToVisual = new Map<number, number>();
  private colToVisual = new Map<number, number>();

  private readonly scrollbarThickness = 10;
  private vScrollbarTrack: HTMLDivElement;
  private vScrollbarThumb: HTMLDivElement;
  private hScrollbarTrack: HTMLDivElement;
  private hScrollbarThumb: HTMLDivElement;
  private lastScrollbarLayout:
    | { showV: boolean; showH: boolean; rowHeaderWidth: number; colHeaderHeight: number; thickness: number }
    | null = null;
  private lastScrollbarThumb: {
    vSize: number | null;
    vOffset: number | null;
    hSize: number | null;
    hOffset: number | null;
  } = { vSize: null, vOffset: null, hSize: null, hOffset: null };
  private readonly scrollbarThumbScratch = {
    v: { size: 0, offset: 0 },
    h: { size: 0, offset: 0 }
  };
  private scrollbarDrag:
    | { axis: "x" | "y"; pointerId: number; grabOffset: number; thumbTravel: number; trackStart: number; maxScroll: number }
    | null = null;

  private selection: SelectionState;
  private selectionRenderer = new SelectionRenderer();
  private formulaSelectionRenderer = new SelectionRenderer({
    fillColor: "transparent",
    borderColor: "transparent",
    activeBorderColor: resolveCssVar("--formula-grid-selection-border", { fallback: resolveCssVar("--selection-border", { fallback: "transparent" }) }),
    fillHandleColor: "transparent",
    borderWidth: 2,
    activeBorderWidth: 3,
    fillHandleSize: 0,
  });
  private readonly selectionListeners = new Set<(selection: SelectionState) => void>();
  private readonly scrollListeners = new Set<(scroll: { x: number; y: number }) => void>();
  private readonly zoomListeners = new Set<(zoom: number) => void>();
  private readonly formulaBarOverlayListeners = new Set<() => void>();

  private editState = false;
  private readonly editStateListeners = new Set<(isEditing: boolean) => void>();
  private focusTargetProvider: (() => HTMLElement | null) | null = null;

  private editor: CellEditorOverlay;
  private suppressFocusRestoreOnNextCommandCommit = false;
  private formulaBar: FormulaBarView | null = null;
  private formulaBarCompletion: FormulaBarTabCompletionController | null = null;
  private formulaRangePreviewTooltip: HTMLDivElement | null = null;
  private formulaRangePreviewTooltipVisible = false;
  private formulaRangePreviewTooltipLastKey: string | null = null;
  private formulaEditCell: { sheetId: string; cell: CellCoord } | null = null;
  private keyboardRangeSelectionActive = false;
  private referencePreview: { start: CellCoord; end: CellCoord } | null = null;
  private referenceHighlights: Array<{ start: CellCoord; end: CellCoord; color: string; active: boolean }> = [];
  private referenceHighlightsSource: Array<{
    range: { startRow: number; startCol: number; endRow: number; endCol: number; sheet?: string };
    color: string;
    active?: boolean;
  }> = [];
  private fillPreviewRange: Range | null = null;
  private showFormulas = false;

  private dragState: DragState | null = null;
  private dragPointerPos: { x: number; y: number } | null = null;
  private dragAutoScrollRaf: number | null = null;

  private drawingGesture: DrawingGestureState | null = null;

  private resizeObserver: ResizeObserver;
  private disposed = false;
  private readonly domAbort = new AbortController();
  private commentsDocUpdateListener: (() => void) | null = null;
  private stopCommentsRootObserver: (() => void) | null = null;
  private commentsUndoScopeAdded = false;

  private readonly inlineEditController: InlineEditController;

  private readonly currentUser: CommentAuthor;
  private readonly commentsDoc: Y.Doc;
  private readonly commentManager: CommentManager;
  private commentsPanelVisible = false;
  private stopCommentPersistence: (() => void) | null = null;
  private commentIndexVersion = 0;
  private lastHoveredCommentCellKey: number | null = null;
  private lastHoveredCommentIndexVersion = -1;
  private sharedHoverCellKey: number | null = null;
  private sharedHoverCellRect: { x: number; y: number; width: number; height: number } | null = null;
  private sharedHoverCellHasComment = false;
  private sharedHoverCellCommentIndexVersion = -1;

  private collabSession: CollabSession | null = null;
  private collabBinderOrigin: object | null = null;
  private imageBytesBinder: ImageBytesBinder | null = null;
  private collabEncryptionKeyStore: CollabEncryptionKeyStore | null = null;
  private reservedRootGuardToastUnsubscribe: (() => void) | null = null;
  private readOnly = false;
  private readOnlyRole: string | null = null;
  private collabPermissionsUnsubscribe: (() => void) | null = null;
  private collabBinder: { destroy: () => void; rehydrate?: () => void; whenIdle?: () => Promise<void> } | null = null;
  private collabBinderInitPromise: Promise<{ destroy: () => void; rehydrate?: () => void; whenIdle?: () => Promise<void> }> | null =
    null;
  private collabSelectionUnsubscribe: (() => void) | null = null;
  private collabPresenceUnsubscribe: (() => void) | null = null;
  private conflictUi: ConflictUiController | null = null;
  private conflictUiContainer: HTMLDivElement | null = null;
  private collaboratorsListUi: CollaboratorsListUiController | null = null;
  private collaboratorsListContainer: HTMLDivElement | null = null;
  private structuralConflictUi: StructuralConflictUiController | null = null;
  private pendingFormulaConflicts: any[] = [];
  private pendingStructuralConflicts: any[] = [];

  private readonly chartStore: ChartStore;
  private chartTheme: ChartTheme = FALLBACK_CHART_THEME;
  private selectedChartId: string | null = null;
  private readonly chartModels = new Map<string, ChartModel>();
  private readonly chartRenderer: ChartRendererAdapter;
  /**
   * Chart ids whose cached `ChartModel.series[*].{categories,values,xValues,yValues}.cache`
   * need to be refreshed before the next on-screen render.
   *
   * We keep this separate from `chartModels` so we can avoid rescanning chart data on
   * every scroll, while still ensuring charts that were off-screen during edits refresh
   * when scrolled into view.
   */
  private readonly dirtyChartIds = new Set<string>();
  /**
   * Whether a chart's referenced ranges include at least one formula cell.
   *
   * This is used as a conservative signal for when a recalculation event may affect chart
   * values even if the triggering cell deltas were outside the chart's direct ranges (e.g.
   * a chart plots `B1` where `B1` is a formula referencing `A1`).
   */
  private readonly chartHasFormulaCells = new Map<string, boolean>();
  private readonly chartRangeRectsCache = new Map<
    string,
    {
      signature: string;
      ranges: Array<{ sheetId: string; startRow: number; endRow: number; startCol: number; endCol: number }>;
    }
  >();
  private readonly chartCanvasStoreAdapter: ChartCanvasStoreAdapter;
  private chartRecordLookupCache: { list: readonly ChartRecord[]; map: Map<string, ChartRecord> } | null = null;
  private readonly chartOverlayImages: ImageStore = { get: () => undefined, set: () => {} };
  private chartOverlayGeom: DrawingGridGeometry | null = null;
  private chartSelectionOverlay: DrawingOverlay | null = null;
  private chartSelectionViewportMemo: { width: number; height: number; dpr: number } | null = null;
  private chartSelectionCanvas: HTMLCanvasElement | null = null;
  private chartDrawingInteraction: DrawingInteractionController | null = null;
  private chartDragState:
    | {
        pointerId: number;
        chartId: string;
        mode: "move" | "resize";
        resizeHandle?: ResizeHandle;
        startClientX: number;
        startClientY: number;
        startAnchor: ChartRecord["anchor"];
      }
    | null = null;
  private chartDragAbort: AbortController | null = null;
  private commentsPanel!: HTMLDivElement;
  private commentsPanelThreads!: HTMLDivElement;
  private commentsPanelCell!: HTMLDivElement;
  private newCommentInput!: HTMLInputElement;
  private newCommentSubmitButton!: HTMLButtonElement;
  private commentsPanelReadOnlyHint!: HTMLDivElement;
  private commentTooltip!: HTMLDivElement;
  private commentTooltipVisible = false;

  private sheetViewBinder: SheetViewBinder | null = null;

  private renderScheduled = false;
  private pendingRenderMode: "full" | "scroll" = "full";
  private statusUpdateScheduled = false;
  private selectionStatsFormatter: Intl.NumberFormat | null = null;
  private selectionSummaryCache:
    | {
        sheetId: string;
        sheetContentVersion: number;
        workbookContentVersion: number;
        computedValuesVersion: number;
        selectionHasFormula: boolean;
        rangesKey: number[];
        summary: SpreadsheetSelectionSummary;
      }
    | null = null;
  private chartContentRefreshScheduled = false;
  private drawingsRenderScheduled = false;
  private clipboardProviderPromise: ReturnType<typeof createClipboardProvider> | null = null;
  private clipboardCopyContext:
    | {
        range: Range;
        payload: { text?: string; html?: string; rtf?: string };
        cells: Array<Array<{ value: unknown; formula: string | null; styleId: number }>>;
      }
    | null = null;
  private dlpContext: ReturnType<typeof createDesktopDlpContext> | null = null;
  private encryptedRangeManager: EncryptedRangeManager | null = null;
  constructor(
    private root: HTMLElement,
    private status: SpreadsheetAppStatusElements,
    opts: {
      workbookId?: string;
      /**
       * Enables sheet-qualified comment cell refs (`${sheetId}!A1`) without requiring a full
       * CollabSession. Intended for desktop collab mode and unit/e2e harnesses.
       */
      collabMode?: boolean;
      limits?: GridLimits;
      formulaBar?: HTMLElement;
      collab?: SpreadsheetAppCollabOptions;
      sheetNameResolver?: SheetNameResolver;
      inlineEdit?: {
        llmClient?: InlineEditLLMClient;
        model?: string;
        auditStore?: AIAuditStore;
        onWorkbookContextBuildStats?: (stats: WorkbookContextBuildStats) => void;
      };
      /**
       * Enables interactive drawing manipulation (click-to-select + drag/resize).
       *
       * Defaults to enabled in shared-grid mode (drawings are otherwise not interactable)
       * and disabled in legacy mode unless explicitly enabled.
       */
      enableDrawingInteractions?: boolean;
    } = {}
  ) {
    this.sheetNameResolver = opts.sheetNameResolver ?? null;
    this.searchWorkbook = new DocumentWorkbookAdapter({ document: this.document, sheetNameResolver: this.sheetNameResolver ?? undefined });
    this.gridMode = resolveDesktopGridMode();
    this.useCanvasCharts = resolveUseCanvasCharts();
    this.limits =
      opts.limits ??
      (this.gridMode === "shared"
        ? { ...DEFAULT_GRID_LIMITS }
        : {
            // Legacy renderer relies on eagerly-built row/col visibility caches; keep its
            // default caps small to avoid O(N) work and large Map allocations.
            ...DEFAULT_GRID_LIMITS,
            maxRows: DEFAULT_DESKTOP_LOAD_MAX_ROWS,
            maxCols: DEFAULT_DESKTOP_LOAD_MAX_COLS
          });
    this.selection = createSelection({ row: 0, col: 0 }, this.limits);
    const rawCollab = opts.collab ?? resolveCollabOptionsFromUrl() ?? resolveCollabOptionsFromStoredConnection(opts.workbookId);
    const jwtPermissions = rawCollab?.token ? tryDeriveCollabSessionPermissionsFromJwtToken(rawCollab.token) : null;
    const collab = rawCollab
      ? {
          ...rawCollab,
          user: jwtPermissions?.userId ? overrideCollabUserIdentityId(rawCollab.user, jwtPermissions.userId) : rawCollab.user,
        }
      : null;
    const localWorkbookId = opts.workbookId ?? collab?.docId ?? "local-workbook";

    const collabEnabled = Boolean(collab);
    this.collabMode = collabEnabled || Boolean(opts.collabMode);
    this.currentUser = collab ? { id: collab.user.id, name: collab.user.name } : { id: "local", name: t("chat.role.user") };

    if (collab) {
      // Persist non-secret collab connection metadata so desktop startup can auto-reconnect
      // after restart. Tokens are stored separately in the secure token store.
      try {
        const workbookKey = String(opts.workbookId ?? "").trim();
        if (workbookKey) {
          saveCollabConnectionForWorkbook({ workbookKey, wsUrl: collab.wsUrl, docId: collab.docId });
        }
      } catch {
        // ignore storage failures
      }

      const sessionPermissions = jwtPermissions
        ? {
            role: jwtPermissions.role,
            rangeRestrictions: jwtPermissions.rangeRestrictions,
            userId: jwtPermissions.userId ?? collab.user.id,
          }
        : { role: "editor", rangeRestrictions: [], userId: collab.user.id };

      // Best-effort: cache the token for this (wsUrl, docId) in session-scoped storage
      // so UI surfaces (sharing/reconnect) can access it without leaving it in the URL.
      if (collab.token) {
        try {
          storeCollabToken({ wsUrl: collab.wsUrl, docId: collab.docId, token: collab.token });
        } catch {
          // ignore
        }
      }

      // Binder writes (DocumentController -> Yjs) must use an origin that is *distinct* from
      // `session.origin` so Yjs writes performed directly through the session (e.g. versioning
      // operations) still propagate back into the DocumentController.
      const binderOrigin = { type: "desktop-document-controller:binder" };
      this.collabBinderOrigin = binderOrigin;

      const persistenceEnabled = collab.persistenceEnabled ?? collab.offlineEnabled ?? true;
      const persistence = persistenceEnabled === false ? undefined : new IndexedDbCollabPersistence();

      const devEncryption =
        typeof window !== "undefined"
          ? resolveDevCollabEncryptionFromSearch({
              search: window.location.search,
              docId: collab.docId,
              defaultSheetId: this.sheetId,
              resolveSheetIdByName: (name) => this.sheetNameResolver?.getSheetIdByName(name) ?? null,
            })
          : null;
      const encryptionKeyStore = new CollabEncryptionKeyStore();
      this.collabEncryptionKeyStore = encryptionKeyStore;
      // When dev-mode encryption is enabled (`collabEncrypt=1`), the encryption key is
      // derived deterministically from `docId` and does not rely on the persisted key store.
      // Skip the potentially expensive/unsupported Tauri keychain hydration so binder startup
      // isn't delayed in simple dev server scenarios.
      const encryptionKeyStoreHydrated = devEncryption ? Promise.resolve() : encryptionKeyStore.hydrateDoc(collab.docId).catch(() => {});
      let encryptionPolicy: ReturnType<typeof createEncryptionPolicyFromDoc> | null = null;

      this.collabSession = createCollabSession({
        docId: collab.docId,
        persistence,
        connection: {
          wsUrl: collab.wsUrl,
          docId: collab.docId,
          token: collab.token,
          disableBc: collab.disableBc,
        },
        comments: {
          // Canonicalize older Array-backed comment roots into the deterministic
          // Map-backed schema after hydration.
          migrateLegacyArrayToMap: true,
        },
        presence: {
          user: collab.user,
          activeSheet: this.sheetId,
          staleAfterMs: 60_000,
          throttleMs: 50,
        },
        encryption:
          devEncryption ??
          ({
            shouldEncryptCell: (cell) => encryptionPolicy?.shouldEncryptCell(cell) ?? false,
            keyForCell: (cell) => {
              const keyId = encryptionPolicy?.keyIdForCell(cell);
              if (!keyId) return null;
              return encryptionKeyStore.getCachedKey(collab.docId, keyId);
            },
          } as any),
        // Enable formula/value conflict monitoring in collab mode.
        formulaConflicts: {
          localUserId: sessionPermissions.userId,
          mode: "formula+value",
          onConflict: (conflict: any) => {
            // Conflicts are surfaced via a minimal DOM UI (ConflictUiController).
            // To exercise manually, edit the same formula concurrently in two clients.
            if (this.conflictUi) {
              this.conflictUi.addConflict(conflict);
            } else {
              // Conflicts can be detected before the UI overlay is mounted (e.g. during
              // initial sync). Queue until the UI is ready.
              this.pendingFormulaConflicts.push(conflict);
            }
          },
        },
        // Enable structural conflict monitoring (move/delete-vs-edit/content/format) in collab mode.
         cellConflicts: {
           localUserId: sessionPermissions.userId,
           onConflict: (conflict: any) => {
             if (this.structuralConflictUi) {
               this.structuralConflictUi.addConflict(conflict);
            } else {
              // Conflicts can be detected before the UI overlay is mounted (e.g. from persisted
              // structural op logs). Queue them until the conflict UI overlay is mounted.
              this.pendingStructuralConflicts.push(conflict);
            }
          },
        },
      });
      encryptionPolicy = createEncryptionPolicyFromDoc(this.collabSession.doc);

      // If the sync-server reserved root guard is enabled, writing to in-doc versioning /
      // branching roots will cause the server to close the websocket (1008 "reserved root
      // mutation"). Surface this as an actionable toast even if the relevant panels are
      // not currently open.
      this.reservedRootGuardToastUnsubscribe?.();
      this.reservedRootGuardToastUnsubscribe = null;
      try {
        let toastShown = false;
        this.reservedRootGuardToastUnsubscribe = subscribeToReservedRootGuardDisconnect(
          this.collabSession.provider as any,
          (detected) => {
            if (!detected) {
              toastShown = false;
              return;
            }
            if (toastShown) return;
            toastShown = true;
            try {
              showToast(reservedRootGuardUiMessage(), "error", { timeoutMs: 15_000 });
            } catch {
              // Best-effort; `showToast` requires a DOM #toast-root and should never block startup.
            }
          },
        );
      } catch {
        // ignore (defensive: never block collab startup on toast wiring)
      }
      if (devEncryption && typeof window !== "undefined") {
        try {
          const params = new URL(window.location.href).searchParams;
          const range = params.get("collabEncryptRange") ?? "Sheet1!A1:C10";
          showToast(
            `Dev: collab cell encryption enabled (${range}). Open a second client without collabEncrypt to verify masked reads (###).`,
            "info",
            { timeoutMs: 10_000 }
          );
        } catch {
          // Best-effort; `showToast` requires a DOM #toast-root and should never block startup.
        }
      }

      // Track permissions changes so the desktop UI can immediately reflect read-only mode
      // (viewer/commenter) and avoid local-only edits.
      const sessionForPermissions = this.collabSession;
      const onPermissionsChanged = (sessionForPermissions as any).onPermissionsChanged;
      if (typeof onPermissionsChanged === "function") {
        this.collabPermissionsUnsubscribe = onPermissionsChanged.call(sessionForPermissions, () => this.syncReadOnlyState());
      } else {
        // Backwards-compat / test stubs: if the session doesn't expose a subscription API,
        // wrap `setPermissions` as a best-effort signal.
        const originalSetPermissions = (sessionForPermissions as any).setPermissions;
        const isMock =
          typeof originalSetPermissions === "function" &&
          // Vitest/Jest-style mock functions expose a `.mock` property. Avoid wrapping them so
          // unit tests can continue to assert on call counts/args.
          typeof (originalSetPermissions as any).mock === "object";
        if (typeof originalSetPermissions === "function" && !isMock) {
          (sessionForPermissions as any).setPermissions = (permissions: any) => {
            originalSetPermissions.call(sessionForPermissions, permissions);
            this.syncReadOnlyState();
          };
        }
      }

      this.encryptedRangeManager = createEncryptedRangeManagerForSession(this.collabSession);

      // Populate `modifiedBy` metadata for any direct `CollabSession.setCell*` writes
      // (used by some conflict resolution + versioning flows) and ensure downstream
      // conflict UX can attribute local edits correctly.
      try {
        this.collabSession.setPermissions(sessionPermissions);
      } catch (err) {
        // JWT payloads are intentionally decoded without signature verification (best-effort so
        // collab links can bootstrap quickly). Treat any derived permissions as untrusted data and
        // never crash the app if the token contains malformed `rangeRestrictions`.
        const message = err instanceof Error ? err.message : String(err);
        console.warn("Failed to apply collab permissions from token; falling back to defaults.", message);

         const fallbackPermissions = { ...sessionPermissions, rangeRestrictions: [] };
         try {
           // Fallback policy: keep the derived role/userId (if present) but drop all range
           // restrictions when they fail validation. Server-side access control is still enforced
           // by the sync server; this only prevents desktop startup from being DoS'd by a bad URL.
           this.collabSession.setPermissions(fallbackPermissions);
         } catch (fallbackErr) {
           const fallbackMessage = fallbackErr instanceof Error ? fallbackErr.message : String(fallbackErr);
           console.warn("Failed to apply fallback collab permissions; continuing with viewer access.", fallbackMessage);
           // Last-resort safe default (should never happen).
           this.collabSession.setPermissions({ role: "viewer", rangeRestrictions: [], userId: sessionPermissions.userId });
         }
       }

      this.sheetViewBinder = bindSheetViewToCollabSession({
        session: this.collabSession,
        documentController: this.document,
        origin: binderOrigin,
      });

      const undoScope: Array<Y.AbstractType<any>> = [
        this.collabSession.cells,
        this.collabSession.sheets,
        this.collabSession.metadata,
        this.collabSession.namedRanges,
      ];

      // Include comments in the undo scope when the comments root already exists in
      // the doc. Avoid instantiating `doc.getMap("comments")` pre-hydration because
      // older documents may still use an Array-backed schema.
      try {
        if (this.collabSession.doc.share.get("comments")) {
          const root = getCommentsRoot(this.collabSession.doc);
          undoScope.push(root.kind === "map" ? root.map : root.array);
        }
      } catch {
        // Best-effort; never block app startup on comment schema issues.
      }

      const undoService = createUndoService({
        mode: "collab",
        doc: this.collabSession.doc,
        scope: undoScope,
        origin: binderOrigin,
      }) as UndoService & { origin?: any };

      // The binder expects an explicit origin token for echo suppression.
      undoService.origin = binderOrigin;
      this.setCollabUndoService(undoService);

      // Ensure conflict monitors treat binder + undo transactions as local so they
      // can log structural operations and avoid misclassifying undo/redo edits as
      // remote.
      for (const origin of undoService.localOrigins ?? []) {
        this.collabSession.localOrigins.add(origin);
      }

      // Comments sync through the shared collaborative Y.Doc when collab is enabled.
      this.commentsDoc = this.collabSession.doc;
      // Avoid eagerly instantiating the `comments` root type before the provider has
      // hydrated the document; older docs may still use a legacy Array-backed schema.
      this.commentManager = createCommentManagerForDoc({
        doc: this.commentsDoc,
        // Gate comment writes on the current CollabSession permissions. We pass a
        // callback (rather than a snapshot boolean) so role updates (if any) are
        // reflected in subsequent comment mutations.
        canComment: () => this.collabSession?.canComment() ?? true,
        // Ensure comment edits are tracked by the binder-origin collaborative UndoManager
        // (so Cmd/Ctrl+Z reverts comment add/edit/reply/resolve just like cell edits).
        transact: (fn) => {
          // If the comments root was created lazily (e.g. first comment add), ensure it
          // is added to the UndoManager scope before we perform the tracked transaction.
          // This keeps early comment edits undoable even if they happen before provider sync.
          this.ensureCommentsUndoScope(null, { allowCreateBeforeSync: true });
          const transact = (undoService as any)?.transact;
          if (typeof transact === "function") {
            transact(fn);
            return;
          }
          // Fallback: run directly with the binder origin so any externally created
          // UndoManager tracking `binderOrigin` still captures the transaction.
          this.commentsDoc.transact(fn, binderOrigin);
        },
      });

      const binderPromise = (async () => {
        // Ensure any previously-imported encryption keys are loaded into the in-memory
        // cache before the binder reads encrypted cells.
        await encryptionKeyStoreHydrated;

        return await bindCollabSessionToDocumentController({
          session: this.collabSession,
          documentController: this.document,
          undoService,
          defaultSheetId: this.sheetId,
          userId: sessionPermissions.userId,
          onEditRejected: (rejected) => {
            showCollabEditRejectedToast(rejected);
          },
          // Opt into binder write semantics required for reliable causal conflict detection.
          // (E.g. represent clears as explicit `formula=null` markers so FormulaConflictMonitor
          // can reason about delete-vs-overwrite concurrency deterministically.)
           formulaConflictsMode: "formula+value",
         });
      })();

      this.collabBinderInitPromise = binderPromise;

      void binderPromise
        .then((binder: any) => {
          if (this.disposed) {
            binder.destroy();
            return;
          }
          this.collabBinder = binder;
        })
        .catch((err: any) => {
          console.error("Failed to bind collab session to DocumentController", err);
        })
        .finally(() => {
          if (this.collabBinderInitPromise === binderPromise) {
            this.collabBinderInitPromise = null;
          }
        });
    } else {
      this.commentsDoc = new Y.Doc();
      this.commentManager = new CommentManager(this.commentsDoc);
    }

    if (this.gridMode === "legacy") {
      const outline = this.getOutlineForSheet(this.sheetId);
      // Seed a simple outline group: rows 2-4 with a summary row at 5 (Excel 1-based indices).
      outline.groupRows(2, 4);
      outline.recomputeOutlineHiddenRows();
      // And columns 2-4 with a summary column at 5.
      outline.groupCols(2, 4);
      outline.recomputeOutlineHiddenCols();
    }

    if (!collabEnabled) {
      // Seed data for navigation tests (used range ends at D5).
      this.document.setCellValue(this.sheetId, { row: 0, col: 0 }, "Seed");
      this.document.setCellValue(this.sheetId, { row: 0, col: 1 }, {
        text: "Rich Bold",
        runs: [
          { start: 0, end: 5, style: {} },
          { start: 5, end: 9, style: { bold: true } }
        ]
      });
      this.document.setCellValue(this.sheetId, { row: 4, col: 3 }, "BottomRight");

      // Seed a small data range for the demo chart without expanding the used range past D5.
      this.document.setCellValue(this.sheetId, { row: 1, col: 0 }, "A");
      this.document.setCellValue(this.sheetId, { row: 1, col: 1 }, 2);
      this.document.setCellValue(this.sheetId, { row: 2, col: 0 }, "B");
      this.document.setCellValue(this.sheetId, { row: 2, col: 1 }, 4);
      this.document.setCellValue(this.sheetId, { row: 3, col: 0 }, "C");
      this.document.setCellValue(this.sheetId, { row: 3, col: 1 }, 3);
      this.document.setCellValue(this.sheetId, { row: 4, col: 0 }, "D");
      this.document.setCellValue(this.sheetId, { row: 4, col: 1 }, 5);
    }

    // Best-effort: keep the WASM engine worker hydrated from the DocumentController.
    // When the WASM module isn't available (e.g. local dev without building it),
    // the app continues to operate using the in-process mock engine.
    void this.initWasmEngine();

    this.gridCanvas = document.createElement("canvas");
    this.gridCanvas.className = "grid-canvas grid-canvas--base";
    this.gridCanvas.setAttribute("aria-hidden", "true");

    this.drawingCanvas = document.createElement("canvas");
    this.drawingCanvas.className = "drawing-layer drawing-layer--overlay grid-canvas--drawings";
    this.drawingCanvas.setAttribute("aria-hidden", "true");
    this.drawingCanvas.setAttribute("data-testid", "drawing-layer-canvas");
    if (this.gridMode === "shared") {
      // Shared-grid overlay stacking is expressed via CSS classes (see charts-overlay.css).
      this.drawingCanvas.classList.add("drawing-layer--shared", "grid-canvas--shared-drawings");
    }

    this.chartCanvas = document.createElement("canvas");
    this.chartCanvas.className = "grid-canvas grid-canvas--chart";
    this.chartCanvas.setAttribute("aria-hidden", "true");
    if (this.gridMode === "shared") {
      // Shared-grid overlay stacking is expressed via CSS classes (see charts-overlay.css).
      this.chartCanvas.classList.add("grid-canvas--shared-chart");
    }

    this.chartOverlayGeom = {
      cellOriginPx: (cell) => this.chartCellOriginPx(cell),
      cellSizePx: (cell) => this.chartCellSizePx(cell),
    };

    this.chartSelectionCanvas = document.createElement("canvas");
    // Chart selection handles are rendered on a separate overlay canvas. Keep the base
    // class list minimal and use semantic classes so CSS can control stacking.
    this.chartSelectionCanvas.className = "grid-canvas chart-selection-canvas";
    this.chartSelectionCanvas.setAttribute("aria-hidden", "true");
    if (this.gridMode === "shared") {
      // Match the selection canvas z-index so chart handles are drawn above charts in shared mode.
      this.chartSelectionCanvas.classList.add("grid-canvas--shared-selection");
    }
    this.chartSelectionOverlay = new DrawingOverlay(this.chartSelectionCanvas, this.chartOverlayImages, this.chartOverlayGeom);

    this.referenceCanvas = document.createElement("canvas");
    this.referenceCanvas.className = "grid-canvas grid-canvas--content";
    this.referenceCanvas.setAttribute("aria-hidden", "true");
    this.auditingCanvas = document.createElement("canvas");
    this.auditingCanvas.className = "grid-canvas grid-canvas--auditing";
    this.auditingCanvas.setAttribute("aria-hidden", "true");
    if (collabEnabled && this.gridMode !== "shared") {
      // Remote presence overlays should render above auditing highlights but below
      // the local selection layer.
      this.presenceCanvas = document.createElement("canvas");
      this.presenceCanvas.className = "grid-canvas grid-canvas--presence";
      this.presenceCanvas.setAttribute("aria-hidden", "true");
    }
    this.selectionCanvas = document.createElement("canvas");
    this.selectionCanvas.className = "grid-canvas grid-canvas--selection";
    this.selectionCanvas.setAttribute("aria-hidden", "true");
    if (this.gridMode === "shared") {
      // Shared-grid overlay stacking is expressed via CSS classes (see charts-overlay.css).
      this.selectionCanvas.classList.add("grid-canvas--shared-selection");
    }

    this.root.appendChild(this.gridCanvas);
    this.root.appendChild(this.drawingCanvas);
    this.root.appendChild(this.chartCanvas);
    this.root.appendChild(this.referenceCanvas);
    this.root.appendChild(this.auditingCanvas);
    if (this.presenceCanvas) this.root.appendChild(this.presenceCanvas);
    this.root.appendChild(this.selectionCanvas);
    this.root.appendChild(this.chartSelectionCanvas);
    // Avoid allocating a fresh `{row,col}` object for every chart cell lookup.
    const chartCoordScratch = { row: 0, col: 0 };
    const getChartCellValue = (sheetId: string, row: number, col: number): unknown => {
      chartCoordScratch.row = row;
      chartCoordScratch.col = col;
      const state = this.document.getCell(sheetId, chartCoordScratch) as {
        value: unknown;
        formula: string | null;
      };
      if (state?.formula != null) {
        // Charts should use computed values for formulas (show-formulas is a display-only toggle).
        return this.getCellComputedValueForSheetInternal(sheetId, chartCoordScratch);
      }
      const value = state?.value ?? null;
      return isRichTextValue(value) ? value.text : value;
    };
    this.chartStore = new ChartStore({
      defaultSheet: this.sheetId,
      sheetNameResolver: this.sheetNameResolver,
      getCellValue: getChartCellValue,
      // Creating/removing charts should not force a full data re-scan for *every* existing chart.
      // `renderCharts(false)` updates chart positioning and ensures newly-created charts have a
      // cached ChartModel. Data refreshes happen only for charts marked dirty by cell/computed
      // changes (see `dirtyChartIds`).
      onChange: () => {
        if (this.useCanvasCharts) this.renderDrawings();
        else this.renderCharts(false);
      }
    });

    this.chartCanvasStoreAdapter = new ChartCanvasStoreAdapter({
      getChart: (chartId) => this.getChartRecordById(chartId),
      getCellValue: getChartCellValue,
      resolveSheetId: (token) => this.resolveSheetIdByName(token),
      getSeriesColors: () => this.chartTheme.seriesColors,
      maxDataCells: MAX_CHART_DATA_CELLS,
    });

    const chartRendererStore: ChartRendererStore = {
      getChartModel: (chartId) => this.chartModels.get(chartId),
      getChartData: () => undefined,
      getChartTheme: () => ({ seriesColors: this.chartTheme.seriesColors }),
    };
    this.chartRenderer = new ChartRendererAdapter(chartRendererStore);

    this.outlineLayer = document.createElement("div");
    this.outlineLayer.className = "outline-layer";
    if (this.gridMode === "shared") {
      // Shared-grid overlay stacking is expressed via CSS classes (see charts-overlay.css).
      this.outlineLayer.classList.add("outline-layer--shared");
    }
    this.root.appendChild(this.outlineLayer);

    // Minimal scrollbars (drawn as DOM overlays, like the React CanvasGrid).
    this.vScrollbarTrack = document.createElement("div");
    this.vScrollbarTrack.setAttribute("aria-hidden", "true");
    this.vScrollbarTrack.setAttribute("data-testid", "scrollbar-track-y");
    this.vScrollbarTrack.className = "grid-scrollbar-track grid-scrollbar-track--vertical";

    this.vScrollbarThumb = document.createElement("div");
    this.vScrollbarThumb.setAttribute("aria-hidden", "true");
    this.vScrollbarThumb.setAttribute("data-testid", "scrollbar-thumb-y");
    this.vScrollbarThumb.className = "grid-scrollbar-thumb";
    this.vScrollbarTrack.appendChild(this.vScrollbarThumb);
    this.root.appendChild(this.vScrollbarTrack);

    this.hScrollbarTrack = document.createElement("div");
    this.hScrollbarTrack.setAttribute("aria-hidden", "true");
    this.hScrollbarTrack.setAttribute("data-testid", "scrollbar-track-x");
    this.hScrollbarTrack.className = "grid-scrollbar-track grid-scrollbar-track--horizontal";

    this.hScrollbarThumb = document.createElement("div");
    this.hScrollbarThumb.setAttribute("aria-hidden", "true");
    this.hScrollbarThumb.setAttribute("data-testid", "scrollbar-thumb-x");
    this.hScrollbarThumb.className = "grid-scrollbar-thumb";
    this.hScrollbarTrack.appendChild(this.hScrollbarThumb);
    this.root.appendChild(this.hScrollbarTrack);

    this.commentsPanel = this.createCommentsPanel();
    this.root.appendChild(this.commentsPanel);

    this.commentTooltip = this.createCommentTooltip();
    this.root.appendChild(this.commentTooltip);

    this.auditingLegend = this.createAuditingLegend();
    this.root.appendChild(this.auditingLegend);

    const gridCtx = this.gridCanvas.getContext("2d");
    const chartCtx = this.chartCanvas.getContext("2d");
    const referenceCtx = this.referenceCanvas.getContext("2d");
    const auditingCtx = this.auditingCanvas.getContext("2d");
    const presenceCtx = this.presenceCanvas ? this.presenceCanvas.getContext("2d") : null;
    const selectionCtx = this.selectionCanvas.getContext("2d");
    if (!gridCtx || !chartCtx || !referenceCtx || !auditingCtx || (this.presenceCanvas && !presenceCtx) || !selectionCtx) {
      throw new Error("Canvas 2D context not available");
    }
    this.gridCtx = gridCtx;
    this.chartCtx = chartCtx;
    this.referenceCtx = referenceCtx;
    this.auditingCtx = auditingCtx;
    this.presenceCtx = presenceCtx;
    this.selectionCtx = selectionCtx;
    if (this.presenceCanvas) {
      this.presenceRenderer = new PresenceRenderer();
    }

    this.editor = new CellEditorOverlay(this.root, {
      onCommit: (commit) => {
        const suppressFocusRestore =
          commit.reason === "command" && this.suppressFocusRestoreOnNextCommandCommit;
        this.suppressFocusRestoreOnNextCommandCommit = false;
        this.updateEditState();
        this.applyEdit(this.sheetId, commit.cell, commit.value);

        if (commit.reason !== "command") {
          const next = navigateSelectionByKey(
            this.selection,
            commit.reason === "enter" ? "Enter" : "Tab",
            { shift: commit.shift, primary: false },
            this.usedRangeProvider(),
            this.limits
          );

          if (next) this.selection = next;
        }
        this.ensureActiveCellVisible();
        this.scrollCellIntoView(this.selection.active);
        if (this.sharedGrid) this.syncSharedGridSelectionFromState({ scrollIntoView: false });
        this.refresh();
        if (!suppressFocusRestore) this.focus();
      },
      onCancel: () => {
        this.suppressFocusRestoreOnNextCommandCommit = false;
        this.updateEditState();
        this.renderSelection();
        this.updateStatus();
        this.focus();
      }
    });

    // Excel behavior: leaving in-cell editing (e.g. clicking another cell, ribbon, etc)
    // should commit the draft text.
    //
    // IMPORTANT: Avoid stealing focus back from whatever surface the user clicked. When the
    // editor blurs to an element other than the grid root itself, suppress the focus-restore
    // logic that normally runs after a command commit.
    const onEditorBlur = (event: FocusEvent) => {
      if (!this.editor.isOpen()) return;
      const next = event.relatedTarget as Node | null;
      // Only restore focus to the grid when the blur target is the grid root itself (e.g.
      // DesktopSharedGrid focusing the container during pointer interactions). If focus moved
      // to any other element (including focusable overlays inside the grid root), don't steal
      // it back.
      this.suppressFocusRestoreOnNextCommandCommit = next !== this.root;
      this.editor.commit("command");
    };
    this.editor.element.addEventListener("blur", onEditorBlur, { signal: this.domAbort.signal });

    this.inlineEditController = new InlineEditController({
      container: this.root,
      document: this.document,
      workbookId: opts.workbookId,
      schemaProvider: createSchemaProviderFromSearchWorkbook(this.searchWorkbook),
      getSheetId: () => this.sheetId,
      sheetNameResolver: this.sheetNameResolver,
      getSelectionRange: () => this.getInlineEditSelectionRange(),
      onApplied: () => {
        this.renderGrid();
        this.renderCharts(false);
        this.renderSelection();
        this.updateStatus();
        this.focus();
      },
      onClosed: () => {
        this.focus();
        this.updateEditState();
      },
      llmClient: opts.inlineEdit?.llmClient,
      model: opts.inlineEdit?.model,
      auditStore: opts.inlineEdit?.auditStore,
      onWorkbookContextBuildStats: opts.inlineEdit?.onWorkbookContextBuildStats,
    });

    if (this.gridMode === "shared") {
      const headerRows = 1;
      const headerCols = 1;
      this.sharedProvider = new DocumentCellProvider({
        document: this.document,
        getSheetId: () => this.sheetId,
        headerRows,
        headerCols,
        rowCount: this.limits.maxRows + headerRows,
        colCount: this.limits.maxCols + headerCols,
        showFormulas: () => this.showFormulas,
        getComputedValue: (cell) => this.getCellComputedValue(cell),
        getCommentMeta: (row, col) => this.commentMetaByCoord.get(row * COMMENT_COORD_COL_STRIDE + col) ?? null
      });

      this.sharedGrid = new DesktopSharedGrid({
        container: this.root,
        provider: this.sharedProvider,
        rowCount: this.limits.maxRows + headerRows,
        colCount: this.limits.maxCols + headerCols,
        frozenRows: headerRows,
        frozenCols: headerCols,
        defaultRowHeight: this.cellHeight,
        defaultColWidth: this.cellWidth,
        imageResolver: this.sharedGridImageResolver,
        enableResize: true,
        enableKeyboard: false,
        canvases: { grid: this.gridCanvas, content: this.referenceCanvas, selection: this.selectionCanvas },
        scrollbars: {
          vTrack: this.vScrollbarTrack,
          vThumb: this.vScrollbarThumb,
          hTrack: this.hScrollbarTrack,
          hThumb: this.hScrollbarThumb
        },
        callbacks: {
          onScroll: (scroll, viewport) => {
            // DesktopSharedGrid batches viewport change notifications through `requestAnimationFrame`.
            // In unit tests, `requestAnimationFrame` is often stubbed to fire synchronously, which
            // means we can receive `onScroll` callbacks during SpreadsheetApp construction (before
            // overlays like `drawingOverlay` are initialized). Defer overlay rendering until the
            // app is fully mounted (`uiReady=true`), but still keep scroll state in sync.
            if (!this.uiReady) {
              this.scrollX = scroll.x;
              this.scrollY = scroll.y;
              return;
            }

            let effectiveViewport = viewport;
            const prevZoom = this.sharedGridZoom;
            const nextZoom = this.sharedGrid?.renderer.getZoom() ?? prevZoom;

            const zoomChanged = nextZoom !== prevZoom;
            if (zoomChanged) {
              this.sharedGridZoom = nextZoom;
              // `CanvasGridRenderer.setZoom()` scales both the default row/col sizes and any
              // existing overrides derived from document state. Avoid re-applying persisted
              // overrides here: rebuilding large override maps on every zoom gesture step can
              // be expensive when many explicit row/col sizes exist.
              this.dispatchZoomChanged();
              this.notifyZoomListeners();
              effectiveViewport = this.sharedGrid?.renderer.getViewportState() ?? viewport;
            }

            const prevX = this.scrollX;
            const prevY = this.scrollY;
            const nextScroll = zoomChanged ? (this.sharedGrid?.renderer.scroll.getScroll() ?? scroll) : scroll;
            this.scrollX = nextScroll.x;
            this.scrollY = nextScroll.y;
            this.clearSharedHoverCellCache();
            this.hideCommentTooltip();
            this.renderDrawings(effectiveViewport);
            if (!this.useCanvasCharts) {
              this.renderCharts(false);
            }
            this.renderAuditing();
            this.renderSelection();
            if (this.scrollX !== prevX || this.scrollY !== prevY) {
              this.notifyScrollListeners();
            }
          },
          onSelectionChange: () => {
            if (this.sharedGridSelectionSyncInProgress) return;
            this.syncSelectionFromSharedGrid();
            this.updateStatus();
          },
          onSelectionRangeChange: () => {
            if (this.sharedGridSelectionSyncInProgress) return;
            this.syncSelectionFromSharedGrid();
            this.updateStatus();
          },
          onRequestCellEdit: (request) => {
            if (this.isReadOnly()) return;
            this.openEditorFromSharedGrid(request);
          },
          onAxisSizeChange: (change) => {
            this.onSharedGridAxisSizeChange(change);
          },
          onRangeSelectionStart: (range) => this.onSharedRangeSelectionStart(range),
          onRangeSelectionChange: (range) => this.onSharedRangeSelectionChange(range),
          onRangeSelectionEnd: () => this.onSharedRangeSelectionEnd(),
          onFillCommit: ({ sourceRange, targetRange, mode }) => {
            // Fill operations should never mutate the sheet while the user is actively editing text
            // (cell editor, formula bar, inline edit). This mirrors the keyboard shortcut guards.
            //
            // Note: DesktopSharedGrid will still expand the selection to the dragged target range
            // after this callback runs. Revert the selection on the next microtask so the UI
            // reflects that no fill occurred.
            if (this.isReadOnly() || this.isEditing()) {
              const selectionSnapshot = {
                ranges: this.selection.ranges.map((r) => ({ ...r })),
                active: { ...this.selection.active },
                anchor: { ...this.selection.anchor },
                activeRangeIndex: this.selection.activeRangeIndex
              };
              queueMicrotask(() => {
                if (!this.sharedGrid) return;
                if (this.disposed) return;
                this.selection = buildSelection(selectionSnapshot, this.limits);
                this.syncSharedGridSelectionFromState({ scrollIntoView: false });
                this.renderSelection();
                this.updateStatus();
                if (this.formulaBar?.isEditing() || this.formulaEditCell) {
                  this.formulaBar?.focus();
                } else {
                  this.focus();
                }
              });
              return;
            }
            const headerRows = this.sharedHeaderRows();
            const headerCols = this.sharedHeaderCols();

            const toFillRange = (range: GridCellRange): FillEngineRange | null => {
              const startRow = Math.max(0, range.startRow - headerRows);
              const endRow = Math.max(0, range.endRow - headerRows);
              const startCol = Math.max(0, range.startCol - headerCols);
              const endCol = Math.max(0, range.endCol - headerCols);
              if (endRow <= startRow || endCol <= startCol) return null;
              return { startRow, endRow, startCol, endCol };
            };

            const source = toFillRange(sourceRange);
            const target = toFillRange(targetRange);
            if (!source || !target) return;

            const sourceCells = (source.endRow - source.startRow) * (source.endCol - source.startCol);
            const targetCells = (target.endRow - target.startRow) * (target.endCol - target.startCol);
            if (sourceCells > MAX_FILL_CELLS || targetCells > MAX_FILL_CELLS) {
              try {
                showToast(
                  `Fill range too large (>${MAX_FILL_CELLS.toLocaleString()} cells). Select fewer cells and try again.`,
                  "warning"
                );
              } catch {
                // `showToast` requires a #toast-root; unit tests don't always include it.
              }

              // DesktopSharedGrid will still expand the selection to the dragged target range
              // after this callback runs. Revert selection back to the pre-fill state on the
              // next microtask turn so the UI reflects that no fill occurred.
              const selectionSnapshot = {
                ranges: this.selection.ranges.map((r) => ({ ...r })),
                active: { ...this.selection.active },
                anchor: { ...this.selection.anchor },
                activeRangeIndex: this.selection.activeRangeIndex
              };
              queueMicrotask(() => {
                if (!this.sharedGrid) return;
                if (this.disposed) return;
                this.selection = buildSelection(selectionSnapshot, this.limits);
                this.syncSharedGridSelectionFromState({ scrollIntoView: false });
                this.renderSelection();
                this.updateStatus();
                this.focus();
              });
              return;
            }

            const fillCoordScratch = { row: 0, col: 0 };
            const getCellComputedValue = (row: number, col: number) => {
              fillCoordScratch.row = row;
              fillCoordScratch.col = col;
              return this.getCellComputedValue(fillCoordScratch) as any;
            };

            // Prefer engine-backed formula shifting when available (handles A:A / 1:1 / ranges, etc).
            const wasm = this.wasmEngine;
            if (wasm && mode !== "copy") {
              const task = applyFillCommitToDocumentControllerWithFormulaRewrite({
                document: this.document,
                sheetId: this.sheetId,
                sourceRange: source,
                targetRange: target,
                mode,
                getCellComputedValue,
                rewriteFormulasForCopyDelta: (requests) => wasm.rewriteFormulasForCopyDelta(requests),
              })
                .catch(() => {
                  // Fall back to the legacy best-effort fill engine if the worker is unavailable.
                  applyFillCommitToDocumentController({
                    document: this.document,
                    sheetId: this.sheetId,
                    sourceRange: source,
                    targetRange: target,
                    mode,
                    getCellComputedValue,
                  });
                })
                .finally(() => {
                  // Ensure non-grid overlays (charts, auditing) refresh after the mutation.
                  this.refresh();
                  this.focus();
                });
              this.idle.track(task);
              return;
            }

            applyFillCommitToDocumentController({
              document: this.document,
              sheetId: this.sheetId,
              sourceRange: source,
              targetRange: target,
              mode,
              getCellComputedValue,
            });

            // Ensure non-grid overlays (charts, auditing) refresh after the mutation.
            this.refresh();
            this.focus();
          }
        }
      });

    }

    const persistedDrawingImages = new IndexedDbImageStore(localWorkbookId);

    // Primary image store used by the drawings overlay. Writes should be undoable (user edits).
    this.drawingImages = new DocumentImageStore(this.document, persistedDrawingImages, { mode: "user" });
    if (this.collabSession) {
      // In collab mode, image bytes are not guaranteed to be available locally (e.g. remote inserts,
      // or inserts made on another device without offline persistence). Bind drawing images to Yjs
      // metadata so collaborators eventually converge on the actual bytes.
      this.imageBytesBinder = bindImageBytesToCollabSession({
        session: this.collabSession,
        // Use a dedicated adapter for hydration so remote bytes don't create local undo history.
        images: new DocumentImageStore(this.document, persistedDrawingImages, { mode: "external", source: "collab" }),
        origin: this.collabBinderOrigin ?? undefined,
      });
    }

    const legacyDrawingGeom: DrawingGridGeometry = {
      cellOriginPx: (cell) => ({
        x: this.visualIndexForCol(cell.col) * this.cellWidth,
        y: this.visualIndexForRow(cell.row) * this.cellHeight,
      }),
      cellSizePx: () => ({ width: this.cellWidth, height: this.cellHeight }),
    };

    const sharedDrawingGeom: DrawingGridGeometry = {
      cellOriginPx: (cell) => {
        const grid = this.sharedGrid;
        if (!grid) return { x: 0, y: 0 };
        const headerRows = this.sharedHeaderRows();
        const headerCols = this.sharedHeaderCols();
        const headerWidth = headerCols > 0 ? grid.renderer.scroll.cols.totalSize(headerCols) : 0;
        const headerHeight = headerRows > 0 ? grid.renderer.scroll.rows.totalSize(headerRows) : 0;
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return {
          x: grid.renderer.scroll.cols.positionOf(gridCol) - headerWidth,
          y: grid.renderer.scroll.rows.positionOf(gridRow) - headerHeight,
        };
      },
      cellSizePx: (cell) => {
        const grid = this.sharedGrid;
        if (!grid) return { width: this.cellWidth, height: this.cellHeight };
        const headerRows = this.sharedHeaderRows();
        const headerCols = this.sharedHeaderCols();
        const gridRow = cell.row + headerRows;
        const gridCol = cell.col + headerCols;
        return { width: grid.renderer.getColWidth(gridCol), height: grid.renderer.getRowHeight(gridRow) };
      },
    };

    this.drawingGeom = this.gridMode === "shared" ? sharedDrawingGeom : legacyDrawingGeom;
    this.drawingChartRenderer = new ChartRendererAdapter(this.createDrawingChartRendererStore());
    this.drawingOverlay = new DrawingOverlay(
      this.drawingCanvas,
      this.drawingImages,
      this.drawingGeom,
      this.drawingChartRenderer,
    );
    this.drawingOverlay.setSelectedId(null);

    const enableDrawingInteractions = opts.enableDrawingInteractions ?? this.drawingsDemoEnabled;
    if (enableDrawingInteractions) {
      const callbacks: DrawingInteractionCallbacks = {
        getViewport: () => this.getDrawingInteractionViewport(this.sharedGrid?.renderer.scroll.getViewportState()),
        getObjects: () => this.listDrawingObjectsForSheet(),
        setObjects: (next) => {
          this.setDrawingObjectsForSheet(next);
          this.scheduleDrawingsRender();
          const selected = this.selectedDrawingId != null ? next.find((obj) => obj.id === this.selectedDrawingId) : undefined;
          if (selected?.kind.type === "chart") {
            // Best-effort: keep chart overlays aligned when moving/resizing chart drawings.
            this.renderCharts(false);
          }
        },
        commitObjects: (next) => {
          this.document.setSheetDrawings(this.sheetId, next, { source: "drawings" });
        },
        beginBatch: ({ label }) => this.document.beginBatch({ label }),
        endBatch: () => this.document.endBatch(),
        cancelBatch: () => this.document.cancelBatch(),
        shouldHandlePointerDown: () => !this.formulaBar?.isFormulaEditing(),
        onPointerDownHit: () => {
          if (this.editor.isOpen()) {
            this.editor.commit("command");
          }
        },
        onSelectionChange: (selectedId) => {
          const prev = this.selectedDrawingId;
          this.selectedDrawingId = selectedId;
          if (prev !== selectedId) {
            this.dispatchDrawingSelectionChanged();
          }
          // Drawings and charts are mutually exclusive selections; selecting a drawing
          // should clear any active chart selection so selection handles don't double-render.
          if (selectedId != null && this.selectedChartId != null) {
            this.setSelectedChartId(null);
          }
          this.drawingOverlay.setSelectedId(selectedId);
          this.renderDrawings();
        },
        requestFocus: () => this.focus(),
      };
      this.drawingInteractionCallbacks = callbacks;
      const interactionElement = this.gridMode === "shared" ? this.selectionCanvas : this.root;
      this.drawingInteractionController = new DrawingInteractionController(interactionElement, this.drawingGeom, callbacks, {
        capture: this.gridMode === "shared",
      });
    }

    if (this.sharedGrid) {
      // Match the legacy header sizing so existing click offsets and overlays stay aligned.
      //
      // Important: set these after `this.drawingOverlay` is constructed since the shared-grid renderer
      // emits viewport-layout change notifications (via rAF) when axis sizes update.
      this.sharedGrid.renderer.setColWidth(0, this.rowHeaderWidth);
      this.sharedGrid.renderer.setRowHeight(0, this.colHeaderHeight);
      this.sharedGridZoom = this.sharedGrid.renderer.getZoom();
    }

    if (this.gridMode === "shared") {
      // Shared-grid mode uses the CanvasGridRenderer selection layer, but we still
      // need pointer movement for comment tooltips.
      this.root.addEventListener("pointermove", (e) => this.onSharedPointerMove(e), {
        passive: true,
        signal: this.domAbort.signal,
      });
      this.root.addEventListener(
        "pointerenter",
        () => this.maybeRefreshRootPosition({ force: true }),
        { passive: true, signal: this.domAbort.signal }
      );
      this.root.addEventListener(
        "pointerleave",
        () => {
          this.clearSharedHoverCellCache();
          this.hideCommentTooltip();
          this.root.style.cursor = "";
        },
        { signal: this.domAbort.signal }
      );
      this.root.addEventListener("keydown", (e) => this.onKeyDown(e), { signal: this.domAbort.signal });
    } else {
      this.root.addEventListener("pointerdown", (e) => this.onPointerDown(e), { signal: this.domAbort.signal });
      this.root.addEventListener("pointermove", (e) => this.onPointerMove(e), {
        passive: true,
        signal: this.domAbort.signal,
      });
      this.root.addEventListener("pointerup", (e) => this.onPointerUp(e), { passive: true, signal: this.domAbort.signal });
      this.root.addEventListener("pointercancel", (e) => this.onPointerUp(e), {
        passive: true,
        signal: this.domAbort.signal,
      });
      this.root.addEventListener(
        "pointerenter",
        () => this.maybeRefreshRootPosition({ force: true }),
        { passive: true, signal: this.domAbort.signal }
      );
      this.root.addEventListener(
        "pointerleave",
        () => {
          this.hideCommentTooltip();
          this.root.style.cursor = "";
        },
        { signal: this.domAbort.signal }
      );
      this.root.addEventListener("keydown", (e) => this.onKeyDown(e), { signal: this.domAbort.signal });
      this.root.addEventListener("wheel", (e) => this.onWheel(e), { passive: false, signal: this.domAbort.signal });

      this.vScrollbarThumb.addEventListener("pointerdown", (e) => this.onScrollbarThumbPointerDown(e, "y"), {
        passive: false,
        signal: this.domAbort.signal
      });
      this.hScrollbarThumb.addEventListener("pointerdown", (e) => this.onScrollbarThumbPointerDown(e, "x"), {
        passive: false,
        signal: this.domAbort.signal
      });
      this.vScrollbarTrack.addEventListener("pointerdown", (e) => this.onScrollbarTrackPointerDown(e, "y"), {
        passive: false,
        signal: this.domAbort.signal
      });
      this.hScrollbarTrack.addEventListener("pointerdown", (e) => this.onScrollbarTrackPointerDown(e, "x"), {
        passive: false,
        signal: this.domAbort.signal
      });
    }

    if (this.useCanvasCharts) {
      const geom = this.chartOverlayGeom;
      if (geom) {
        const anchorsEqual = (a: DrawingObject["anchor"], b: DrawingObject["anchor"]): boolean => {
          if (a.type !== b.type) return false;
          if (a.type === "absolute") {
            return a.pos.xEmu === (b as any).pos.xEmu && a.pos.yEmu === (b as any).pos.yEmu && a.size.cx === (b as any).size.cx && a.size.cy === (b as any).size.cy;
          }
          if (a.type === "oneCell") {
            const bb = b as any;
            return (
              a.from.cell.row === bb.from.cell.row &&
              a.from.cell.col === bb.from.cell.col &&
              a.from.offset.xEmu === bb.from.offset.xEmu &&
              a.from.offset.yEmu === bb.from.offset.yEmu &&
              a.size.cx === bb.size.cx &&
              a.size.cy === bb.size.cy
            );
          }
          // twoCell
          const bb = b as any;
          return (
            a.from.cell.row === bb.from.cell.row &&
            a.from.cell.col === bb.from.cell.col &&
            a.from.offset.xEmu === bb.from.offset.xEmu &&
            a.from.offset.yEmu === bb.from.offset.yEmu &&
            a.to.cell.row === bb.to.cell.row &&
            a.to.cell.col === bb.to.cell.col &&
            a.to.offset.xEmu === bb.to.offset.xEmu &&
            a.to.offset.yEmu === bb.to.offset.yEmu
          );
        };

        this.chartDrawingInteraction = new DrawingInteractionController(
          this.root,
          geom,
          {
            getViewport: () => this.getDrawingInteractionViewport(),
            getObjects: () => this.listCanvasChartDrawingObjectsForSheet(this.sheetId, 0),
            setObjects: (next) => {
              const current = this.listCanvasChartDrawingObjectsForSheet(this.sheetId, 0);
              const prevById = new Map<number, DrawingObject>();
              for (const obj of current) prevById.set(obj.id, obj);

              for (const obj of next) {
                const prev = prevById.get(obj.id);
                if (!prev) continue;
                if (anchorsEqual(prev.anchor, obj.anchor)) continue;
                const chartId = obj.kind.type === "chart" ? obj.kind.chartId : undefined;
                if (typeof chartId !== "string" || chartId.trim() === "") continue;
                this.chartStore.updateChartAnchor(chartId, drawingAnchorToChartAnchor(obj.anchor));
              }
            },
            onSelectionChange: (selectedId) => {
              const selected =
                selectedId != null ? this.listCanvasChartDrawingObjectsForSheet(this.sheetId, 0).find((o) => o.id === selectedId) : null;
              const nextChartId =
                selected?.kind.type === "chart" && typeof selected.kind.chartId === "string" ? selected.kind.chartId : null;

              // Drawings and charts are mutually exclusive selections. Selecting a chart
              // should clear any drawing selection so selection handles don't double-render.
              if (nextChartId != null && this.selectedDrawingId != null) {
                this.selectedDrawingId = null;
                this.drawingInteractionController?.setSelectedId(null);
                this.dispatchDrawingSelectionChanged();
              }

              this.selectedChartId = nextChartId;
              this.renderDrawings();
            },
          },
          { capture: true },
        );
      }
    } else {
      // Chart interactions use hit testing against chart bounds rather than enabling pointer
      // events on the chart DOM itself (charts remain `pointer-events: none` so grid scrolling
      // and selection behave consistently). Use a capture listener so we can intercept
      // chart pointerdowns before grid selection handlers run.
      this.root.addEventListener("pointerdown", (e) => this.onChartPointerDownCapture(e), {
        capture: true,
        passive: false,
        signal: this.domAbort.signal,
      });
    }

    // Drawings interactions also require capture-based hit testing because in shared-grid mode
    // pointer events are handled by the full-size `selectionCanvas` (which includes headers)
    // while drawings are rendered onto `drawingCanvas` (which spans the full grid root and is
    // clipped to the cell grid body area under headers).
    //
    // Use a capture listener so we can intercept drawing clicks before grid selection handlers run.
    this.root.addEventListener("pointerdown", (e) => this.onDrawingPointerDownCapture(e), {
      capture: true,
      passive: false,
      signal: this.domAbort.signal,
    });

    this.root.addEventListener("dragover", (e) => this.onGridDragOver(e), { signal: this.domAbort.signal });
    this.root.addEventListener("drop", (e) => this.onGridDrop(e), { signal: this.domAbort.signal });

    // If the user copies/cuts from an input/contenteditable (formula bar, comments, etc),
    // the system clipboard content changes and any prior "internal copy" context used for
    // style/formula shifting should be discarded. Grid copy/cut uses `preventDefault()`
    // and the Clipboard API, so it should not trigger these native events.
    if (typeof document !== "undefined") {
      const clearClipboardContext = () => {
        this.clipboardCopyContext = null;
      };
      document.addEventListener("copy", clearClipboardContext, { capture: true, signal: this.domAbort.signal });
      document.addEventListener("cut", clearClipboardContext, { capture: true, signal: this.domAbort.signal });

      if (typeof window !== "undefined") {
        window.addEventListener("blur", clearClipboardContext, { signal: this.domAbort.signal });
      }
      document.addEventListener(
        "visibilitychange",
        () => {
          if ((document as any).hidden) clearClipboardContext();
        },
        { signal: this.domAbort.signal }
      );
    }

    this.resizeObserver = new ResizeObserver(() => this.onResize());
    this.resizeObserver.observe(this.root);

    const emitCommentsChanged = () => this.dispatchCommentsChanged();

    if (!collabEnabled) {
      // Save so we can detach cleanly in `destroy()`.
      this.commentsDocUpdateListener = () => {
        this.reindexCommentCells();
        this.refresh();
        emitCommentsChanged();
      };
      this.commentsDoc.on("update", this.commentsDocUpdateListener);
    } else {
      // Collab mode: comments live inside the shared workbook Y.Doc. Avoid listening to
      // `doc.on("update")` (which would fire for every cell edit); instead, observe just
      // the comments root once the provider has hydrated the doc.
      const session = this.collabSession;
      if (session) {
        const provider = session.provider;
        const attach = () => {
          if (this.disposed) return;
          if (this.stopCommentsRootObserver) return;
          let root: ReturnType<typeof getCommentsRoot> | null = null;
          try {
            root = getCommentsRoot(this.commentsDoc);
          } catch {
            // Best-effort; never block app startup on comment schema issues.
          }
          if (!root) return;

          // Ensure comments participate in collaborative undo/redo. We defer this
          // until after provider sync so we don't accidentally clobber legacy
          // Array-backed roots by instantiating a Map too early.
          this.ensureCommentsUndoScope(root);

          const handler = () => {
            this.reindexCommentCells();
            this.refresh();
            emitCommentsChanged();
          };

          if (root.kind === "map") {
            root.map.observeDeep(handler);
            this.stopCommentsRootObserver = () => root?.kind === "map" && root.map.unobserveDeep(handler);
          } else {
            root.array.observeDeep(handler);
            this.stopCommentsRootObserver = () => root?.kind === "array" && root.array.unobserveDeep(handler);
          }

          // Initial index after hydration.
          this.reindexCommentCells();
          this.refresh();
          emitCommentsChanged();
        };

        if (provider && typeof provider.on === "function") {
          const onSync = (isSynced: boolean) => {
            if (!isSynced) return;
            provider.off?.("sync", onSync);
            attach();
          };
          provider.on("sync", onSync);
          if ((provider as any).synced) onSync(true);
          // If the user creates a comment before the provider reports `sync=true`,
          // the `comments` root will already exist locally. Attach the comments
          // observer immediately in that case so the UI updates in real time.
          //
          // This remains safe for legacy Array-backed docs because we only attach
          // once the root exists (and `getCommentsRoot` peeks at the underlying
          // placeholder before choosing a constructor).
          try {
            if (this.commentsDoc.share.get("comments")) attach();
          } catch {
            // Best-effort.
          }

          // When local persistence is enabled, the persisted doc state may be loaded
          // before the WebSocket provider emits `sync=true`. If comments are present
          // in that persisted state, attach the comments observer as soon as local
          // persistence hydration completes so comment indicators/panels are populated
          // immediately (even if the provider never reports synced, e.g. offline).
          if (typeof (session as any).whenLocalPersistenceLoaded === "function") {
            void session
              .whenLocalPersistenceLoaded()
              .then(() => {
                try {
                  if (this.commentsDoc.share.get("comments")) attach();
                } catch {
                  // Best-effort.
                }
              })
              .catch(() => {
                // ignore
              });
          }
        } else {
          attach();
        }
      }
    }

    this.auditingUnsubscribe = this.document.on("change", (payload: any) => {
      // Outline state (row/col grouping + hidden flags) is tracked locally per sheet.
      // Ensure we don't retain outline state for sheets that have been deleted from the document.
      const sheetMetaDeltas = Array.isArray(payload?.sheetMetaDeltas) ? payload.sheetMetaDeltas : [];
      for (const delta of sheetMetaDeltas) {
        const sheetId = delta?.sheetId;
        if (typeof sheetId !== "string" || sheetId === "") continue;
        if (delta?.after == null) {
          this.outlinesBySheet.delete(sheetId);
        }
      }
      if (payload?.source === "applyState") {
        // `DocumentController.applyState` can delete sheets without emitting sheetMetaDeltas.
        // Reconcile the outline map against the current sheet ids to avoid leaking stale state.
        const existing = new Set(this.document.getSheetIds());
        for (const key of this.outlinesBySheet.keys()) {
          if (!existing.has(key)) this.outlinesBySheet.delete(key);
        }
      }

      this.auditingCache.clear();
      this.auditingLastCellKey = null;
      if (this.auditingMode !== "off") {
        this.scheduleAuditingUpdate();
      }

      // Track which charts' underlying data ranges were touched so off-screen charts can
      // refresh their cached series data when scrolled into view.
      this.markChartsDirtyFromDeltas(payload?.deltas);

      // When formulas may have been recalculated, the chart ranges can change even if the
      // triggering cell deltas were outside the chart's direct ranges (e.g. a chart plots
      // `B1` where `B1` is `=A1*2`, and the user edits `A1`).
      //
      // If we're not relying on the WASM engine's computed-value deltas, conservatively mark
      // charts that contain formulas as dirty so their cached data refreshes on the next
      // render.
      if (payload?.recalc) {
        const sheetCount = (this.document as any)?.model?.sheets?.size;
        const useEngineCache =
          (typeof sheetCount === "number" ? sheetCount : this.document.getSheetIds().length) <= 1;
        const hasWasmEngine = Boolean(this.wasmEngine && !this.wasmSyncSuspended);
        if (!hasWasmEngine || !useEngineCache) {
          this.markFormulaChartsDirty();
        }
      }

      // DocumentController changes can also include sheet-level view deltas
      // (e.g. frozen panes). In shared-grid mode, frozen panes must be pushed
      // down to the CanvasGridRenderer explicitly.
      //
      // Avoid re-syncing axis sizes back into the *same* shared-grid renderer after local
      // resize/auto-fit interactions. Those interactions update the renderer directly during
      // the drag, and rebuilding large override maps here can be expensive for sheets with
      // many explicit row/col sizes.
      const source = typeof payload?.source === "string" ? payload.source : "";
      const hasActiveSheetViewDelta =
        Array.isArray(payload?.sheetViewDeltas) &&
        payload.sheetViewDeltas.some((delta: any) => delta?.sheetId === this.sheetId);
      if (hasActiveSheetViewDelta) {
        if (source !== "sharedGridAxis") {
          this.syncFrozenPanes();
        }
        // Background image selection is persisted in the document's sheet view state;
        // keep the active sheet's rendered pattern aligned with the latest view deltas
        // (including undo/redo and collaboration updates).
        this.syncActiveSheetBackgroundImage();
      }
    });

    // SpreadsheetApp's legacy (non-shared) renderer only repaints when explicitly asked.
    // Remote/collaboration updates arrive via `DocumentController.applyExternalDeltas()` and
    // emit `document.on("change")`, but do not necessarily go through any local UI action
    // that would call `refresh()`. Ensure we schedule a repaint for externally-sourced
    // deltas so collaboration never appears "stuck" until the next scroll/input event.
    this.externalRepaintUnsubscribe = this.document.on("change", (payload: any) => {
      const source = typeof payload?.source === "string" ? payload.source : "";
      const isExternalSource =
        source === "collab" ||
        source === "backend" ||
        source === "python" ||
        source === "macro" ||
        source === "pivot" ||
        source === "extension" ||
        source === "sheetRename" ||
        source === "sheetDelete" ||
        source === "applyState";
      if (!isExternalSource) return;
      // Shared-grid mode repaints via CanvasGridRenderer/provider invalidations; avoid scheduling
      // redundant SpreadsheetApp refreshes (which mainly exist for the legacy renderer).
      if (!this.sharedGrid) {
        this.refresh("scroll");
      }
      // External edits can affect the active cell's displayed value (e.g. direct edits to the active
      // cell, edits inside the current selection affecting summary stats, or formula dependency
      // changes). Schedule a debounced status/formula bar update when it is likely to matter.
      const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
      if (deltas.length > 0) {
        const active = this.selection?.active ?? null;
        const touchesActive =
          active != null &&
          deltas.some(
            (d: any) =>
              d &&
              String(d.sheetId ?? "") === this.sheetId &&
              Number(d.row) === active.row &&
              Number(d.col) === active.col
          );

        const activeState =
          active != null ? (this.document.getCell(this.sheetId, active) as { formula: string | null } | null) : null;
        const activeIsFormula = activeState?.formula != null;

        const wantsSelectionStats = Boolean(this.status.selectionSum || this.status.selectionAverage || this.status.selectionCount);

        const touchesSelection = wantsSelectionStats
          ? deltas.some((d: any) => {
              if (!d) return false;
              if (String(d.sheetId ?? "") !== this.sheetId) return false;
              const row = Number(d.row);
              const col = Number(d.col);
              if (!Number.isInteger(row) || row < 0) return false;
              if (!Number.isInteger(col) || col < 0) return false;
              return this.selection?.ranges?.some((r) => {
                const startRow = Math.min(r.startRow, r.endRow);
                const endRow = Math.max(r.startRow, r.endRow);
                const startCol = Math.min(r.startCol, r.endCol);
                const endCol = Math.max(r.startCol, r.endCol);
                return row >= startRow && row <= endRow && col >= startCol && col <= endCol;
              });
            })
          : false;

        if (touchesActive || activeIsFormula || touchesSelection) {
          this.scheduleStatusUpdate();
        }
      }
      // Similarly, chart data caches are refreshed lazily via `dirtyChartIds`. Schedule a debounced
      // chart redraw so visible charts reflect remote data edits in real time.
      this.scheduleChartContentRefresh(payload);
    });

    // Drawings/images may update via document-level deltas (Task 148+) without going
    // through SpreadsheetApp UI actions that call `refresh()`. Listen for drawing/image
    // change payloads and re-render the active sheet's drawing overlay.
    const drawingUnsubs: Array<() => void> = [];
    const invalidateAndRenderDrawings = (reason?: string) => {
      // Keep memory bounded: only cache the active sheet's objects.
      this.drawingObjectsCache = null;
      this.drawingHitTestIndex = null;
      this.drawingHitTestIndexObjects = null;
      this.scheduleDrawingsRender(reason);
    };

    drawingUnsubs.push(
      this.document.on("change", (payload: any) => {
        if (this.disposed) return;
        if (!this.uiReady) return;
        if (!this.documentChangeAffectsDrawings(payload)) return;
        const source = typeof payload?.source === "string" ? payload.source : "";
        if (source === "applyState") {
          this.drawingOverlay.clearImageCache();
        }
          const imageDeltas: any[] = Array.isArray(payload?.imageDeltas)
          ? payload.imageDeltas
          : Array.isArray(payload?.imagesDeltas)
            ? payload.imagesDeltas
            : [];
        const activeDesiredBackgroundId = this.getSheetBackgroundImageId(this.sheetId) ?? null;
        let activeBackgroundNeedsReload = false;

        for (const delta of imageDeltas) {
          const imageId = typeof delta?.imageId === "string" ? delta.imageId : typeof delta?.id === "string" ? delta.id : null;
          if (!imageId) continue;
          this.drawingOverlay.invalidateImage(imageId);

          // Keep the workbook-scoped in-cell/background image store aligned with DocumentController
          // image deltas so callers can populate images via `DocumentController.setImage(...)`.
          //
          // Note: This is intentionally best-effort; ignore malformed deltas.
          try {
            const after = (delta as any)?.after ?? null;
            if (!after) {
              this.imageStore.delete(imageId);
            } else if (typeof after === "object") {
              const bytes: unknown = (after as any).bytes;
              if (bytes instanceof Uint8Array) {
                const mimeTypeRaw: unknown = (after as any).mimeType;
                const mimeType =
                  typeof mimeTypeRaw === "string" && mimeTypeRaw.trim() !== "" ? mimeTypeRaw : "application/octet-stream";
                this.imageStore.set(imageId, { bytes, mimeType });
              }
            }

            // Image bytes may have changed; invalidate decoded bitmaps so patterns re-decode.
            this.workbookImageBitmaps.invalidate(imageId);

            if (imageId === activeDesiredBackgroundId) {
              activeBackgroundNeedsReload = true;
            }
          } catch {
            // ignore
          }
        }

        if (activeBackgroundNeedsReload) {
          // Force a reload even when the active sheet points at the same image id.
          this.activeSheetBackgroundAbort?.abort();
          this.activeSheetBackgroundAbort = null;
          this.activeSheetBackgroundImageId = null;
          this.activeSheetBackgroundBitmap = null;
          this.syncActiveSheetBackgroundImage();
        }
        this.handleWorkbookImageDeltasForBackground(payload);
        invalidateAndRenderDrawings("document:change");
      }),
    );

    // Optional dedicated event streams (if/when DocumentController adds them).
    drawingUnsubs.push(
      this.document.on("drawings", (payload: any) => {
        if (this.disposed) return;
        if (!this.uiReady) return;
        const sheetId = typeof payload?.sheetId === "string" ? payload.sheetId : null;
        if (sheetId && sheetId !== this.sheetId) return;
        invalidateAndRenderDrawings("document:drawings");
      }),
    );
    drawingUnsubs.push(
      this.document.on("images", () => {
        if (this.disposed) return;
        if (!this.uiReady) return;
        this.drawingOverlay.clearImageCache();
        invalidateAndRenderDrawings("document:images");
      }),
    );

    this.drawingsUnsubscribe = () => {
      for (const unsub of drawingUnsubs) unsub();
    };

    if (!collabEnabled && typeof window !== "undefined") {
      try {
        this.stopCommentPersistence = bindDocToStorage(this.commentsDoc, window.localStorage, "formula:comments");
      } catch {
        // Ignore persistence failures (e.g. storage disabled).
      }
    }

    if (opts.formulaBar) {
      const nameBoxDropdownProvider: NameBoxDropdownProvider = {
        getItems: () => {
          const items: ReturnType<NameBoxDropdownProvider["getItems"]> = [];
          const seen = new Set<string>();
          const push = (item: ReturnType<NameBoxDropdownProvider["getItems"]>[number]): void => {
            if (seen.has(item.key)) return;
            seen.add(item.key);
            items.push(item);
          };

          const formatSheetPrefix = (token: string): string => {
            const formatted = formatSheetNameForA1(token);
            return formatted ? `${formatted}!` : "";
          };
          const normalizeDocRange = (range: any): Range | null => {
            if (!range) return null;
            const { startRow, endRow, startCol, endCol } = range as any;
            if (
              !Number.isInteger(startRow) ||
              !Number.isInteger(endRow) ||
              !Number.isInteger(startCol) ||
              !Number.isInteger(endCol) ||
              startRow < 0 ||
              endRow < 0 ||
              startCol < 0 ||
              endCol < 0
            ) {
              return null;
            }

            return {
              startRow: Math.min(startRow, endRow),
              endRow: Math.max(startRow, endRow),
              startCol: Math.min(startCol, endCol),
              endCol: Math.max(startCol, endCol),
            };
          };

          for (const entry of this.searchWorkbook.names.values()) {
            const e: any = entry as any;
            const name = typeof e?.name === "string" ? String(e.name).trim() : "";
            if (!name) continue;
            const sheetName = typeof e?.sheetName === "string" ? String(e.sheetName) : "";
            const range = normalizeDocRange(e?.range);
            const description = range
              ? sheetName
                ? `${formatSheetPrefix(sheetName)}${rangeToA1(range)}`
                : rangeToA1(range)
              : undefined;
            push({
              kind: "namedRange",
              key: `namedRange:${name}`,
              label: name,
              reference: range ? name : "",
              description,
            });
          }

          for (const tableEntry of this.searchWorkbook.tables.values()) {
            const t: any = tableEntry as any;
            const name = typeof t?.name === "string" ? String(t.name).trim() : "";
            if (!name) continue;
            const sheetName = typeof t?.sheetName === "string" ? String(t.sheetName) : "";
            const range = normalizeDocRange({
              startRow: t?.startRow,
              startCol: t?.startCol,
              endRow: t?.endRow,
              endCol: t?.endCol,
            });
            const description = range
              ? sheetName
                ? `${formatSheetPrefix(sheetName)}${rangeToA1(range)}`
                : rangeToA1(range)
              : undefined;

            const structuredOk = /^[A-Za-z_][A-Za-z0-9_]*$/.test(name);
            const reference = (() => {
              if (!range) return "";
              // Structured refs don't require a sheet prefix, and allow `parseGoTo` to resolve
              // the target sheet from the workbook table metadata.
              if (structuredOk) return `${name}[#All]`;
              // Non-identifier table names require falling back to a sheet-qualified A1 reference.
              // If we don't have a sheet name, leave the entry as "edit-only" rather than navigating
              // to the wrong sheet.
              if (!sheetName) return "";
              return `${formatSheetPrefix(sheetName)}${rangeToA1(range)}`;
            })();

            push({
              kind: "table",
              key: `table:${name}`,
              label: name,
              reference,
              description,
            });
          }

          return items;
        },
      };
      this.formulaBar = new FormulaBarView(
        opts.formulaBar,
        {
        onBeginEdit: () => {
          if (this.isReadOnly()) {
            const cell = this.selection.active;
            showCollabEditRejectedToast([
              { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
            ]);
            return;
          }
          this.formulaEditCell = { sheetId: this.sheetId, cell: { ...this.selection.active } };
          this.syncSharedGridInteractionMode();
          this.updateEditState();
        },
        onGoTo: (reference) => this.goTo(reference),
        onOpenNameBoxMenu: () => this.openNameBoxMenu(),
        onCommit: (text, commit) => this.commitFormulaBar(text, commit),
        onCancel: () => this.cancelFormulaBar(),
        onHoverRange: (range) => {
          this.referencePreview = range
            ? {
                start: { row: range.start.row, col: range.start.col },
                end: { row: range.end.row, col: range.end.col }
              }
            : null;
          if (this.sharedGrid) {
            this.syncSharedGridInteractionMode();
            if (range) {
              const gridRange = this.gridRangeFromDocRange({
                startRow: range.start.row,
                endRow: range.end.row,
                startCol: range.start.col,
                endCol: range.end.col
              });
              this.sharedGrid.setRangeSelection(gridRange);
            } else {
              this.sharedGrid.clearRangeSelection();
            }
          } else {
            this.renderReferencePreview();
          }
        },
        onHoverRangeWithText: (range, refText) => {
          const allowed = this.isFormulaRangePreviewAllowed(refText);
          if (!allowed) {
            // Avoid showing misleading previews/outlines for sheet-qualified refs, named ranges, or table refs
            // that point at a different sheet than the active one.
            this.hideFormulaRangePreviewTooltip();
            this.referencePreview = null;
            if (this.sharedGrid) {
              this.sharedGrid.clearRangeSelection();
            } else {
              this.renderReferencePreview();
            }
            return;
          }
          this.updateFormulaRangePreviewTooltip(range, refText);
        },
        onReferenceHighlights: (highlights) => {
          this.referenceHighlightsSource = highlights;
          this.referenceHighlights = this.computeReferenceHighlightsForSheet(this.sheetId, this.referenceHighlightsSource);
          if (this.sharedGrid) this.syncSharedGridReferenceHighlights();
          this.renderReferencePreview();
          // Keep other views (e.g. split-view secondary pane) in sync with formula-bar reference
          // highlights and formula-editing state. Only emit while the formula bar is actively editing
          // to avoid spurious callbacks from hover behavior in view mode.
          if (this.formulaBar?.isEditing()) {
            for (const listener of this.formulaBarOverlayListeners) {
              listener();
            }
          }
        }
      },
        {
          nameBoxDropdownProvider,
          getWasmEngine: () => this.wasmEngine,
          // SpreadsheetApp does not currently expose an explicit "engine locale" getter.
          // Use the document language (set by the i18n layer) as a best-effort proxy.
          getLocaleId: () => (typeof document !== "undefined" ? document.documentElement?.lang : "") || "en-US",
          referenceStyle: "A1",
        }
      );

      // Excel-style range selection mode: while the formula bar is editing a formula, focus may
      // temporarily move to the grid so keyboard navigation can select a range. When focus returns
      // to the formula bar textarea, end any in-progress keyboard-driven pointing gesture so
      // subsequent arrow keys behave like normal caret navigation.
      this.formulaBar.textarea.addEventListener("focus", () => this.endKeyboardRangeSelection(), {
        signal: this.domAbort.signal,
      });

      // Provide workbook schema context (defined names + tables) to the spreadsheet-frontend
      // reference extractor so formula-bar range highlighting can resolve named ranges and
      // structured table references (e.g. `Table1[Amount]`).
      //
      // NOTE: `DocumentWorkbookAdapter.tables` is keyed by normalized names, but the resolver in
      // `@formula/spreadsheet-frontend` matches case-insensitively and also checks `table.name`.
      this.formulaBar.model.setExtractFormulaReferencesOptions({ tables: this.searchWorkbook.tables as any });
      this.formulaBar.model.setNameResolver((name) => {
        const entry: any = this.searchWorkbook.getName(name);
        const range = entry?.range;
        if (
          !range ||
          typeof range.startRow !== "number" ||
          typeof range.startCol !== "number" ||
          typeof range.endRow !== "number" ||
          typeof range.endCol !== "number"
        ) {
          return null;
        }
        const sheet = typeof entry?.sheetName === "string" && entry.sheetName.trim() ? entry.sheetName.trim() : undefined;
        return {
          startRow: range.startRow,
          startCol: range.startCol,
          endRow: range.endRow,
          endCol: range.endCol,
          sheet,
        };
      });
      this.formulaBar.setArgumentPreviewProvider((expr) => this.evaluateFormulaBarArgumentPreview(expr));
      this.formulaRangePreviewTooltip = this.createFormulaRangePreviewTooltip();
      opts.formulaBar.appendChild(this.formulaRangePreviewTooltip);

      this.formulaBarCompletion = new FormulaBarTabCompletionController({
        formulaBar: this.formulaBar,
        document: this.document,
        getSheetId: () => this.sheetId,
        getEngineClient: () => this.wasmEngine,
        sheetNameResolver: this.sheetNameResolver ?? undefined,
        limits: this.limits,
        schemaProvider: {
          getNamedRanges: () => {
            const formatSheetPrefix = (id: string): string => {

              const token = formatSheetNameForA1(id);
              return token ? `${token}!` : "";
            };

            const out: Array<{ name: string; range?: string }> = [];
            for (const entry of this.searchWorkbook.names.values()) {
              const e: any = entry as any;
              const name = typeof e?.name === "string" ? (e.name as string) : "";
              if (!name) continue;
              const sheetName = typeof e?.sheetName === "string" ? (e.sheetName as string) : "";
              const range = e?.range;
              const rangeText = sheetName && range ? `${formatSheetPrefix(sheetName)}${rangeToA1(range)}` : undefined;
              out.push({ name, range: rangeText });
            }
            return out;
          },
          getTables: () =>
            Array.from(this.searchWorkbook.tables.values())
              .map((table: any) => ({
                name: typeof table?.name === "string" ? table.name : "",
                sheetName: typeof table?.sheetName === "string" ? table.sheetName : undefined,
                startRow: typeof table?.startRow === "number" ? table.startRow : undefined,
                startCol: typeof table?.startCol === "number" ? table.startCol : undefined,
                endRow: typeof table?.endRow === "number" ? table.endRow : undefined,
                endCol: typeof table?.endCol === "number" ? table.endCol : undefined,
                columns: Array.isArray(table?.columns) ? table.columns.map((c: unknown) => String(c)) : [],
              }))
              .filter((t: { name: string; columns: string[] }) => t.name.length > 0 && t.columns.length > 0),
          getCacheKey: () => `schema:${Number((this.searchWorkbook as any).schemaVersion) || 0}`,
        },
      });
    }

    // Apply initial read-only UI state now that optional UI surfaces (formula bar)
    // may be mounted.
    this.syncReadOnlyState();

    if (this.gridMode === "legacy") {
      // Precompute row/col visibility + mappings before any initial render work.
      //
      // NOTE: This work is intentionally *skipped* in shared-grid mode so the app
      // can support Excel-scale sheets without O(maxRows/maxCols) upfront work.
      this.rebuildAxisVisibilityCache();
    }

    if (!collabEnabled) {
      // Seed a demo chart using the chart store helpers so it matches the logic
      // used by AI chart creation.
      this.chartStore.createChart({
        chart_type: "bar",
        // Use an unqualified range so ChartStore can resolve the active sheet id even
        // when a `sheetNameResolver` is present (display names may not include "Sheet1").
        data_range: "A2:B5",
        title: "Example Chart"
      });
    }

    const workbookId = localWorkbookId;
    const dlp = createDesktopDlpContext({ documentId: workbookId });
    this.dlpContext = dlp;
    this.aiCellFunctions = new AiCellFunctionEngine({
      onUpdate: () => this.refresh(),
      workbookId,
      sheetNameResolver: this.sheetNameResolver,
      cache: { persistKey: "formula:ai_cell_cache" },
    });

    // Initial layout + render.
    this.onResize();
    if (this.sharedGrid) {
      // Ensure the shared renderer selection layer matches the app selection model.
      this.syncSharedGridSelectionFromState();
    }

    if (collabEnabled && this.collabSession) {
      // Conflicts UI (mounted once; new conflicts stream in via the monitor callbacks).
      this.conflictUiContainer = document.createElement("div");
      // Avoid blocking grid interactions unless the user is interacting with the conflict UI itself.
      this.conflictUiContainer.classList.add("conflict-ui-overlay");
      this.root.appendChild(this.conflictUiContainer);

      // Collaborators list (names + colors), similar to a Google Docs avatar strip.
      // This is an always-on overlay, but is non-interactive so it never blocks grid pointer events.
      this.collaboratorsListContainer = document.createElement("div");
      const statusbarMain = typeof document !== "undefined" ? document.querySelector(".statusbar__main") : null;
      if (statusbarMain instanceof HTMLElement) {
        // Prefer mounting in the status bar when available so we never occlude grid content.
        this.collaboratorsListContainer.classList.add("presence-collaborators-statusbar");
        const collabBlock = statusbarMain.querySelector(".statusbar__collab");
        if (collabBlock) statusbarMain.insertBefore(this.collaboratorsListContainer, collabBlock);
        else statusbarMain.appendChild(this.collaboratorsListContainer);
      } else {
        // Unit tests often mount SpreadsheetApp with a minimal DOM; fall back to an in-grid overlay.
        this.collaboratorsListContainer.classList.add("presence-collaborators-overlay");
        this.root.appendChild(this.collaboratorsListContainer);
      }

      // Cap visible items to keep the list compact; show a "+N" pill when more are present.
      this.collaboratorsListUi = new CollaboratorsListUiController({
        container: this.collaboratorsListContainer,
        maxVisible: statusbarMain instanceof HTMLElement ? 3 : 5,
      });

      this.conflictUi = new ConflictUiController({
        container: this.conflictUiContainer,
        sheetNameResolver: this.sheetNameResolver,
        onNavigateToCell: (cellRef: { sheetId: string; row: number; col: number }) => this.navigateToConflictCell(cellRef),
        resolveUserLabel: (userId: string) => this.resolveRemoteUserLabel(userId),
        monitor: {
          resolveConflict: (id: string, chosen: unknown) => {
            const monitor = this.collabSession?.formulaConflictMonitor;
            return monitor ? monitor.resolveConflict(id, chosen) : false;
          },
        },
      });

      if (this.pendingFormulaConflicts.length > 0) {
        const queued = this.pendingFormulaConflicts;
        this.pendingFormulaConflicts = [];
        for (const conflict of queued) {
          this.conflictUi.addConflict(conflict);
        }
      }

      // Re-enable pointer events for the conflict UX primitives.
      const toastRoot = this.conflictUiContainer.querySelector<HTMLElement>('[data-testid="conflict-toast-root"]');
      if (toastRoot) {
        toastRoot.classList.add("conflict-ui-toast-root");
      }
      const dialogRoot = this.conflictUiContainer.querySelector<HTMLElement>('[data-testid="conflict-dialog-root"]');
      if (dialogRoot) {
        dialogRoot.classList.add("conflict-ui-dialog-root");
      }

      this.structuralConflictUi = new StructuralConflictUiController({
        container: this.conflictUiContainer,
        sheetNameResolver: this.sheetNameResolver,
        onNavigateToCell: (cellRef: { sheetId: string; row: number; col: number }) => this.navigateToConflictCell(cellRef),
        resolveUserLabel: (userId: string) => this.resolveRemoteUserLabel(userId),
        monitor: {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          resolveConflict: (id: string, resolution: any) => {
            const monitor = this.collabSession?.cellConflictMonitor;
            return monitor ? monitor.resolveConflict(id, resolution) : false;
          },
        },
      });

      const structuralToastRoot = this.conflictUiContainer.querySelector<HTMLElement>(
        '[data-testid="structural-conflict-toast-root"]',
      );
      if (structuralToastRoot) {
        structuralToastRoot.classList.add("structural-conflict-ui-toast-root");
      }
      const structuralDialogRoot = this.conflictUiContainer.querySelector<HTMLElement>(
        '[data-testid="structural-conflict-dialog-root"]',
      );
      if (structuralDialogRoot) {
        structuralDialogRoot.classList.add("structural-conflict-ui-dialog-root");
      }

      if (this.pendingStructuralConflicts.length > 0) {
        const queued = this.pendingStructuralConflicts;
        this.pendingStructuralConflicts = [];
        for (const conflict of queued) {
          this.structuralConflictUi.addConflict(conflict);
        }
      }

      const presence = this.collabSession.presence;
      if (presence) {
        const defaultPresenceColor = resolveCssVar("--formula-grid-remote-presence-default", {
          fallback: resolveCssVar("--accent", { fallback: resolveCssVar("--text-primary", { fallback: "CanvasText" }) }),
        });
        // Publish local selection state.
        this.collabSelectionUnsubscribe = this.subscribeSelection((selection) => {
          presence.setCursor({ row: selection.active.row, col: selection.active.col });
          presence.setSelections(selection.ranges);
        });

        // Render remote presences.
        this.collabPresenceUnsubscribe = presence.subscribe((presences: any[]) => {
          this.remotePresences = (Array.isArray(presences) ? presences : []).map((p) => {
            const cursor =
              p?.cursor && typeof p.cursor.row === "number" && typeof p.cursor.col === "number"
                ? { row: Math.trunc(p.cursor.row), col: Math.trunc(p.cursor.col) }
                : null;
            const selections = Array.isArray(p?.selections)
              ? p.selections
                  .map((r: any) =>
                    r &&
                    typeof r.startRow === "number" &&
                    typeof r.startCol === "number" &&
                    typeof r.endRow === "number" &&
                    typeof r.endCol === "number"
                      ? {
                          startRow: Math.trunc(r.startRow),
                          startCol: Math.trunc(r.startCol),
                          endRow: Math.trunc(r.endRow),
                          endCol: Math.trunc(r.endCol),
                        }
                      : null
                  )
                  .filter((r: any) => r != null)
              : [];

            return {
              id: String(p?.id ?? ""),
              name: String(p?.name ?? t("presence.anonymous")),
              color: String(p?.color || defaultPresenceColor),
              cursor,
              selections,
            } satisfies GridPresence;
          });

          // Update the collaborators list UI. This is intentionally decoupled from cursor/selection
          // movement: the UI controller diffs relevant fields (name/color/sheet) to avoid reflowing
          // on every remote cursor update.
          if (this.collaboratorsListUi) {
            const allPresences = presence.getRemotePresences({ includeOtherSheets: true });
            const activeSheet = this.sheetId;
            const collaborators: CollaboratorListEntry[] = allPresences.map((p: any) => {
              const id = String(p?.id ?? "");
              const clientId = typeof p?.clientId === "number" ? p.clientId : -1;
              const key = `${id || "client"}:${clientId}`;
              const name = String(p?.name ?? t("presence.anonymous"));
              const color = String(p?.color || defaultPresenceColor);
              const sheetId = String(p?.activeSheet ?? "");
              const sheetName =
                sheetId && sheetId !== activeSheet
                  ? this.sheetNameResolver?.getSheetNameById(sheetId) ?? sheetId
                  : null;
              return { key, name, color, sheetName };
            });
            this.collaboratorsListUi.setCollaborators(collaborators);
          }

          if (this.sharedGrid) {
            const headerRows = this.sharedHeaderRows();
            const headerCols = this.sharedHeaderCols();
            const mapped: GridPresence[] = this.remotePresences.map((p) => ({
              ...p,
              cursor: p.cursor ? { row: p.cursor.row + headerRows, col: p.cursor.col + headerCols } : null,
              selections: p.selections.map((r) => ({
                startRow: r.startRow + headerRows,
                startCol: r.startCol + headerCols,
                endRow: r.endRow + headerRows,
                endCol: r.endCol + headerCols,
              })),
            }));
            this.sharedGrid.renderer.setRemotePresences(mapped);
          } else {
            this.renderPresence();
          }
        });
      }
    }

    this.uiReady = true;
    this.editState = this.isEditing();
  }

  private dispatchViewChanged(): void {
    // Used by the Ribbon to sync pressed state for view toggles (e.g. when toggled
    // via keyboard shortcuts).
    if (typeof window === "undefined") return;
    window.dispatchEvent(new CustomEvent("formula:view-changed"));
  }

  private dispatchDrawingsChanged(): void {
    if (typeof window === "undefined") return;
    window.dispatchEvent(new Event("formula:drawings-changed"));
  }

  private dispatchDrawingSelectionChanged(): void {
    if (typeof window === "undefined") return;
    window.dispatchEvent(new CustomEvent("formula:drawing-selection-changed", { detail: { id: this.selectedDrawingId } }));
  }

  private dispatchCommentsChanged(): void {
    if (typeof window === "undefined") return;
    window.dispatchEvent(new Event("formula:comments-changed"));
  }

  private dispatchZoomChanged(): void {
    if (typeof window === "undefined") return;
    window.dispatchEvent(new CustomEvent("formula:zoom-changed"));
  }

  private dispatchReadOnlyChanged(): void {
    if (typeof window === "undefined") return;
    window.dispatchEvent(new CustomEvent("formula:read-only-changed", { detail: { readOnly: this.readOnly, role: this.readOnlyRole } }));
  }

  private syncReadOnlyState(): void {
    const session = this.collabSession;
    const nextReadOnly = (() => {
      if (!session) return false;
      const fn = (session as any).isReadOnly;
      if (typeof fn !== "function") return false;
      try {
        return Boolean(fn.call(session));
      } catch {
        return false;
      }
    })();
    const nextRole = (() => {
      if (!session) return null;
      // `getPermissions` is the preferred API, but some tests stub CollabSession and may omit it.
      const perms = (session as any).getPermissions?.();
      const role = typeof perms?.role === "string" ? String(perms.role) : null;
      return role;
    })();

    if (nextReadOnly === this.readOnly && nextRole === this.readOnlyRole) return;
    this.readOnly = nextReadOnly;
    this.readOnlyRole = nextRole;

    // If permissions flipped to read-only mid-edit, cancel any active edit surfaces so the UI
    // doesn't appear to accept local-only changes.
    if (nextReadOnly) {
      if (this.editor.isOpen()) {
        try {
          this.editor.cancel();
        } catch {
          // ignore
        }
      }
      if (this.formulaBar?.isEditing()) {
        try {
          this.formulaBar.cancelEdit();
        } catch {
          // ignore
        }
      }
      if (this.inlineEditController.isOpen()) {
        try {
          this.inlineEditController.close();
        } catch {
          // ignore
        }
      }
    }

    const indicator = this.status.readOnlyIndicator;
    if (indicator) {
      if (nextReadOnly) {
        indicator.hidden = false;
        indicator.textContent = nextRole ? `Read-only (${nextRole})` : "Read-only";
      } else {
        indicator.hidden = true;
        indicator.textContent = "";
      }
    }

    try {
      this.formulaBar?.setReadOnly(nextReadOnly, { role: nextRole });
    } catch {
      // ignore formula bar wiring failures (eg. missing DOM in tests)
    }

    this.dispatchReadOnlyChanged();

    // Permissions changes can also affect comment capabilities (e.g. viewer ↔ commenter).
    // If the comments panel is open, re-render so composer/action disabled states stay in sync.
    try {
      this.renderCommentsPanel();
    } catch {
      // ignore (tests may construct partial app instances)
    }
  }

  destroy(): void {
    this.disposed = true;
    // Ensure any drawing interaction controller listeners are released promptly.
    this.drawingInteractionController?.dispose();
    this.drawingInteractionController = null;
    this.drawingInteractionCallbacks = null;

    // Clear any cached image bytes kept in the drawing ImageStore.
    try {
      (this.drawingImages as any)?.dispose?.();
    } catch {
      // ignore
    }

    // If the "insert image" input was created, ensure it (and its event handler)
    // are released even if the app object remains referenced after dispose.
    if (this.insertImageInput) {
      try {
        this.insertImageInput.onchange = null;
      } catch {
        // ignore
      }
      try {
        this.insertImageInput.remove();
      } catch {
        // ignore
      }
      this.insertImageInput = null;
    }

    // Ensure overlay caches (ImageBitmaps, parsed XML) are released promptly.
    this.drawingOverlay?.destroy?.();
    this.chartSelectionOverlay?.destroy?.();
    this.chartRenderer.destroy();
    this.activeSheetBackgroundAbort?.abort();
    this.activeSheetBackgroundAbort = null;
    this.workbookImageBitmaps.clear();
    this.activeSheetBackgroundBitmap = null;
    this.sheetViewBinder?.destroy();
    this.sheetViewBinder = null;
    this.imageBytesBinder?.destroy();
    this.imageBytesBinder = null;
    this.domAbort.abort();
    this.chartDragAbort?.abort();
    this.chartDragAbort = null;
    this.chartDragState = null;
    this.chartDrawingInteraction?.dispose();
    this.chartDrawingInteraction = null;
    this.setCollabUndoService(null);
    this.reservedRootGuardToastUnsubscribe?.();
    this.reservedRootGuardToastUnsubscribe = null;
    this.collabPermissionsUnsubscribe?.();
    this.collabPermissionsUnsubscribe = null;
    this.collabSelectionUnsubscribe?.();
    this.collabSelectionUnsubscribe = null;
    this.collabPresenceUnsubscribe?.();
    this.collabPresenceUnsubscribe = null;
    this.collabBinder?.destroy();
    this.collabBinder = null;
    this.collabBinderInitPromise = null;
    // Tests sometimes inject a minimal collab session stub (e.g. `{ presence }`)
    // to exercise presence behavior without spinning up a provider. Use optional
    // method calls so teardown remains resilient.
    this.collabSession?.disconnect?.();
    this.collabSession?.destroy?.();
    this.collabSession = null;
    this.collabEncryptionKeyStore = null;
    this.encryptedRangeManager = null;
    this.stopCommentsRootObserver?.();
    this.stopCommentsRootObserver = null;
    if (this.commentsDocUpdateListener) {
      this.commentsDoc.off("update", this.commentsDocUpdateListener);
      this.commentsDocUpdateListener = null;
    }

    this.conflictUi?.destroy();
    this.conflictUi = null;
    this.collaboratorsListUi?.destroy();
    this.collaboratorsListUi = null;
    this.collaboratorsListContainer?.remove();
    this.collaboratorsListContainer = null;
    this.structuralConflictUi?.destroy();
    this.structuralConflictUi = null;
    this.pendingFormulaConflicts = [];
    this.pendingStructuralConflicts = [];

    this.formulaBarCompletion?.destroy();
    this.syncFormulaRangePreviewTooltipDescribedBy(false);
    this.formulaRangePreviewTooltip?.remove();
    this.formulaRangePreviewTooltip = null;
    this.formulaRangePreviewTooltipVisible = false;
    this.formulaRangePreviewTooltipLastKey = null;
    this.sharedGrid?.destroy();
    this.sharedGrid = null;
    this.sharedProvider = null;
    this.wasmUnsubscribe?.();
    this.wasmUnsubscribe = null;
    this.auditingUnsubscribe?.();
    this.auditingUnsubscribe = null;
    this.externalRepaintUnsubscribe?.();
    this.externalRepaintUnsubscribe = null;
    this.drawingsUnsubscribe?.();
    this.drawingsUnsubscribe = null;
    this.wasmEngine?.terminate();
    this.wasmEngine = null;
    this.stopCommentPersistence?.();
    this.stopCommentPersistence = null;
    this.resizeObserver.disconnect();
    if (this.dragAutoScrollRaf != null) {
      if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(this.dragAutoScrollRaf);
      else globalThis.clearTimeout(this.dragAutoScrollRaf);
      this.dragAutoScrollRaf = null;
    }
    this.outlineButtons.clear();
    this.chartModels.clear();
    this.dirtyChartIds.clear();
    this.chartHasFormulaCells.clear();
    this.chartRangeRectsCache.clear();
    this.chartSelectionViewportMemo = null;
    this.conflictUiContainer = null;

    // Drop references to drawings state so a disposed app instance does not retain
    // large drawing metadata/images if it is kept alive (e.g. tests/hot reload).
    //
    // Do this after disposing drawing interaction controllers since they may cancel an
    // in-progress gesture by restoring the pre-gesture object list.
    this.selectedDrawingId = null;
    this.drawingObjectsCache = null;
    this.drawingHitTestIndex = null;
    this.drawingHitTestIndexObjects = null;
    this.drawingObjects = [];
    this.root.replaceChildren();
  }

  /**
   * Alias for `destroy()` (preferred by some call sites/tests).
   */
  dispose(): void {
    this.destroy();
  }

  /**
   * Request a full redraw of the grid + overlays.
   *
   * This is intentionally public so external callers (e.g. AI tool execution)
   * can update the UI after mutating the DocumentController.
   */
  refresh(mode: "full" | "scroll" = "full"): void {
    if (this.disposed) return;
    if (this.renderScheduled) {
      // Upgrade a pending scroll-only render to a full render if needed.
      if (mode === "full") this.pendingRenderMode = "full";
      return;
    }

    this.renderScheduled = true;
    this.pendingRenderMode = mode;

    const schedule =
      typeof requestAnimationFrame === "function"
        ? requestAnimationFrame
        : (cb: FrameRequestCallback) =>
            globalThis.setTimeout(() => cb(typeof performance !== "undefined" ? performance.now() : Date.now()), 0);

    schedule(() => {
      this.renderScheduled = false;
      if (this.disposed) return;
      const renderMode = this.pendingRenderMode;
      const renderContent = renderMode === "full";
      this.pendingRenderMode = "full";
      this.renderGrid();
      if (this.useCanvasCharts && renderContent) {
        this.invalidateCanvasChartsForActiveSheet();
      }
      this.renderDrawings();
      if (!this.useCanvasCharts) {
        this.renderCharts(false);
      }
      this.renderReferencePreview();
      this.renderAuditing();
      this.renderPresence();
      this.renderSelection();
      if (renderMode === "full") this.updateStatus();
    });
  }

  private scheduleStatusUpdate(): void {
    if (this.disposed) return;
    if (!this.uiReady) return;
    // A pending full refresh will update status anyway.
    if (this.renderScheduled && this.pendingRenderMode === "full") return;
    if (this.statusUpdateScheduled) return;
    this.statusUpdateScheduled = true;

    const schedule =
      typeof requestAnimationFrame === "function"
        ? requestAnimationFrame
        : (cb: FrameRequestCallback) =>
            globalThis.setTimeout(() => cb(typeof performance !== "undefined" ? performance.now() : Date.now()), 0);

    schedule(() => {
      this.statusUpdateScheduled = false;
      if (this.disposed) return;
      if (!this.uiReady) return;
      this.updateStatus();
    });
  }

  private scheduleChartContentRefresh(payload?: any): void {
    if (this.disposed) return;
    if (!this.uiReady) return;
    // A pending refresh will render chart content anyway.
    if (this.renderScheduled) return;
    if (this.chartContentRefreshScheduled) return;
    const charts = this.chartStore.listCharts().filter((chart) => chart.sheetId === this.sheetId);
    if (charts.length === 0) return;

    // In non-engine modes (or multi-sheet), formula charts can change due to edits outside their
    // direct ranges. If we see a recalc event, conservatively mark formula-containing charts as
    // dirty so their cached series data will refresh on the next render.
    if (payload?.recalc) {
      const sheetCount = (this.document as any)?.model?.sheets?.size;
      const useEngineCache =
        (typeof sheetCount === "number" ? sheetCount : this.document.getSheetIds().length) <= 1;
      const hasWasmEngine = Boolean(this.wasmEngine && !this.wasmSyncSuspended);
      if (!hasWasmEngine || !useEngineCache) {
        this.markFormulaChartsDirty();
      }
    }

    if (this.dirtyChartIds.size === 0) return;

    if (!this.sharedGrid) {
      // Chart rect computations depend on frozen pane geometry; keep legacy caches current.
      this.ensureViewportMappingCurrent();
    }

    const intersects = (
      a: { x: number; y: number; width: number; height: number },
      b: { x: number; y: number; width: number; height: number },
    ): boolean => {
      return !(
        a.x + a.width < b.x ||
        b.x + b.width < a.x ||
        a.y + a.height < b.y ||
        b.y + b.height < a.y
      );
    };

    const layout = this.chartOverlayLayout(this.sharedGrid ? this.sharedGrid.renderer.scroll.getViewportState() : undefined);
    const { frozenRows, frozenCols } = this.getFrozen();

    const hasVisibleDirtyChart = charts.some((chart) => {
      if (!this.dirtyChartIds.has(chart.id)) return false;

      const rect = this.chartAnchorToViewportRect(chart.anchor);
      if (!rect) return false;

      const chartRect = { x: rect.left, y: rect.top, width: rect.width, height: rect.height };

      const fromRow = chart.anchor.kind === "oneCell" || chart.anchor.kind === "twoCell" ? chart.anchor.fromRow : Number.POSITIVE_INFINITY;
      const fromCol = chart.anchor.kind === "oneCell" || chart.anchor.kind === "twoCell" ? chart.anchor.fromCol : Number.POSITIVE_INFINITY;
      const inFrozenRows = fromRow < frozenRows;
      const inFrozenCols = fromCol < frozenCols;
      const pane =
        inFrozenRows && inFrozenCols
          ? layout.paneRects.topLeft
          : inFrozenRows && !inFrozenCols
            ? layout.paneRects.topRight
            : !inFrozenRows && inFrozenCols
              ? layout.paneRects.bottomLeft
              : layout.paneRects.bottomRight;

      if (pane.width <= 0 || pane.height <= 0) return false;
      return intersects(chartRect, pane);
    });

    if (!hasVisibleDirtyChart) return;

    this.chartContentRefreshScheduled = true;

    const schedule =
      typeof requestAnimationFrame === "function"
        ? requestAnimationFrame
        : (cb: FrameRequestCallback) =>
            globalThis.setTimeout(() => cb(typeof performance !== "undefined" ? performance.now() : Date.now()), 0);

    schedule(() => {
      this.chartContentRefreshScheduled = false;
      if (this.disposed) return;
      if (!this.uiReady) return;
      this.renderCharts(false);
    });
  }

  private scheduleDrawingsRender(reason?: string): void {
    if (this.disposed) return;
    if (!this.uiReady) return;
    // A pending refresh will redraw drawings (and other overlays) anyway.
    if (this.renderScheduled) return;
    if (this.drawingsRenderScheduled) return;

    // Keep the reason in the signature so callers can provide context without allocations.
    void reason;

    this.drawingsRenderScheduled = true;

    const schedule =
      typeof requestAnimationFrame === "function"
        ? requestAnimationFrame
        : (cb: FrameRequestCallback) =>
            globalThis.setTimeout(() => cb(typeof performance !== "undefined" ? performance.now() : Date.now()), 0);

    schedule(() => {
      this.drawingsRenderScheduled = false;
      if (this.disposed) return;
      if (!this.uiReady) return;
      // In shared-grid mode, axis size changes (and related scroll clamping/alignment) can update the
      // renderer's internal scroll offsets even when the user didn't actively scroll. Keep our legacy
      // scroll state in sync before computing drawing viewport geometry so overlays remain pixel-aligned.
      if (this.sharedGrid) {
        const scroll = this.sharedGrid.getScroll();
        this.scrollX = scroll.x;
        this.scrollY = scroll.y;
      }
      this.renderDrawings();
    });
  }

  private markChartsDirtyFromDeltas(deltas: unknown): void {
    if (!Array.isArray(deltas) || deltas.length === 0) return;

    const charts = this.chartStore.listCharts().filter((chart) => chart.sheetId === this.sheetId);
    if (charts.length === 0) return;

    type RangeRect = { chartId: string; startRow: number; endRow: number; startCol: number; endCol: number };
    const rangesBySheet = new Map<string, RangeRect[]>();

    const getChartRanges = (chart: ChartRecord): Array<{ sheetId: string; startRow: number; endRow: number; startCol: number; endCol: number }> => {
      let signature = "";
      for (const ser of chart.series ?? []) {
        signature += `|${ser.categories ?? ""}|${ser.values ?? ""}|${ser.xValues ?? ""}|${ser.yValues ?? ""}`;
      }

      const cached = this.chartRangeRectsCache.get(chart.id);
      if (cached && cached.signature === signature) return cached.ranges;

      const ranges: Array<{ sheetId: string; startRow: number; endRow: number; startCol: number; endCol: number }> = [];
      for (const ser of chart.series ?? []) {
        const refs = [ser.categories, ser.values, ser.xValues, ser.yValues];
        for (const rangeRef of refs) {
          if (typeof rangeRef !== "string" || rangeRef.trim() === "") continue;
          const parsed = parseA1Range(rangeRef);
          if (!parsed) continue;
          const resolvedSheetId = parsed.sheetName ? this.resolveSheetIdByName(parsed.sheetName) : chart.sheetId;
          if (!resolvedSheetId) continue;
          ranges.push({
            sheetId: resolvedSheetId,
            startRow: parsed.startRow,
            endRow: parsed.endRow,
            startCol: parsed.startCol,
            endCol: parsed.endCol,
          });
        }
      }

      this.chartRangeRectsCache.set(chart.id, { signature, ranges });
      return ranges;
    };

    for (const chart of charts) {
      const ranges = getChartRanges(chart);
      for (const range of ranges) {
        let list = rangesBySheet.get(range.sheetId);
        if (!list) {
          list = [];
          rangesBySheet.set(range.sheetId, list);
        }
        list.push({ chartId: chart.id, startRow: range.startRow, endRow: range.endRow, startCol: range.startCol, endCol: range.endCol });
      }
    }

    const affected = new Set<string>();
    for (const delta of deltas) {
      const sheetId = String((delta as any)?.sheetId ?? "");
      const row = Number((delta as any)?.row);
      const col = Number((delta as any)?.col);
      if (!sheetId) continue;
      if (!Number.isInteger(row) || row < 0) continue;
      if (!Number.isInteger(col) || col < 0) continue;

      const ranges = rangesBySheet.get(sheetId);
      if (!ranges) continue;

      for (const range of ranges) {
        if (row < range.startRow || row > range.endRow) continue;
        if (col < range.startCol || col > range.endCol) continue;
        affected.add(range.chartId);
      }
    }

    for (const id of affected) this.dirtyChartIds.add(id);
  }

  private markFormulaChartsDirty(): void {
    const charts = this.chartStore.listCharts().filter((chart) => chart.sheetId === this.sheetId);
    if (charts.length === 0) return;

    for (const chart of charts) {
      const hasFormula = this.chartHasFormulaCells.get(chart.id);
      // If we haven't scanned the chart ranges yet, be conservative and assume formulas
      // might be present so charts don't get stuck with stale cached values.
      if (hasFormula === false) continue;
      this.dirtyChartIds.add(chart.id);
    }
  }

  focus(): void {
    const target = this.focusTargetProvider?.() ?? null;
    if (target && target.isConnected) {
      try {
        (target as any).focus?.({ preventScroll: true });
        return;
      } catch {
        // Fall back to regular focus below.
      }

      try {
        target.focus();
        return;
      } catch {
        // Fall back to focusing the primary grid.
      }
    }

    this.root.focus();
  }

  setFocusTargetProvider(provider: (() => HTMLElement | null) | null): void {
    this.focusTargetProvider = provider;
  }

  /**
   * Focus the formula bar input, if present.
   *
   * This is intentionally public so UI chrome (e.g. sheet tabs) can preserve
   * "formula-bar-driven navigation" workflows while the user is selecting ranges
   * across sheets.
   *
   * `opts.cursor` can be used to control the selection on focus (`"end"` / `"all"`).
   */
  focusFormulaBar(opts: { cursor?: "end" | "all" } = {}): void {
    this.formulaBar?.focus(opts);
  }

  /**
   * After a sheet activation that was triggered by navigation UI (tabs, sheet switcher,
   * overflow menu, Ctrl+PgUp/PgDn, etc.), restore focus to the appropriate editing surface.
   *
   * - Default: focus the grid so the user can immediately type or use shortcuts.
   * - While actively editing a formula in the formula bar (range selection mode): keep the
   *   formula bar focused so sheet switching doesn't interrupt formula editing.
   */
  focusAfterSheetNavigation(): void {
    if (this.formulaBar?.isFormulaEditing()) {
      this.formulaBar.focus();
      return;
    }
    this.focus();
  }

  isCellEditorOpen(): boolean {
    return this.editor.isOpen();
  }

  isFormulaBarEditing(): boolean {
    return Boolean(this.formulaBar?.isEditing() || this.formulaEditCell);
  }

  /**
   * Returns true when the formula bar is actively editing a formula (draft text starts with `=`).
   *
   * This is intentionally narrower than `isFormulaBarEditing()`: we only want to treat the formula
   * bar as a special case for shortcuts like Ctrl/Cmd+PgUp/PgDn when the user is in Excel-like range
   * selection mode while building a formula.
   */
  isFormulaBarFormulaEditing(): boolean {
    return Boolean(this.formulaBar?.isFormulaEditing());
  }

  onFormulaBarOverlayChange(listener: () => void): () => void {
    this.formulaBarOverlayListeners.add(listener);
    listener();
    return () => {
      this.formulaBarOverlayListeners.delete(listener);
    };
  }

  getFormulaReferenceHighlights(): Array<{ start: CellCoord; end: CellCoord; color: string; active: boolean }> {
    return this.referenceHighlights.map((h) => ({
      start: { ...h.start },
      end: { ...h.end },
      color: h.color,
      active: h.active,
    }));
  }

  isEditing(): boolean {
    return this.isCellEditorOpen() || this.isFormulaBarEditing() || this.inlineEditController.isOpen();
  }

  /**
   * Commit any in-progress edits (cell editor / formula bar) without moving selection.
   *
   * This is intended for "command" entry points like File → Save/New/Open/Close/Quit so:
   * - unsaved-change prompts see the latest edit
   * - saves include the latest edit
   */
  commitPendingEditsForCommand(): void {
    if (this.disposed) return;
    // Commit in-cell edits first; if the formula bar is also editing, its commit should
    // win as the most explicit user intent.
    this.editor.commit("command");
    this.formulaBar?.commitEdit();
  }

  onEditStateChange(listener: (isEditing: boolean) => void): () => void {
    this.editStateListeners.add(listener);
    listener(this.isEditing());
    return () => {
      this.editStateListeners.delete(listener);
    };
  }

  private updateEditState(): void {
    const next = this.isEditing();
    if (next === this.editState) return;
    this.editState = next;
    for (const listener of this.editStateListeners) {
      listener(next);
    }
  }

  /**
   * Public clipboard commands (used by split view + other non-primary grid surfaces).
   *
   * These intentionally mirror the keyboard shortcut behavior in `onKeyDown`:
   * - no-ops while editing (cell editor, formula bar, inline edit)
   * - uses the async clipboard helpers tracked by `IdleTracker` so tests can await `whenIdle()`
   */
  copy(): void {
    if (this.inlineEditController.isOpen()) return;
    if (this.isEditing()) return;
    const target = typeof document !== "undefined" ? (document.activeElement as HTMLElement | null) : null;
    if (target) {
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return;
    }
    this.idle.track(this.copySelectionToClipboard());
  }

  cut(): void {
    if (this.inlineEditController.isOpen()) return;
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    const focusTarget = typeof document !== "undefined" ? (document.activeElement as HTMLElement | null) : null;
    if (focusTarget) {
      const tag = focusTarget.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || focusTarget.isContentEditable) return;
    }

    const shouldRestoreFocus = focusTarget != null && focusTarget !== this.root;
    const op = this.cutSelectionToClipboard();
    const tracked = op.finally(() => {
      if (!shouldRestoreFocus) return;
      if (!focusTarget) return;
      if (!focusTarget.isConnected) return;
      try {
        focusTarget.focus();
      } catch {
        // Ignore focus restore failures.
      }
    });
    this.idle.track(tracked);
  }

  paste(): void {
    if (this.inlineEditController.isOpen()) return;
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    const focusTarget = typeof document !== "undefined" ? (document.activeElement as HTMLElement | null) : null;
    if (focusTarget) {
      const tag = focusTarget.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || focusTarget.isContentEditable) return;
    }

    const shouldRestoreFocus = focusTarget != null && focusTarget !== this.root;
    const op = this.pasteClipboardToSelection();
    const tracked = op.finally(() => {
      if (!shouldRestoreFocus) return;
      if (!focusTarget) return;
      if (!focusTarget.isConnected) return;
      try {
        focusTarget.focus();
      } catch {
        // Ignore focus restore failures.
      }
    });
    this.idle.track(tracked);
  }

  clearSelection(): void {
    if (this.inlineEditController.isOpen()) return;
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.clearSelectionContentsInternal();
    this.refresh();
  }

  async clipboardCopy(): Promise<void> {
    await this.copySelectionToClipboard();
  }

  async clipboardCut(): Promise<void> {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    await this.cutSelectionToClipboard();
  }

  async clipboardPaste(): Promise<void> {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    await this.pasteClipboardToSelection();
  }

  async clipboardPasteSpecial(
    mode: "all" | "values" | "formulas" | "formats" = "all",
    options: { transpose?: boolean } = {}
  ): Promise<void> {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (!this.shouldHandleSpreadsheetClipboardCommand()) return;

    const normalized: "all" | "values" | "formulas" | "formats" =
      mode === "values" || mode === "formulas" || mode === "formats" ? mode : "all";

    const promise = this.pasteClipboardToSelection({ mode: normalized, transpose: options.transpose === true });
    this.idle.track(promise);
    await promise;
  }

  clearContents(): void {
    // `clearSelectionContents()` is a public command surface used by the command registry / context
    // menus. It already triggers a refresh, but `clearContents()` is also called by some internal
    // UX flows; avoid double-refreshing by calling the underlying mutation directly here.
    this.clearSelectionContentsInternal();
    this.refresh();
  }

  setFormulaBarDraft(text: string, opts: { cursorStart?: number; cursorEnd?: number } = {}): void {
    const bar = this.formulaBar;
    if (!bar) return;

    // Ensure the formula bar enters editing mode so the view renders the textarea
    // and keeps its internal model in sync.
    if (!bar.isEditing()) {
      this.formulaEditCell = { sheetId: this.sheetId, cell: { ...this.selection.active } };
      this.syncSharedGridInteractionMode();
      this.updateEditState();
      bar.model.beginEdit();
    }

    const textarea = bar.textarea;
    textarea.value = text;
    const start = opts.cursorStart ?? text.length;
    const end = opts.cursorEnd ?? start;
    textarea.setSelectionRange(start, end);
    // Drive the existing input listeners so the FormulaBarModel updates and the
    // highlight/hint UI stays consistent.
    textarea.dispatchEvent(new Event("input", { bubbles: true }));
  }

  insertIntoFormulaBar(
    text: string,
    opts: { replaceSelection?: boolean; focus?: boolean; cursorOffset?: number } = {}
  ): void {
    const bar = this.formulaBar;
    if (!bar) return;

    const focus = opts.focus ?? true;
    const replaceSelection = opts.replaceSelection ?? true;
    const cursorOffset = opts.cursorOffset ?? text.length;
    const focusTextarea = (textarea: HTMLTextAreaElement): void => {
      // The formula bar textarea is `display: none` while not editing. When we insert text
      // programmatically we may focus before the view's next render frame toggles the
      // `formula-bar--editing` class, so eagerly add the class to make focus reliable.
      bar.root.classList.add("formula-bar--editing");
      try {
        textarea.focus({ preventScroll: true });
      } catch {
        textarea.focus();
      }
    };

    // If the user isn't already editing the formula bar, treat insertion as a full
    // replacement (Excel-esque "start editing with this template").
    if (!bar.isEditing()) {
      // Avoid relying on the textarea focus handler to begin edit; drive the
      // FormulaBarModel directly so programmatic insertions work reliably.
      this.formulaEditCell = { sheetId: this.sheetId, cell: { ...this.selection.active } };
      this.syncSharedGridInteractionMode();
      this.updateEditState();
      bar.model.beginEdit();

      const textarea = bar.textarea;
      textarea.value = text;
      const cursor = Math.max(0, Math.min(cursorOffset, textarea.value.length));
      textarea.setSelectionRange(cursor, cursor);
      textarea.dispatchEvent(new Event("input", { bubbles: true }));
      if (focus) focusTextarea(textarea);
      return;
    }

    const textarea = bar.textarea;
    const current = textarea.value;
    const selStart = textarea.selectionStart ?? current.length;
    const selEnd = textarea.selectionEnd ?? current.length;

    const start = Math.max(0, Math.min(Math.min(selStart, selEnd), current.length));
    const end = Math.max(0, Math.min(Math.max(selStart, selEnd), current.length));

    const insertAt = replaceSelection ? start : end;
    // When editing an existing formula, callers often want to insert a function
    // template like `=SUM()` at the caret. Avoid injecting an extra `=` in the
    // middle of a formula (Excel behavior inserts `SUM()` inside an existing `=` draft).
    const stripLeadingEquals = current.trimStart().startsWith("=") && insertAt > 0 && text.startsWith("=");
    const textToInsert = stripLeadingEquals ? text.slice(1) : text;
    const effectiveCursorOffset = stripLeadingEquals ? Math.max(0, cursorOffset - 1) : cursorOffset;

    const nextValue = replaceSelection
      ? current.slice(0, start) + textToInsert + current.slice(end)
      : current.slice(0, insertAt) + textToInsert + current.slice(insertAt);

    textarea.value = nextValue;
    const cursor = Math.max(0, Math.min(insertAt + effectiveCursorOffset, textarea.value.length));
    textarea.setSelectionRange(cursor, cursor);
    textarea.dispatchEvent(new Event("input", { bubbles: true }));
    if (focus) focusTextarea(textarea);
  }

  async whenIdle(): Promise<void> {
    // `wasmSyncPromise` and `auditingIdlePromise` are growing promise chains (see `enqueueWasmSync`
    // and `scheduleAuditingUpdate`). While awaiting them, additional work can be appended
    // (e.g. a batched `setCells` update queued after an engine apply). Loop until both
    // chains stabilize so callers (notably Playwright) can reliably await the
    // "no more background work pending" condition.
    while (true) {
      const wasm = this.wasmSyncPromise;
      const auditing = this.auditingIdlePromise;
      await Promise.all([this.idle.whenIdle(), wasm.catch(() => {}), auditing.catch(() => {})]);
      if (this.wasmSyncPromise === wasm && this.auditingIdlePromise === auditing) return;
    }
  }

  getRecalcCount(): number {
    return this.engine.recalcCount;
  }

  getDocument(): DocumentController {
    return this.document;
  }

  undo(): boolean {
    if (this.isReadOnly()) return false;
    if (this.editor.isOpen()) return false;
    if (this.formulaBar?.isEditing()) return false;
    return this.applyUndoRedo("undo");
  }

  redo(): boolean {
    if (this.isReadOnly()) return false;
    if (this.editor.isOpen()) return false;
    if (this.formulaBar?.isEditing()) return false;
    return this.applyUndoRedo("redo");
  }

  getUndoRedoState(): UndoRedoState {
    const editing = this.isEditing();
    const allowUndoRedo = !editing && !this.isReadOnly();
    if (this.collabUndoService) {
      return {
        canUndo: allowUndoRedo && this.collabUndoService.canUndo(),
        canRedo: allowUndoRedo && this.collabUndoService.canRedo(),
        // Collab undo/redo is managed by Yjs; DocumentController history labels are not
        // meaningful in this mode.
        undoLabel: null,
        redoLabel: null,
      };
    }
    return {
      canUndo: allowUndoRedo && Boolean(this.document.canUndo),
      canRedo: allowUndoRedo && Boolean(this.document.canRedo),
      undoLabel: (this.document.undoLabel as string | null) ?? null,
      redoLabel: (this.document.redoLabel as string | null) ?? null,
    };
  }

  /**
   * Collaboration integration hook.
   *
   * When a CollabSession is bound to this app's DocumentController, the integration
   * layer should set a collab undo service so keyboard shortcuts use Yjs UndoManager
   * semantics (local-only undo that never overwrites remote edits).
   */
  setCollabUndoService(undoService: UndoService | null): void {
    this.collabUndoService = undoService;
  }

  getSearchWorkbook(): DocumentWorkbookAdapter {
    return this.searchWorkbook;
  }

  getCurrentSheetId(): string {
    return this.sheetId;
  }

  /**
   * Returns the user-facing display name for a stable sheet id.
   */
  getSheetDisplayNameById(sheetId: string): string {
    return this.resolveSheetDisplayNameById(sheetId);
  }

  /**
   * Returns the user-facing display name for the currently active sheet.
   */
  getCurrentSheetDisplayName(): string {
    return this.resolveSheetDisplayNameById(this.sheetId);
  }

  /**
   * Replace the workbook-scoped image store.
   *
   * This is currently used by:
   * - in-cell images (shared-grid mode via `CanvasGridImageResolver`)
   * - worksheet background images (tiled pattern behind the grid)
   */
  setWorkbookImages(images: ImageEntry[]): void {
    // Cancel any in-flight background image decode so stale results don't race
    // the updated image bytes (and to avoid leaking decoded bitmaps if the
    // previous decode completes after caches are cleared).
    this.activeSheetBackgroundAbort?.abort();
    this.activeSheetBackgroundAbort = null;

    const doc: any = this.document as any;

    // Normalize existing image store into a stable map keyed by trimmed ids.
    const existingById = new Map<string, unknown>();
    const existingImages: unknown = doc?.images;
    if (existingImages instanceof Map) {
      for (const [rawId, entry] of existingImages.entries()) {
        const id = String(rawId ?? "").trim();
        if (!id) continue;
        if (!existingById.has(id)) existingById.set(id, entry);
      }
    } else if (existingImages && typeof existingImages === "object") {
      for (const [rawId, entry] of Object.entries(existingImages as Record<string, unknown>)) {
        const id = String(rawId ?? "").trim();
        if (!id) continue;
        if (!existingById.has(id)) existingById.set(id, entry);
      }
    }

    const nextById = new Map<string, ImageEntry>();
    for (const entry of images) {
      const id = typeof entry?.id === "string" ? entry.id.trim() : "";
      if (!id) continue;
      if (!(entry?.bytes instanceof Uint8Array)) continue;
      nextById.set(id, entry);
    }

    // Prefer a non-undoable "external deltas" path for hydration/testing hooks.
    // This avoids polluting undo history and (by default) does not mark the document dirty.
    if (typeof doc?.applyExternalImageDeltas === "function") {
      /** @type {any[]} */
      const deltas: any[] = [];

      for (const [id, before] of existingById.entries()) {
        if (nextById.has(id)) continue;
        deltas.push({ imageId: id, before, after: null });
      }

      for (const [id, entry] of nextById.entries()) {
        const before = existingById.get(id) ?? null;
        deltas.push({ imageId: id, before, after: { bytes: entry.bytes, mimeType: entry.mimeType } });
      }

      try {
        doc.applyExternalImageDeltas(deltas, { source: "setWorkbookImages", markDirty: false });
      } catch {
        // ignore
      }
    } else {
      // Fallback: apply via undoable mutations (legacy/older controller builds).
      const existingIds = Array.from(existingById.keys());
      // Batch so workbook image hydration becomes a single undo step.
      this.document.beginBatch({ label: "Set workbook images" });
      try {
        for (const id of existingIds) {
          if (nextById.has(id)) continue;
          try {
            this.document.deleteImage(id, { source: "setWorkbookImages" });
          } catch {
            // ignore
          }
        }

        for (const [id, entry] of nextById) {
          try {
            this.document.setImage(id, { bytes: entry.bytes, mimeType: entry.mimeType }, { source: "setWorkbookImages" });
          } catch {
            // ignore
          }
        }
      } finally {
        this.document.endBatch();
      }
    }

    // ImageBitmap caches live inside individual CanvasGridRenderer instances. When the
    // workbook image store changes, clear cached decoded bitmaps so reused image ids redraw.
    this.sharedGrid?.renderer?.clearImageCache?.();

    // Image bytes may have changed; invalidate decoded bitmaps.
    this.workbookImageBitmaps.clear();
    // Force a reload even when the active sheet points at the same image id.
    this.activeSheetBackgroundImageId = null;
    this.activeSheetBackgroundBitmap = null;
    this.syncActiveSheetBackgroundImage();
  }

  /**
   * Set (or clear) a sheet-level tiled background image.
   */
  setSheetBackgroundImageId(sheetId: string, imageId: string | null): void {
    const key = String(sheetId ?? "").trim();
    if (!key) return;
    const normalized = imageId ? String(imageId).trim() : null;
    const doc: any = this.document as any;
    if (typeof doc.setSheetBackgroundImageId === "function") {
      try {
        doc.setSheetBackgroundImageId(key, normalized || null);
      } catch {
        // ignore
      }
    }
    if (key === this.sheetId) {
      this.syncActiveSheetBackgroundImage();
    }
  }

  getSheetBackgroundImageId(sheetId: string): string | null {
    const key = String(sheetId ?? "").trim();
    if (!key) return null;
    const doc: any = this.document as any;
    if (typeof doc.getSheetBackgroundImageId === "function") {
      try {
        return doc.getSheetBackgroundImageId(key) ?? null;
      } catch {
        return null;
      }
    }
    const view = this.document.getSheetView(key) as any;
    const id = view?.backgroundImageId;
    return typeof id === "string" && id.trim() !== "" ? id : null;
  }

  /**
   * Resolve a sheet display name (as written in formula text) to a stable sheet id.
   *
   * Matching is case-insensitive. When provided, the app's `sheetNameResolver`
   * (typically backed by workbook metadata) is used first, with a fallback to
   * matching raw sheet ids for legacy formulas and tests.
   */
  getSheetIdByName(name: string): string | null {
    return this.resolveSheetIdByName(name);
  }

  getScroll(): { x: number; y: number } {
    return { x: this.scrollX, y: this.scrollY };
  }

  subscribeScroll(listener: (scroll: { x: number; y: number }) => void): () => void {
    this.scrollListeners.add(listener);
    listener(this.getScroll());
    return () => this.scrollListeners.delete(listener);
  }

  subscribeZoom(listener: (zoom: number) => void): () => void {
    this.zoomListeners.add(listener);
    listener(this.getZoom());
    return () => this.zoomListeners.delete(listener);
  }

  setScroll(x: number, y: number): void {
    const changed = this.setScrollInternal(x, y);
    // Shared-grid mode repaints via CanvasGridRenderer and invokes the onScroll callback; the
    // legacy renderer needs an explicit scroll-mode refresh.
    if (changed && !this.sharedGrid) this.refresh("scroll");
  }

  private notifyScrollListeners(): void {
    if (this.scrollListeners.size === 0) return;
    const scroll = this.getScroll();
    for (const listener of this.scrollListeners) {
      listener(scroll);
    }
  }

  private notifyZoomListeners(): void {
    if (this.zoomListeners.size === 0) return;
    const zoom = this.getZoom();
    for (const listener of this.zoomListeners) {
      listener(zoom);
    }
  }
  getGridMode(): DesktopGridMode {
    return this.gridMode;
  }

  supportsZoom(): boolean {
    return this.sharedGrid != null;
  }

  getZoom(): number {
    return this.sharedGrid ? this.sharedGrid.getZoom() : 1;
  }

  setZoom(nextZoom: number): void {
    if (!this.sharedGrid) return;
    const clamped = clampZoom(nextZoom);

    const prev = this.sharedGrid.getZoom();
    if (Math.abs(prev - clamped) < 1e-6) return;

    // DesktopSharedGrid.setZoom will resync scrollbars + emit an onScroll callback
    // (which is where shared-grid overlays are re-positioned and zoom listeners are notified).
    this.sharedGrid.setZoom(clamped);
  }

  zoomToSelection(): void {
    if (!this.sharedGrid) return;
    const ranges = this.selection.ranges;
    if (ranges.length === 0) return;

    let startRow = Number.POSITIVE_INFINITY;
    let startCol = Number.POSITIVE_INFINITY;
    let endRow = Number.NEGATIVE_INFINITY;
    let endCol = Number.NEGATIVE_INFINITY;

    for (const range of ranges) {
      const grid = this.gridRangeFromDocRange(range);
      startRow = Math.min(startRow, grid.startRow);
      startCol = Math.min(startCol, grid.startCol);
      endRow = Math.max(endRow, grid.endRow);
      endCol = Math.max(endCol, grid.endCol);
    }

    if (
      !Number.isFinite(startRow) ||
      !Number.isFinite(startCol) ||
      !Number.isFinite(endRow) ||
      !Number.isFinite(endCol)
    ) {
      return;
    }

    const renderer = this.sharedGrid.renderer;
    const zoom = renderer.getZoom();
    const { rows, cols } = renderer.scroll;
    const selectionWidthPx = cols.positionOf(endCol) - cols.positionOf(startCol);
    const selectionHeightPx = rows.positionOf(endRow) - rows.positionOf(startRow);
    if (selectionWidthPx <= 0 || selectionHeightPx <= 0) return;

    const selectionWidthBase = selectionWidthPx / zoom;
    const selectionHeightBase = selectionHeightPx / zoom;

    const viewport = renderer.scroll.getViewportState();
    const padding = 8;
    const availableWidth = Math.max(1, viewport.width - viewport.frozenWidth - padding * 2);
    const availableHeight = Math.max(1, viewport.height - viewport.frozenHeight - padding * 2);

    const nextZoom = Math.min(availableWidth / selectionWidthBase, availableHeight / selectionHeightBase);
    this.setZoom(nextZoom);
    this.sharedGrid.scrollToCell(startRow, startCol, { align: "start", padding });
  }

  getShowFormulas(): boolean {
    return this.showFormulas;
  }

  setShowFormulas(enabled: boolean): void {
    if (enabled === this.showFormulas) return;
    this.showFormulas = enabled;
    this.sharedProvider?.invalidateAll();
    this.refresh();
    this.dispatchViewChanged();
  }

  toggleShowFormulas(): void {
    this.setShowFormulas(!this.showFormulas);
  }

  getCollabSession(): CollabSession | null {
    return this.collabSession;
  }

  /**
   * Best-effort helper to await any pending DocumentController → Yjs binder work.
   *
   * This is primarily used during teardown/quit flows to reduce the chance of
   * flushing local persistence before queued binder writes have been applied to
   * the shared Yjs document (encryption can make binder writes async).
   */
  async whenCollabBinderIdle(): Promise<void> {
    // If the binder is still being initialized (async import, encryption key store hydration),
    // wait for it so we don't flush local persistence before DocumentController edits have
    // started propagating into Yjs at all.
    const pending = this.collabBinderInitPromise;
    if (pending) {
      try {
        await pending;
      } catch {
        // ignore
      }
      // Allow the binder promise `.then(...)` to install `this.collabBinder`.
      await new Promise<void>((resolve) => queueMicrotask(resolve));
    }
    const binder = this.collabBinder;
    if (!binder || typeof binder.whenIdle !== "function") return;
    try {
      await binder.whenIdle();
    } catch {
      // Best-effort; never block callers (e.g. quit) on binder errors.
    }
  }

  getCollabEncryptionKeyStore(): CollabEncryptionKeyStore | null {
    return this.collabEncryptionKeyStore;
  }

  getEncryptedRangeManager(): EncryptedRangeManager | null {
    return this.encryptedRangeManager;
  }

  /**
   * Force the collab binder to rehydrate from Yjs.
   *
   * This is used when local state changes without mutating the shared document
   * (e.g. importing an encryption key).
   */
  rehydrateCollabBinder(): void {
    this.collabBinder?.rehydrate?.();
  }

  /**
   * Returns true when the current collaboration session is read-only (viewer/commenter).
   *
   * Non-collab/local workbooks are always editable.
   */
  isReadOnly(): boolean {
    const session = this.collabSession;
    if (!session) return false;
    try {
      return Boolean(session.isReadOnly());
    } catch {
      return false;
    }
  }

  /**
   * Test-only helper: returns the viewport-relative rect for a cell address.
   */
  getCellRectA1(a1: string): { x: number; y: number; width: number; height: number } | null {
    const cell = parseA1(a1);
    const rect = this.getCellRect(cell);
    if (!rect) return null;
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  }

  /**
   * Test/e2e-only helper: returns an observable snapshot of the active sheet's drawings (pictures, charts, shapes).
   *
   * This intentionally avoids leaking private overlay internals while still providing stable ids + viewport-relative
   * geometry for Playwright assertions.
   *
   * NOTE: The returned `rectPx` values are in the same coordinate space as `getCellRectA1` (viewport-relative, with
   * header offsets already applied), and use the same anchor-to-rect conversion as the drawings overlay/hit-testing.
   */
  getDrawingsDebugState(): {
    sheetId: string;
    selectedId: number | null;
    drawings: Array<{
      id: number;
      kind: string;
      zOrder: number;
      anchor: unknown;
      rectPx: { x: number; y: number; width: number; height: number } | null;
    }>;
  } {
    // SpreadsheetApp can receive shared-grid viewport notifications before the constructor finishes
    // initializing optional overlays. Keep this API resilient so Playwright can probe early.
    const overlay = (this as any).drawingOverlay as DrawingOverlay | undefined;
    if (!overlay) {
      return { sheetId: this.sheetId, selectedId: null, drawings: [] };
    }

    let viewport: DrawingViewport;
    try {
      viewport = this.getDrawingInteractionViewport(this.sharedGrid ? this.sharedGrid.renderer.scroll.getViewportState() : undefined);
    } catch {
      return { sheetId: this.sheetId, selectedId: null, drawings: [] };
    }

    const objects = this.getDrawingObjects(this.sheetId);
    const drawings = objects.map((obj) => {
      let rectPx: { x: number; y: number; width: number; height: number } | null = null;
      try {
        const rect = drawingObjectToViewportRect(obj, viewport, this.drawingGeom);
        rectPx = { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
      } catch {
        rectPx = null;
      }

      return {
        id: obj.id,
        kind: obj.kind.type,
        zOrder: obj.zOrder,
        anchor: obj.anchor as unknown,
        rectPx,
      };
    });

    return { sheetId: this.sheetId, selectedId: this.getSelectedDrawingId(), drawings };
  }

  /**
   * Test/e2e-only helper: returns the viewport-relative rect for the drawing's anchor bounds.
   *
   * The coordinate space matches `getCellRectA1`.
   */
  getDrawingRectPx(id: number): { x: number; y: number; width: number; height: number } | null {
    const overlay = (this as any).drawingOverlay as DrawingOverlay | undefined;
    if (!overlay) return null;

    const targetId = Number(id);
    if (!Number.isSafeInteger(targetId)) return null;

    const obj = this.getDrawingObjects(this.sheetId).find((o) => o.id === targetId) ?? null;
    if (!obj) return null;

    let viewport: DrawingViewport;
    try {
      viewport = this.getDrawingInteractionViewport(this.sharedGrid ? this.sharedGrid.renderer.scroll.getViewportState() : undefined);
    } catch {
      return null;
    }

    try {
      const rect = drawingObjectToViewportRect(obj, viewport, this.drawingGeom);
      if (
        !Number.isFinite(rect.x) ||
        !Number.isFinite(rect.y) ||
        !Number.isFinite(rect.width) ||
        !Number.isFinite(rect.height) ||
        rect.width <= 0 ||
        rect.height <= 0
      ) {
        return null;
      }
      return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
    } catch {
      return null;
    }
  }

  /**
   * Test/e2e-only helper: returns viewport-relative points for the 4 corner resize handles.
   *
   * These points are intended for Playwright to reliably target resize handles, and match the
   * positions used by the drawings overlay's selection handle rendering.
   */
  getDrawingHandlePointsPx(id: number): {
    nw: { x: number; y: number };
    ne: { x: number; y: number };
    se: { x: number; y: number };
    sw: { x: number; y: number };
  } | null {
    const overlay = (this as any).drawingOverlay as DrawingOverlay | undefined;
    if (!overlay) return null;

    const targetId = Number(id);
    if (!Number.isSafeInteger(targetId)) return null;

    const obj = this.getDrawingObjects(this.sheetId).find((o) => o.id === targetId) ?? null;
    if (!obj) return null;

    let viewport: DrawingViewport;
    try {
      viewport = this.getDrawingInteractionViewport(this.sharedGrid ? this.sharedGrid.renderer.scroll.getViewportState() : undefined);
    } catch {
      return null;
    }

    let bounds: { x: number; y: number; width: number; height: number };
    try {
      bounds = drawingObjectToViewportRect(obj, viewport, this.drawingGeom);
    } catch {
      return null;
    }

    let centers: ReturnType<typeof getResizeHandleCenters>;
    try {
      centers = getResizeHandleCenters(bounds, obj.transform);
    } catch {
      return null;
    }

    const pick = (handle: "nw" | "ne" | "se" | "sw"): { x: number; y: number } | null => {
      const found = centers.find((c) => c.handle === handle);
      if (!found) return null;
      return { x: found.x, y: found.y };
    };

    const nw = pick("nw");
    const ne = pick("ne");
    const se = pick("se");
    const sw = pick("sw");
    if (!nw || !ne || !se || !sw) return null;

    return { nw, ne, se, sw };
  }

  /**
   * Returns grid renderer perf stats (shared grid only).
   */
  getGridPerfStats(): unknown {
    if (!this.sharedGrid) return null;
    const stats = this.sharedGrid.getPerfStats();
    return {
      enabled: stats.enabled,
      lastFrameMs: stats.lastFrameMs,
      cellsPainted: stats.cellsPainted,
      cellFetches: stats.cellFetches,
      dirtyRects: { ...stats.dirtyRects },
      blitUsed: stats.blitUsed
    };
  }

  setGridPerfStatsEnabled(enabled: boolean): void {
    this.sharedGrid?.setPerfStatsEnabled(enabled);
    this.dispatchViewChanged();
  }

  getFrozen(): { frozenRows: number; frozenCols: number } {
    const view = this.document.getSheetView(this.sheetId) as { frozenRows?: number; frozenCols?: number } | null;
    const normalize = (value: unknown, max: number): number => {
      const num = Number(value);
      if (!Number.isFinite(num)) return 0;
      return Math.max(0, Math.min(Math.trunc(num), max));
    };
    return {
      frozenRows: normalize(view?.frozenRows, this.limits.maxRows),
      frozenCols: normalize(view?.frozenCols, this.limits.maxCols),
    };
  }

  private syncFrozenPanes(): void {
    const { frozenRows, frozenCols } = this.getFrozen();
    if (this.sharedGrid) {
      // Shared-grid mode uses frozen panes for headers + user freezes.
      const headerRows = 1;
      const headerCols = 1;
      this.sharedGrid.renderer.setFrozen(headerRows + frozenRows, headerCols + frozenCols);
      this.syncSharedGridAxisSizesFromDocument();
      // Scrollbars + overlays update via DesktopSharedGrid's renderer viewport subscription.
      return;
    }

    const didClamp = this.clampScroll();
    if (didClamp) this.hideCommentTooltip();
    this.syncScrollbars();
    if (didClamp) this.notifyScrollListeners();
    this.refresh();
  }

  private syncActiveSheetBackgroundImage(): void {
    const doc: any = this.document as any;
    const desiredIdRaw =
      typeof doc.getSheetBackgroundImageId === "function"
        ? doc.getSheetBackgroundImageId(this.sheetId)
        : (this.document.getSheetView(this.sheetId) as any)?.backgroundImageId;
    const desiredId = typeof desiredIdRaw === "string" && desiredIdRaw.trim() !== "" ? desiredIdRaw : null;
    if (desiredId === this.activeSheetBackgroundImageId && this.activeSheetBackgroundBitmap) {
      return;
    }

    // If the desired image id hasn't changed and we already have an in-flight decode,
    // keep it running so we don't churn/abort repeatedly on incidental refresh calls.
    if (desiredId === this.activeSheetBackgroundImageId && this.activeSheetBackgroundAbort) {
      return;
    }

    // If the background id changed, clear current state and repaint immediately so stale
    // patterns disappear (even if the new image is still decoding).
    if (desiredId !== this.activeSheetBackgroundImageId) {
      // Cancel any in-flight background decode for the prior image id.
      this.activeSheetBackgroundAbort?.abort();
      this.activeSheetBackgroundAbort = null;

      this.activeSheetBackgroundImageId = desiredId;
      this.activeSheetBackgroundBitmap = null;
      this.activeSheetBackgroundLoadToken += 1;
      if (this.sharedGrid) {
        this.sharedGrid.renderer.setBackgroundPatternImage(null);
      } else {
        this.refresh();
      }
    }

    if (!desiredId) {
      this.activeSheetBackgroundAbort?.abort();
      this.activeSheetBackgroundAbort = null;
      return;
    }

    let entry: ImageEntry | undefined;
    try {
      entry = lookupImageEntry(desiredId, doc.images) ?? normalizeImageEntry(desiredId, doc.getImage?.(desiredId));
    } catch {
      entry = undefined;
    }
    if (!entry) return;

    const token = ++this.activeSheetBackgroundLoadToken;
    this.activeSheetBackgroundAbort?.abort();
    const abort = typeof AbortController !== "undefined" ? new AbortController() : null;
    this.activeSheetBackgroundAbort = abort;
    const signal = abort?.signal;
    const promise = this.workbookImageBitmaps
      .get(entry, signal ? { signal } : undefined)
      .then((bitmap) => {
        if (this.disposed) return;
        if (signal?.aborted) return;
        if (token !== this.activeSheetBackgroundLoadToken) return;
        if (this.activeSheetBackgroundImageId !== desiredId) return;
        this.activeSheetBackgroundBitmap = bitmap;
        if (this.sharedGrid) {
          this.sharedGrid.renderer.setBackgroundPatternImage(bitmap);
        } else {
          this.refresh();
        }
      })
      .catch((err) => {
        if (signal?.aborted || (err as any)?.name === "AbortError") return;
        // Ignore decode failures; treat as "no background image".
        if (token !== this.activeSheetBackgroundLoadToken) return;
        if (this.activeSheetBackgroundImageId !== desiredId) return;
        this.activeSheetBackgroundBitmap = null;
        if (this.sharedGrid) {
          this.sharedGrid.renderer.setBackgroundPatternImage(null);
        } else {
          this.refresh();
        }
      })
      .finally(() => {
        if (this.activeSheetBackgroundAbort === abort) {
          this.activeSheetBackgroundAbort = null;
        }
      });

    // Track in the idle monitor so tests can deterministically await image decode + repaint.
    this.idle.track(promise);
  }

  private handleWorkbookImageDeltasForBackground(payload: any): void {
    if (!payload || typeof payload !== "object") return;

    const desiredId = this.getSheetBackgroundImageId(this.sheetId);
    if (!desiredId) return;

    const deltas: unknown = (payload as any).imageDeltas ?? (payload as any).imagesDeltas;
    if (!Array.isArray(deltas) || deltas.length === 0) return;

    const touched = deltas.some((d) => {
      const anyDelta = d as any;
      const imageId =
        typeof anyDelta?.imageId === "string" ? anyDelta.imageId : typeof anyDelta?.id === "string" ? anyDelta.id : null;
      return typeof imageId === "string" && imageId.trim() === desiredId;
    });
    if (!touched) return;

    // Cancel any in-flight background decode before invalidating the ImageBitmap cache.
    //
    // This avoids a subtle leak: `ImageBitmapCache.invalidate()` drops the cache entry while the
    // underlying `createImageBitmap` promise is still in-flight; if that promise later resolves
    // while a waiter is still attached, the decoded ImageBitmap can escape the cache and never be
    // closed. Aborting first ensures waiters are released so the stale decode result is closed.
    this.activeSheetBackgroundAbort?.abort();
    this.activeSheetBackgroundAbort = null;

    // Force a reload even when the background id itself is unchanged.
    this.workbookImageBitmaps.invalidate(desiredId);
    this.activeSheetBackgroundImageId = null;
    this.activeSheetBackgroundBitmap = null;
    this.syncActiveSheetBackgroundImage();
  }

  private syncSharedGridAxisSizesFromDocument(): void {
    if (!this.sharedGrid) return;

    const view = this.document.getSheetView(this.sheetId) as {
      colWidths?: Record<string, number>;
      rowHeights?: Record<string, number>;
    } | null;

    const zoom = this.sharedGrid.renderer.getZoom();
    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    const { rowCount, colCount } = this.sharedGrid.renderer.scroll.getCounts();
    const maxDocRows = Math.max(0, rowCount - headerRows);
    const maxDocCols = Math.max(0, colCount - headerCols);

    // Shared-grid hide/unhide semantics: only `OutlineEntry.hidden.user` is treated as hidden.
    // (Outline- and filter-hidden rows/cols are intentionally ignored so existing shared-grid
    // outline compatibility tests remain valid.)
    const HIDDEN_AXIS_SIZE_BASE = 2; // CSS px at zoom=1
    const outline = this.getOutlineForSheet(this.sheetId);

    const docColOverridesBase = new Map<number, number>();
    for (const [key, value] of Object.entries(view?.colWidths ?? {})) {
      const col = Number(key);
      if (!Number.isInteger(col) || col < 0) continue;
      if (col >= maxDocCols) continue;
      const size = Number(value);
      if (!Number.isFinite(size) || size <= 0) continue;
      docColOverridesBase.set(col, size);
    }
    for (const [summaryIndex, entry] of outline.cols.entries) {
      if (!entry.hidden.user) continue;
      const docCol = Number(summaryIndex) - 1; // outline indices are 1-based
      if (!Number.isInteger(docCol) || docCol < 0) continue;
      if (docCol >= maxDocCols) continue;
      // Hidden overrides must take precedence over persisted widths.
      docColOverridesBase.set(docCol, HIDDEN_AXIS_SIZE_BASE);
    }

    const docRowOverridesBase = new Map<number, number>();
    for (const [key, value] of Object.entries(view?.rowHeights ?? {})) {
      const row = Number(key);
      if (!Number.isInteger(row) || row < 0) continue;
      if (row >= maxDocRows) continue;
      const size = Number(value);
      if (!Number.isFinite(size) || size <= 0) continue;
      docRowOverridesBase.set(row, size);
    }
    for (const [summaryIndex, entry] of outline.rows.entries) {
      if (!entry.hidden.user) continue;
      const docRow = Number(summaryIndex) - 1; // outline indices are 1-based
      if (!Number.isInteger(docRow) || docRow < 0) continue;
      if (docRow >= maxDocRows) continue;
      docRowOverridesBase.set(docRow, HIDDEN_AXIS_SIZE_BASE);
    }

    // Batch apply to avoid N-per-index invalidation + worst-case O(n^2) `VariableSizeAxis` updates.
    // Ensure indices are applied in ascending order (CanvasGridRenderer expects this for predictable
    // axis updates).
    const colSizes = new Map<number, number>();
    for (let i = 0; i < headerCols; i += 1) {
      colSizes.set(i, this.sharedGrid.renderer.getColWidth(i));
    }
    const docCols = [...docColOverridesBase.keys()].sort((a, b) => a - b);
    for (const docCol of docCols) {
      const base = docColOverridesBase.get(docCol);
      if (base == null) continue;
      colSizes.set(docCol + headerCols, base * zoom);
    }

    const rowSizes = new Map<number, number>();
    for (let i = 0; i < headerRows; i += 1) {
      rowSizes.set(i, this.sharedGrid.renderer.getRowHeight(i));
    }
    const docRows = [...docRowOverridesBase.keys()].sort((a, b) => a - b);
    for (const docRow of docRows) {
      const base = docRowOverridesBase.get(docRow);
      if (base == null) continue;
      rowSizes.set(docRow + headerRows, base * zoom);
    }

    this.sharedGrid.renderer.applyAxisSizeOverrides({ rows: rowSizes, cols: colSizes }, { resetUnspecified: true });

    // `drawingGeom` is stable by reference but reads live shared-grid scroll state.
    // Row/col size overrides change `cellOriginPx` / `cellSizePx`, so any cached
    // drawings bounds must be recomputed.
    const drawingOverlay = (this as any).drawingOverlay as DrawingOverlay | undefined;
    drawingOverlay?.invalidateSpatialIndex();
  }

  freezePanes(): void {
    if (this.isEditing()) return;
    const active = this.selection.active;
    this.document.setFrozen(this.sheetId, active.row, active.col, { label: t("command.view.freezePanes") });
  }

  freezeTopRow(): void {
    if (this.isEditing()) return;
    this.document.setFrozen(this.sheetId, 1, 0, { label: t("command.view.freezeTopRow") });
  }

  freezeFirstColumn(): void {
    if (this.isEditing()) return;
    this.document.setFrozen(this.sheetId, 0, 1, { label: t("command.view.freezeFirstColumn") });
  }

  unfreezePanes(): void {
    if (this.isEditing()) return;
    this.document.setFrozen(this.sheetId, 0, 0, { label: t("command.view.unfreezePanes") });
  }

  addChart(spec: CreateChartSpec): CreateChartResult {
    return this.chartStore.createChart(spec);
  }

  setChartTheme(theme: ChartTheme): void {
    this.chartTheme = theme;
    // Keep imported chart rendering aligned with the workbook palette too.
    this.formulaChartModelStore.setDefaultTheme({ seriesColors: theme.seriesColors });
    // Imported charts are rendered via the drawings overlay; repaint so the new palette applies.
    this.renderDrawings();
    this.dispatchDrawingsChanged();
    this.renderCharts(false);
  }

  /**
   * Populate the imported-chart model store from JSON-serialized Rust `ChartModel`s.
   *
   * Expected input shape (best-effort; callers may omit fields):
   * - `{ chart_id: string, rel_id?: string, model: unknown }` (snake_case; Tauri)
   * - `{ chartId: string, relId?: string, model: unknown }` (camelCase)
   */
  setImportedChartModels(entries: unknown): void {
    this.formulaChartModelStore.clear();

    if (!Array.isArray(entries)) {
      this.renderDrawings();
      this.dispatchDrawingsChanged();
      return;
    }

    const sheetIdByName = new Map<string, string>();
    for (const id of this.document.getSheetIds()) {
      const meta = this.document.getSheetMeta(id);
      const name = typeof meta?.name === "string" ? meta.name.trim() : "";
      if (!name) continue;
      // Excel sheet names are effectively case-insensitive; normalize for matching.
      if (!sheetIdByName.has(name.toLowerCase())) {
        sheetIdByName.set(name.toLowerCase(), id);
      }
    }

    for (const entry of entries) {
      if (!entry || typeof entry !== "object") continue;
      const e = entry as any;
      const model = e.model;
      if (model == null) continue;

      let chartId: string | null =
        typeof e.chart_id === "string"
          ? e.chart_id
          : typeof e.chartId === "string"
            ? e.chartId
            : null;

      const sheetName =
        typeof e.sheet_name === "string" ? e.sheet_name : typeof e.sheetName === "string" ? e.sheetName : null;
      const drawingObjectId =
        typeof e.drawing_object_id === "number"
          ? e.drawing_object_id
          : typeof e.drawingObjectId === "number"
            ? e.drawingObjectId
            : null;

      // Some import paths expose charts with a stable `${sheetId}:${drawingObjectId}` key, while
      // others use `${sheetName}:${drawingObjectId}`. Keep both keys in the model store so drawing
      // adapters can use either identifier.
      const sheetObjectChartId =
        sheetName && drawingObjectId != null ? FormulaChartModelStore.chartIdFromSheetObject(sheetName, drawingObjectId) : null;

      const stableSheetId =
        sheetName && sheetName.trim() !== "" ? (sheetIdByName.get(sheetName.trim().toLowerCase()) ?? null) : null;
      const stableChartId =
        stableSheetId && drawingObjectId != null ? FormulaChartModelStore.chartIdFromSheetObject(stableSheetId, drawingObjectId) : null;

      if (!chartId && sheetObjectChartId) {
        chartId = sheetObjectChartId;
      }

      const relId: string | null =
        typeof e.rel_id === "string"
          ? e.rel_id
          : typeof e.relId === "string"
            ? e.relId
            : null;

      try {
        if (chartId && chartId.trim() !== "") {
          this.formulaChartModelStore.setFormulaModelChartModel(chartId, model);
        }
        if (stableChartId && stableChartId.trim() !== "" && stableChartId !== chartId) {
          this.formulaChartModelStore.setFormulaModelChartModel(stableChartId, model);
        }
        if (sheetObjectChartId && sheetObjectChartId.trim() !== "" && sheetObjectChartId !== chartId) {
          this.formulaChartModelStore.setFormulaModelChartModel(sheetObjectChartId, model);
        }
        // Back-compat: some drawing adapters may identify charts by the drawing relationship id
        // (`rId*`) when sheet/object context isn't available.
        if (relId && relId.trim() !== "" && relId !== chartId) {
          this.formulaChartModelStore.setFormulaModelChartModel(relId, model);
        }
      } catch {
        // Best-effort: ignore malformed chart models so other charts still render.
      }
    }

    // Re-render the drawings overlay so chart placeholders can upgrade to real charts.
    this.renderDrawings();
    this.dispatchDrawingsChanged();
  }

  listCharts(): readonly ChartRecord[] {
    return this.chartStore.listCharts();
  }

  getChartViewportRect(chartId: string): { left: number; top: number; width: number; height: number } | null {
    const chart = this.getChartRecordById(chartId);
    if (!chart) return null;
    return this.chartAnchorToViewportRect(chart.anchor);
  }

  getSelectedChartId(): string | null {
    return this.selectedChartId;
  }

  async insertPicturesFromFiles(files: File[], opts?: { placeAt?: CellCoord }): Promise<void> {
    if (this.isReadOnly()) {
      throw new Error("Workbook is read-only.");
    }
    if (this.isEditing()) {
      throw new Error("Finish editing before inserting a picture.");
    }

    const normalized = Array.isArray(files)
      ? (files.filter((file) => file && typeof (file as File).name === "string") as File[])
      : [];
    if (normalized.length === 0) return;

    const oversized: File[] = [];
    const accepted: File[] = [];
    for (const file of normalized) {
      const size = typeof (file as any)?.size === "number" ? (file as any).size : null;
      // If we can't determine size, reject rather than risk allocating huge buffers.
      if (size == null || size > MAX_INSERT_IMAGE_BYTES) oversized.push(file);
      else accepted.push(file);
    }

    const toastSkippedOversized = () => {
      if (oversized.length === 0) return;
      const mb = Math.round(MAX_INSERT_IMAGE_BYTES / 1024 / 1024);
      const message =
        oversized.length === 1
          ? `Image too large (>${mb}MB). Choose a smaller file.`
          : `Skipped ${oversized.length} images larger than ${mb}MB: ${oversized
              .map((f) => (typeof f?.name === "string" && f.name.trim() ? f.name.trim() : "unnamed"))
              .join(", ")}`;
      try {
        showToast(message, "warning");
      } catch {
        // `showToast` requires a #toast-root; unit tests don't always include it.
      }
    };

    if (accepted.length === 0) {
      toastSkippedOversized();
      return;
    }

    const guessMimeType = (name: string): string => {
      const ext = String(name ?? "").split(".").pop()?.toLowerCase();
      switch (ext) {
        case "png":
          return "image/png";
        case "jpg":
        case "jpeg":
          return "image/jpeg";
        case "gif":
          return "image/gif";
        case "bmp":
          return "image/bmp";
        case "webp":
          return "image/webp";
        case "svg":
          return "image/svg+xml";
        default:
          return "application/octet-stream";
      }
    };

    const uuid = (): string => {
      const randomUuid = (globalThis as any).crypto?.randomUUID as (() => string) | undefined;
      if (typeof randomUuid === "function") {
        try {
          return randomUuid.call((globalThis as any).crypto);
        } catch {
          // Fall through to pseudo-random below.
        }
      }
      return `${Date.now().toString(16)}_${Math.random().toString(16).slice(2)}`;
    };

    const placeAt = opts?.placeAt ?? this.selection.active;
    const anchorCell = { row: placeAt.row, col: placeAt.col };

    // Use the visible grid area (cell canvas) to bound inserted image size.
    const viewport = this.getDrawingRenderViewport(this.sharedGrid?.renderer.scroll.getViewportState?.());
    const cellAreaW = Math.max(1, viewport.width - (viewport.headerOffsetX ?? 0));
    const cellAreaH = Math.max(1, viewport.height - (viewport.headerOffsetY ?? 0));
    const maxW = cellAreaW * 0.6;
    const maxH = cellAreaH * 0.6;
    const zoom = Number.isFinite(viewport.zoom) && (viewport.zoom as number) > 0 ? (viewport.zoom as number) : 1;

    const existingObjects = this.listDrawingObjectsForSheet(this.sheetId);

    // Ensure new pictures stack on top of existing drawings.
    let nextZOrder = 0;
    for (const obj of existingObjects) {
      const z = Number(obj.zOrder);
      if (Number.isFinite(z) && z >= nextZOrder) nextZOrder = z + 1;
    }

    const docAny = this.document as any;
    if (typeof docAny.getSheetDrawings !== "function" || typeof docAny.setSheetDrawings !== "function") {
      throw new Error("Picture insertion is not supported in this build.");
    }

    const existingDrawings = (() => {
      try {
        const raw = docAny.getSheetDrawings(this.sheetId);
        return Array.isArray(raw) ? raw : [];
      } catch {
        return [];
      }
    })();

    const readFileBytes = async (file: File): Promise<Uint8Array> => {
      // Preferred path: the standard File/Blob `arrayBuffer()` API.
      if (typeof (file as any)?.arrayBuffer === "function") {
        return new Uint8Array(await file.arrayBuffer());
      }

      // JSDOM (and some older environments) may not implement `File.arrayBuffer`.
      // Fall back to FileReader when available.
      if (typeof (globalThis as any).FileReader === "function") {
        const reader = new (globalThis as any).FileReader() as FileReader;
        const buf = await new Promise<ArrayBuffer>((resolve, reject) => {
          reader.onerror = () => reject(reader.error ?? new Error("FileReader failed"));
          reader.onload = () => resolve(reader.result as ArrayBuffer);
          reader.readAsArrayBuffer(file);
        });
        return new Uint8Array(buf);
      }

      // Last resort: try the Fetch API's Body helpers.
      if (typeof (globalThis as any).Response === "function") {
        const buf = await new (globalThis as any).Response(file).arrayBuffer();
        return new Uint8Array(buf);
      }

      throw new Error("Reading file bytes is not supported in this environment.");
    };

    const decodeImagePixelSizeViaImage = async (file: File): Promise<{ width: number; height: number } | null> => {
      const ua = typeof navigator !== "undefined" ? String(navigator.userAgent ?? "") : "";
      if (/jsdom|happy-dom/i.test(ua)) return null;

      if (typeof Image === "undefined") return null;
      if (typeof URL === "undefined" || typeof (URL as any).createObjectURL !== "function") return null;

      const url = (URL as any).createObjectURL(file) as string;
      try {
        const img = new Image();
        const loadPromise = new Promise<void>((resolve, reject) => {
          img.onload = () => resolve();
          img.onerror = () => reject(new Error("Image decode failed"));
        });
        img.src = url;

        const timeoutMs = 5_000;
        const timeout = new Promise<void>((_resolve, reject) => {
          setTimeout(() => reject(new Error("Image decode timed out")), timeoutMs);
        });

        const decodePromise = typeof (img as any).decode === "function" ? (img as any).decode() : loadPromise;
        await Promise.race([decodePromise, loadPromise, timeout]);

        const width = Number((img as any).naturalWidth ?? (img as any).width ?? 0);
        const height = Number((img as any).naturalHeight ?? (img as any).height ?? 0);
        if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
          return { width, height };
        }
        return null;
      } catch {
        return null;
      } finally {
        try {
          (URL as any).revokeObjectURL?.(url);
        } catch {
          // ignore
        }
      }
    };

    const MAX_CONCURRENT_DECODES = 4;

    const prepared = await mapWithConcurrencyLimit(accepted, MAX_CONCURRENT_DECODES, async (file, i) => {
      const bytes = await readFileBytes(file);
      if (bytes.byteLength > MAX_INSERT_IMAGE_BYTES) {
        throw new Error(`File is too large (${bytes.byteLength} bytes, max ${MAX_INSERT_IMAGE_BYTES}).`);
      }

      const mimeType = file.type && file.type.trim() ? file.type : guessMimeType(file.name);
      const ext = (() => {
        const raw = String(file.name ?? "").split(".").pop()?.toLowerCase();
        return raw && raw !== file.name ? raw : null;
      })();
      const imageId = `image_${uuid()}${ext ? `.${ext}` : ""}`;

      const imageEntry: ImageEntry = { id: imageId, bytes, mimeType };

      const decoded = await (async (): Promise<{ width: number; height: number } | null> => {
        if (typeof createImageBitmap === "function") {
          try {
            const bitmap = await this.drawingOverlay.preloadImage(imageEntry);
            const width = Number((bitmap as any)?.width);
            const height = Number((bitmap as any)?.height);
            if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
              return { width, height };
            }
          } catch {
            // ignore
          }
        }
        return decodeImagePixelSizeViaImage(file);
      })();

      const fallback = { width: 320, height: 240 };
      const rawW = decoded?.width ?? fallback.width;
      const rawH = decoded?.height ?? fallback.height;
      const widthPx = typeof rawW === "number" && Number.isFinite(rawW) && rawW > 0 ? rawW : fallback.width;
      const heightPx = typeof rawH === "number" && Number.isFinite(rawH) && rawH > 0 ? rawH : fallback.height;

      const scale = Math.min(1, maxW / widthPx, maxH / heightPx);
      const targetScreenW = Math.max(1, widthPx * scale);
      const targetScreenH = Math.max(1, heightPx * scale);

      // Store anchors in base (unzoomed) EMUs so render-time zoom scaling produces the desired
      // on-screen size.
      const targetBaseW = targetScreenW / zoom;
      const targetBaseH = targetScreenH / zoom;

      const offsetScreenPx = 16 * i;
      const offsetBasePx = offsetScreenPx / zoom;

      const anchor: DrawingAnchor = {
        type: "oneCell",
        from: {
          cell: anchorCell,
          offset: { xEmu: pxToEmu(offsetBasePx), yEmu: pxToEmu(offsetBasePx) },
        },
        size: { cx: pxToEmu(targetBaseW), cy: pxToEmu(targetBaseH) },
      };

      const drawingId = createDrawingObjectId();
      const drawing: DrawingObject = {
        id: drawingId,
        kind: { type: "image", imageId },
        anchor,
        zOrder: nextZOrder + i,
        size: anchor.size,
      };

      return { imageEntry, drawing };
    });

    this.document.beginBatch({ label: "Insert Picture" });
    try {
      for (const entry of prepared) {
        // Persist picture bytes out-of-band (IndexedDB) so they survive reloads without
        // bloating DocumentController snapshot payloads.
        this.drawingImages.set(entry.imageEntry);
        try {
          this.imageBytesBinder?.onLocalImageInserted(entry.imageEntry);
        } catch {
          // Best-effort: never fail picture insertion due to collab image propagation.
        }
      }

      // Store drawing ids as strings in the DocumentController payload for back-compat with
      // integrations that treat drawing ids as JSON-friendly keys. The UI continues to use
      // numeric ids (see `DrawingObject.id`).
      const nextDrawings = [
        ...existingDrawings,
        ...prepared.map((p) => ({ ...(p.drawing as any), id: String(p.drawing.id) })),
      ];
      docAny.setSheetDrawings(this.sheetId, nextDrawings, { label: "Insert Picture" });
      this.document.endBatch();
      toastSkippedOversized();
    } catch (err) {
      this.document.cancelBatch();
      throw err;
    }

    if (prepared.length > 0) {
      const insertedObjects = prepared.map((p) => p.drawing);
      const lastInsertedId = insertedObjects[insertedObjects.length - 1]!.id;
      this.setDrawingObjectsForSheet([...existingObjects, ...insertedObjects]);
      const prevSelected = this.selectedDrawingId;
      this.selectedDrawingId = lastInsertedId;
      this.drawingOverlay.setSelectedId(lastInsertedId);
      if (this.gridMode === "shared") {
        this.ensureDrawingInteractionController().setSelectedId(lastInsertedId);
      }
      if (prevSelected !== lastInsertedId) {
        this.dispatchDrawingSelectionChanged();
      }
    }

    // Ensure the drawings overlay is up-to-date after the batch completes.
    this.renderDrawings();
    this.renderSelection();
    this.focus();
  }


  private arrangeSelectedDrawing(direction: "forward" | "backward" | "front" | "back"): void {
    const sheetId = this.sheetId;
    const selectedId = this.selectedDrawingId;
    if (selectedId == null) return;
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;

    const docAny = this.document as any;
    if (typeof docAny.getSheetDrawings !== "function" || typeof docAny.setSheetDrawings !== "function") {
      return;
    }

    let drawingsRaw: unknown = null;
    try {
      drawingsRaw = docAny.getSheetDrawings(sheetId);
    } catch {
      drawingsRaw = null;
    }
    const drawings = Array.isArray(drawingsRaw) ? drawingsRaw : [];
    if (drawings.length < 2) return;

    const ordered = drawings
      .map((drawing, idx) => ({
        drawing,
        idx,
        z: Number((drawing as any)?.zOrder ?? (drawing as any)?.z_order ?? idx),
        idKey: String((drawing as any)?.id ?? ""),
      }))
      .sort((a, b) => {
        const diff = a.z - b.z;
        if (diff !== 0) return diff;
        return a.idx - b.idx;
      });

    const selectedKey = String(selectedId);
    let index = ordered.findIndex((entry) => entry.idKey === selectedKey);
    if (index === -1) {
      // Drawings may be stored with non-numeric ids (e.g. historical snapshots). In that case,
      // `selectedDrawingId` is a stable numeric mapping produced by the adapter layer. Resolve the
      // selected raw entry by comparing via `convertDocumentSheetDrawingsToUiDrawingObjects`.
      for (let i = 0; i < ordered.length; i += 1) {
        const entry = ordered[i]!;
        let uiId: number | null = null;
        try {
          uiId = convertDocumentSheetDrawingsToUiDrawingObjects([entry.drawing], { sheetId })[0]?.id ?? null;
        } catch {
          uiId = null;
        }
        if (uiId === selectedId) {
          index = i;
          break;
        }
      }
    }
    if (index === -1) return;

    let nextOrder = ordered;
    if (direction === "forward") {
      if (index >= ordered.length - 1) return;
      nextOrder = ordered.slice();
      const tmp = nextOrder[index]!;
      nextOrder[index] = nextOrder[index + 1]!;
      nextOrder[index + 1] = tmp;
    } else if (direction === "backward") {
      if (index <= 0) return;
      nextOrder = ordered.slice();
      const tmp = nextOrder[index]!;
      nextOrder[index] = nextOrder[index - 1]!;
      nextOrder[index - 1] = tmp;
    } else if (direction === "front") {
      if (index >= ordered.length - 1) return;
      nextOrder = ordered.slice();
      const [entry] = nextOrder.splice(index, 1);
      if (!entry) return;
      nextOrder.push(entry);
    } else {
      if (index <= 0) return;
      nextOrder = ordered.slice();
      const [entry] = nextOrder.splice(index, 1);
      if (!entry) return;
      nextOrder.unshift(entry);
    }

    // Renormalize zOrder to match render order (0..n-1) and keep the stored array ordered
    // by zOrder so the overlay/hit-test layers can avoid per-frame sorting.
    const next = nextOrder.map(({ drawing }, zOrder) => {
      const cloned = { ...(drawing as any), zOrder } as any;
      if ("z_order" in cloned) delete cloned.z_order;
      return cloned;
    });

    const label =
      direction === "forward"
        ? "Bring Forward"
        : direction === "backward"
          ? "Send Backward"
          : direction === "front"
            ? "Bring To Front"
            : "Send To Back";

    this.document.beginBatch({ label });
    try {
      docAny.setSheetDrawings(sheetId, next);
      this.document.endBatch();
    } catch (err) {
      this.document.cancelBatch();
      throw err;
    }
  }

  private getChartRecordById(chartId: string): ChartRecord | undefined {
    const id = String(chartId ?? "");
    if (!id) return undefined;
    const list = this.chartStore.listCharts();
    const cached = this.chartRecordLookupCache;
    if (!cached || cached.list !== list) {
      const map = new Map<string, ChartRecord>();
      for (const chart of list) map.set(chart.id, chart);
      this.chartRecordLookupCache = { list, map };
      return map.get(id);
    }
    return cached.map.get(id);
  }

  private enqueueWasmSync(task: (engine: EngineClient) => Promise<void>): Promise<void> {
    const engine = this.wasmEngine;
    if (!engine) return Promise.resolve();

    const run = async () => {
      // The engine may have been replaced/terminated while the task was queued.
      if (this.wasmEngine !== engine) return;
      await task(engine);
    };

    this.wasmSyncPromise = this.wasmSyncPromise
      .catch(() => {
        // Ignore prior errors so the chain keeps flowing.
      })
      .then(run)
      .catch(() => {
        // Ignore WASM sync failures; the DocumentController remains the source of truth.
      });

    return this.wasmSyncPromise;
  }

  /**
   * Replace the DocumentController state from a snapshot, then hydrate the WASM engine in one step.
   *
   * This avoids N-per-cell RPC roundtrips during version restore by using the engine JSON load path.
   */
  async restoreDocumentState(snapshot: Uint8Array): Promise<void> {
    this.wasmSyncSuspended = true;
    try {
      // Ensure any in-flight sync operations finish before we replace the workbook.
      await this.wasmSyncPromise;
      this.clearComputedValuesByCoord();
      this.document.applyState(snapshot);
      // The DocumentController snapshot format can include workbook-scoped image bytes
      // (`snapshot.images`). Keep the UI-level in-cell image store aligned with the
      // newly-restored workbook so `CellValue::Image` references can resolve.
      this.syncInCellImageStoreFromDocument();
      // `applyState` can replace workbook-scoped image bytes. Clear decoded bitmap caches so we
      // don't show stale images for reused ids across workbooks/versions.
      this.activeSheetBackgroundAbort?.abort();
      this.activeSheetBackgroundAbort = null;
      this.workbookImageBitmaps.clear();
      this.activeSheetBackgroundImageId = null;
      this.activeSheetBackgroundBitmap = null;
      // ImageBitmap caches also live inside individual CanvasGridRenderer instances (shared-grid
      // mode). Clear them so we don't show stale in-cell images for reused ids across versions.
      this.sharedGrid?.renderer?.clearImageCache?.();
      let sheetChanged = false;
      const sheetIds = this.document.getSheetIds();
      if (sheetIds.length > 0 && !sheetIds.includes(this.sheetId)) {
        this.sheetId = sheetIds[0];
        sheetChanged = true;
        this.reindexCommentCells();
        this.chartStore.setDefaultSheet(this.sheetId);
      }
      this.syncActiveSheetBackgroundImage();
      // In collab mode, comment indicators/tooltips are indexed per active sheet (to avoid
      // collisions between `SheetA!A1` and `SheetB!A1`). If a restore changes the active sheet
      // (or just changes the sheet list used for legacy comment fallback), ensure the comment
      // indexes track the new active sheet.
      if (this.collabMode) {
        this.reindexCommentCells();
      }
      this.referencePreview = null;
      this.referenceHighlights = this.computeReferenceHighlightsForSheet(this.sheetId, this.referenceHighlightsSource);
      if (this.sharedGrid) this.syncSharedGridReferenceHighlights();
      if (sheetChanged) {
        // Row/col visibility (outline hidden rows/cols) is sheet-local in the legacy renderer.
        // If a restore changed the active sheet id, rebuild the visibility caches before any redraw.
        this.rebuildAxisVisibilityCache();
      }
      this.syncFrozenPanes();
      if (this.wasmEngine) {
        await this.enqueueWasmSync(async (engine) => {
          const changes = await engineHydrateFromDocument(engine, this.document);
          this.applyComputedChanges(changes);
        });
      }
      const presence = this.collabSession?.presence;
      if (presence) {
        // Ensure presence updates are always associated with the correct sheet.
        // Note: `setCursor` / `setSelections` may broadcast immediately (throttleMs=0),
        // so update `activeSheet` first.
        presence.setActiveSheet(this.sheetId);
        presence.setCursor(this.selection.active);
        presence.setSelections(this.selection.ranges);
      }
    } finally {
      this.wasmSyncSuspended = false;
    }
  }

  private syncInCellImageStoreFromDocument(): void {
    this.imageStore.clear();
    const doc: any = this.document as any;
    const images: unknown = doc?.images;
    if (!(images instanceof Map)) return;

    for (const [id, raw] of images.entries()) {
      const imageId = typeof id === "string" ? id : String(id ?? "");
      if (!imageId) continue;
      if (!raw || typeof raw !== "object") continue;
      const entry = raw as any;
      const bytes: unknown = entry.bytes;
      if (!(bytes instanceof Uint8Array)) continue;

      const mimeTypeRaw: unknown = entry.mimeType;
      const mimeType = typeof mimeTypeRaw === "string" && mimeTypeRaw.trim() !== "" ? mimeTypeRaw : "application/octet-stream";

      this.imageStore.set(imageId, { bytes, mimeType });
    }
  }

  private async initWasmEngine(): Promise<void> {
    if (this.wasmEngine) return;
    if (typeof Worker === "undefined") return;

    let postInitHydrate: Promise<void> | null = null;

    // Include WASM initialization in the idle tracker. This ensures e2e tests (and any UI that
    // awaits `whenIdle()`) don't race the initial engine hydration.
    const prior = this.wasmSyncPromise;
    this.wasmSyncPromise = prior
      .catch(() => {})
      .then(async () => {
        // Another init call may have completed while we were waiting on the prior promise.
        if (this.wasmEngine) return;

        const env = (import.meta as any)?.env as Record<string, unknown> | undefined;
        const wasmModuleUrl =
          typeof env?.VITE_FORMULA_WASM_MODULE_URL === "string" ? env.VITE_FORMULA_WASM_MODULE_URL : undefined;
        const wasmBinaryUrl =
          typeof env?.VITE_FORMULA_WASM_BINARY_URL === "string" ? env.VITE_FORMULA_WASM_BINARY_URL : undefined;

        let engine: EngineClient | null = null;
        try {
          engine = createEngineClient({ wasmModuleUrl, wasmBinaryUrl });
          let changedDuringInit = false;
          const unsubscribeInit = this.document.on("change", () => {
            changedDuringInit = true;
          });
          try {
            await engine.init();

            // `engineHydrateFromDocument` is relatively expensive, but we need to ensure we don't miss
            // edits that happen while the WASM worker is booting. If the user edits the document
            // between hydrating the worker and subscribing to incremental deltas, the engine can get
            // permanently out of sync (until a full reload). Track any DocumentController changes
            // during the hydrate step and retry once to converge.
            //
            // This is especially important for fast e2e runs that start editing immediately after
            // navigation.
            for (let attempt = 0; attempt < 2; attempt += 1) {
              changedDuringInit = false;
              this.clearComputedValuesByCoord();
              const changes = await engineHydrateFromDocument(engine, this.document);
              this.applyComputedChanges(changes);
              if (!changedDuringInit) break;
            }
          } finally {
            unsubscribeInit();
          }

          this.wasmEngine = engine;
          this.wasmUnsubscribe = this.document.on("change", (payload: any) => {
            if (!this.wasmEngine || this.wasmSyncSuspended) return;

            const source = typeof payload?.source === "string" ? payload.source : "";

            if (source === "applyState") {
              this.clearComputedValuesByCoord();
              void this.enqueueWasmSync(async (worker) => {
                const changes = await engineHydrateFromDocument(worker, this.document);
                this.applyComputedChanges(changes);
              });
              return;
            }

            const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
            const rowStyleDeltas = Array.isArray(payload?.rowStyleDeltas) ? payload.rowStyleDeltas : [];
            const colStyleDeltas = Array.isArray(payload?.colStyleDeltas) ? payload.colStyleDeltas : [];
            const sheetStyleDeltas = Array.isArray(payload?.sheetStyleDeltas) ? payload.sheetStyleDeltas : [];
            const sheetViewDeltas = Array.isArray(payload?.sheetViewDeltas) ? payload.sheetViewDeltas : [];
            const hasStyles = rowStyleDeltas.length > 0 || colStyleDeltas.length > 0 || sheetStyleDeltas.length > 0;
            const hasViews = sheetViewDeltas.length > 0;

            const recalc = payload?.recalc;
            const wantsRecalc = recalc === true;

            // Formatting-only / view-only payloads often omit cell deltas. Avoid scheduling a WASM
            // task unless the payload can impact calculation results.
            if (deltas.length === 0 && !hasStyles && !hasViews) {
              if (wantsRecalc) {
                void this.enqueueWasmSync(async (worker) => {
                  const changes = await worker.recalculate();
                  this.applyComputedChanges(changes);
                });
              }
              return;
            }

            void this.enqueueWasmSync(async (worker) => {
              const changes = await engineApplyDocumentChange(worker, payload, {
                getStyleById: (styleId) => (this.document as any)?.styleTable?.get?.(styleId),
              });
              this.applyComputedChanges(changes);
            });
          });

          // `initWasmEngine` runs asynchronously and can overlap with early user edits (or e2e
          // interactions) before the `document.on("change")` listener is installed. If the
          // DocumentController changed while `engineHydrateFromDocument` was in-flight, those
          // deltas could be missed, leaving the worker with an incomplete view of inputs.
          //
          // Re-hydrate once through the serialized WASM queue to guarantee the worker matches the
          // latest DocumentController state before incremental deltas begin flowing.
          //
          // Note: do not `await` inside this init chain (it would deadlock by waiting on the
          // promise chain we're currently building).
          postInitHydrate = this.enqueueWasmSync(async (worker) => {
            const changes = await engineHydrateFromDocument(worker, this.document);
            this.applyComputedChanges(changes);
          });
          } catch {
            // Ignore initialization failures (e.g. missing WASM bundle).
            engine?.terminate();
            this.wasmEngine = null;
            this.wasmUnsubscribe?.();
            this.wasmUnsubscribe = null;
            this.clearComputedValuesByCoord();
          }
      })
      .catch(() => {
        // Ignore WASM init failures; the app continues to function using the in-process mock engine.
      });

    await this.wasmSyncPromise;
    if (postInitHydrate) {
      try {
        await postInitHydrate;
      } catch {
        // ignore
      }
    }
  }

  /**
   * Switch the active sheet id and re-render.
   */
  activateSheet(sheetId: string): void {
    if (!sheetId) return;
    if (sheetId === this.sheetId) return;
    this.sheetId = sheetId;
    this.drawingObjectsCache = null;
    this.drawingHitTestIndex = null;
    this.drawingHitTestIndexObjects = null;
    this.selectedDrawingId = null;
    this.syncActiveSheetBackgroundImage();
    if (this.collabMode) this.reindexCommentCells();
    const presence = this.collabSession?.presence;
    if (presence) {
      presence.setActiveSheet(this.sheetId);
      presence.setCursor(this.selection.active);
      presence.setSelections(this.selection.ranges);
    }
    this.chartStore.setDefaultSheet(sheetId);
    this.referencePreview = null;
    this.referenceHighlights = this.computeReferenceHighlightsForSheet(this.sheetId, this.referenceHighlightsSource);
    if (this.sharedGrid) this.syncSharedGridReferenceHighlights();
    if (this.sharedGrid) {
      const { frozenRows, frozenCols } = this.getFrozen();
      const headerRows = 1;
      const headerCols = 1;
      this.sharedGrid.renderer.setFrozen(headerRows + frozenRows, headerCols + frozenCols);
      this.syncSharedGridAxisSizesFromDocument();
      this.sharedGrid.scrollTo(this.scrollX, this.scrollY);
    } else {
      // Row/col visibility is sheet-local; when switching sheets in the legacy renderer,
      // rebuild the visibility caches for the newly active sheet.
      this.rebuildAxisVisibilityCache();
    }
    this.renderGrid();
    this.renderDrawings();
    this.renderCharts(true);
    this.renderReferencePreview();
    if (this.sharedGrid) {
      // Switching sheets updates the provider data source but does not emit document
      // changes. Force a full redraw so the CanvasGridRenderer pulls from the new
      // sheet's data.
      this.sharedProvider?.invalidateAll();
      this.syncSharedGridSelectionFromState();
    }
    this.renderReferencePreview();
    this.renderSelection();
    this.updateStatus();
    // Sheet switches can leave formula-bar hover tooltips/overlays in a stale state (e.g. the user was
    // hovering a ref span or editing a formula referencing a different sheet). Clear any visible tooltip
    // and re-sync overlays from the formula bar's current hover state.
    this.hideFormulaRangePreviewTooltip();
    this.syncFormulaBarHoverRangeOverlays();
  }

  /**
   * Programmatically set the active cell (and optionally change sheets).
   */
  activateCell(target: { sheetId?: string; row: number; col: number }, options?: { scrollIntoView?: boolean; focus?: boolean }): void {
    const scrollIntoView = options?.scrollIntoView !== false;
    const focus = options?.focus !== false;
    let sheetChanged = false;
    if (target.sheetId && target.sheetId !== this.sheetId) {
      this.sheetId = target.sheetId;
      this.drawingObjectsCache = null;
      this.drawingHitTestIndex = null;
      this.drawingHitTestIndexObjects = null;
      this.selectedDrawingId = null;
      this.syncActiveSheetBackgroundImage();
      if (this.collabMode) this.reindexCommentCells();
      this.collabSession?.presence?.setActiveSheet(this.sheetId);
      this.chartStore.setDefaultSheet(target.sheetId);
      this.referencePreview = null;
      this.referenceHighlights = this.computeReferenceHighlightsForSheet(this.sheetId, this.referenceHighlightsSource);
      if (this.sharedGrid) this.syncSharedGridReferenceHighlights();
      if (this.sharedGrid) {
        const { frozenRows, frozenCols } = this.getFrozen();
        const headerRows = 1;
        const headerCols = 1;
        this.sharedGrid.renderer.setFrozen(headerRows + frozenRows, headerCols + frozenCols);
        this.syncSharedGridAxisSizesFromDocument();
        this.sharedGrid.scrollTo(this.scrollX, this.scrollY);
      } else {
        this.rebuildAxisVisibilityCache();
      }
      this.renderGrid();
      this.renderCharts(true);
      this.sharedProvider?.invalidateAll();
      sheetChanged = true;
    }
    this.selection = setActiveCell(this.selection, { row: target.row, col: target.col }, this.limits);
    let didScroll = false;
    if (scrollIntoView) {
      this.ensureActiveCellVisible();
      didScroll = this.scrollCellIntoView(this.selection.active);
    }
    // In shared-grid mode, we explicitly scroll the active cell into view above. Avoid triggering
    // a redundant `scrollToCell` (and extra scroll event) when syncing selection ranges.
    if (this.sharedGrid) this.syncSharedGridSelectionFromState({ scrollIntoView: false });
    else if (didScroll) this.ensureViewportMappingCurrent();
    if (sheetChanged) {
      const presence = this.collabSession?.presence;
      if (presence) {
        presence.setCursor(this.selection.active);
        presence.setSelections(this.selection.ranges);
      }
    }
    this.renderSelection();
    this.updateStatus();
    if (sheetChanged) {
      this.hideFormulaRangePreviewTooltip();
      this.syncFormulaBarHoverRangeOverlays();
    }
    if (sheetChanged) {
      // Sheet changes always require a full redraw (grid + charts may differ).
      this.refresh();
    } else if (didScroll) {
      this.refresh("scroll");
    }
    if (focus) this.focus();
  }

  /**
   * Programmatically set the selection range (and optionally change sheets).
   */
  selectRange(target: { sheetId?: string; range: Range }, options?: { scrollIntoView?: boolean; focus?: boolean }): void {
    const scrollIntoView = options?.scrollIntoView !== false;
    const focus = options?.focus !== false;
    let sheetChanged = false;
    if (target.sheetId && target.sheetId !== this.sheetId) {
      this.sheetId = target.sheetId;
      this.drawingObjectsCache = null;
      this.drawingHitTestIndex = null;
      this.drawingHitTestIndexObjects = null;
      this.selectedDrawingId = null;
      this.syncActiveSheetBackgroundImage();
      if (this.collabMode) this.reindexCommentCells();
      this.collabSession?.presence?.setActiveSheet(this.sheetId);
      this.chartStore.setDefaultSheet(target.sheetId);
      this.referencePreview = null;
      this.referenceHighlights = this.computeReferenceHighlightsForSheet(this.sheetId, this.referenceHighlightsSource);
      if (this.sharedGrid) this.syncSharedGridReferenceHighlights();
      if (this.sharedGrid) {
        const { frozenRows, frozenCols } = this.getFrozen();
        const headerRows = 1;
        const headerCols = 1;
        this.sharedGrid.renderer.setFrozen(headerRows + frozenRows, headerCols + frozenCols);
        this.sharedGrid.scrollTo(this.scrollX, this.scrollY);
      } else {
        this.rebuildAxisVisibilityCache();
      }
      this.renderGrid();
      this.renderCharts(true);
      this.sharedProvider?.invalidateAll();
      sheetChanged = true;
    }
    const active = { row: target.range.startRow, col: target.range.startCol };
    this.selection = buildSelection(
      { ranges: [target.range], active, anchor: active, activeRangeIndex: 0 },
      this.limits
    );
    let didScroll = false;
    if (scrollIntoView) {
      this.ensureActiveCellVisible();
      const activeRange = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0];
      const didScrollRange = activeRange ? this.scrollRangeIntoView(activeRange) : false;
      // Even if the range is too large to fit in the viewport, the active cell should never
      // become "lost" offscreen.
      const didScrollCell = this.scrollCellIntoView(this.selection.active);
      didScroll = didScrollRange || didScrollCell;
    }
    // In shared-grid mode, we explicitly scroll the active cell/range into view above. Avoid
    // triggering a redundant `scrollToCell` (and extra scroll event) when syncing selection ranges.
    if (this.sharedGrid) this.syncSharedGridSelectionFromState({ scrollIntoView: false });
    else if (didScroll) this.ensureViewportMappingCurrent();
    if (sheetChanged) {
      const presence = this.collabSession?.presence;
      if (presence) {
        presence.setCursor(this.selection.active);
        presence.setSelections(this.selection.ranges);
      }
    }
    this.renderSelection();
    this.updateStatus();
    if (sheetChanged) {
      this.hideFormulaRangePreviewTooltip();
      this.syncFormulaBarHoverRangeOverlays();
    }
    if (sheetChanged) {
      this.refresh();
    } else if (didScroll) {
      this.refresh("scroll");
    }
    if (focus) this.focus();
  }

  /**
   * Programmatically set selection ranges using shared-grid (CanvasGridRenderer) coordinates.
   *
   * This is primarily intended for split-view selection syncing, where the secondary pane uses
   * `@formula/grid` selection ranges that include the header row/col.
   */
  setSharedGridSelectionRanges(
    ranges: GridCellRange[] | null,
    options?: {
      activeIndex?: number;
      activeCell?: { row: number; col: number } | null;
      scrollIntoView?: boolean;
      focus?: boolean;
    },
  ): void {
    const scrollIntoView = options?.scrollIntoView !== false;
    const focus = options?.focus !== false;

    // Shared-grid mode: delegate to the shared grid so selection semantics (multi-range, active cell)
    // match the renderer and we reuse its selection-change notifications.
    if (this.sharedGrid) {
      this.sharedGrid.setSelectionRanges(ranges, {
        activeIndex: options?.activeIndex,
        activeCell: options?.activeCell,
        scrollIntoView,
      });
      if (focus) this.focus();
      return;
    }

    // Legacy grid mode does not support multi-range or explicit active-cell selection. Mirror the
    // active range only.
    if (!ranges || ranges.length === 0) return;
    const activeIndex = options?.activeIndex ?? 0;
    const idx = Math.max(0, Math.min(activeIndex, ranges.length - 1));
    const activeRange = ranges[idx] ?? ranges[0];
    if (!activeRange) return;

    const docRange = this.docRangeFromGridRange(activeRange);
    this.selectRange({ range: docRange }, { scrollIntoView, focus });
  }

  /**
   * Returns the shared-grid (CanvasGridRenderer) provider when the app is running in shared-grid mode.
   *
   * This is used by the split-view secondary pane so it can reuse the primary grid's provider/cache
   * when available.
   */
  getSharedGridProvider(): DocumentCellProvider | null {
    return this.sharedProvider;
  }

  /**
   * Returns the shared-grid image resolver when running in shared-grid mode.
   *
   * This is primarily used by the split-view secondary pane so it can render the
   * same in-cell images as the primary grid.
   */
  getSharedGridImageResolver(): CanvasGridImageResolver | null {
    if (this.gridMode !== "shared") return null;
    return this.sharedGridImageResolver;
  }

  /**
   * Returns drawing-layer objects (pictures, charts, shapes) for a given sheet.
   *
   * This is used by the split-view secondary pane so drawings render in both panes.
   */
  getDrawingObjects(sheetId: string = this.sheetId): DrawingObject[] {
    const id = String(sheetId ?? this.sheetId);
    if (!id) return [];
    const baseObjects = this.listDrawingObjectsForSheet(id);
    if (!this.useCanvasCharts) return baseObjects;

    const maxZ = baseObjects.reduce((acc, obj) => Math.max(acc, obj.zOrder), -1);
    const charts = this.listCanvasChartDrawingObjectsForSheet(id, maxZ + 1);
    return charts.length > 0 ? [...baseObjects, ...charts] : baseObjects;
  }

  /**
   * Replace drawing objects for the active sheet.
   *
   * Intended for unit tests and ephemeral UI-only drawing edits. This updates the in-memory
   * draw-object cache and triggers a re-render; it does not persist the drawings back to the
   * workbook model.
   */
  setDrawingObjects(objects: DrawingObject[] | null | undefined): void {
    if (this.disposed) return;
    const next = Array.isArray(objects) ? objects : [];
    this.setDrawingObjectsForSheet(next);
    this.renderDrawings(this.sharedGrid ? this.sharedGrid.renderer.scroll.getViewportState() : undefined);
  }

  /**
   * Returns the ImageStore used by the drawings overlay (picture bitmaps).
   *
   * This is used by the split-view secondary pane so it can render the same pictures
   * as the primary grid.
   */
  getDrawingImages(): ImageStore {
    return this.drawingImages;
  }

  /**
   * Best-effort cleanup for the persistent drawings image store (IndexedDB).
   *
   * This scans the current workbook drawings for referenced `imageId`s and deletes any
   * unreferenced records from IndexedDB. This helps keep local persistence bounded after
   * users delete/reinsert pictures.
   */
  async garbageCollectDrawingImages(): Promise<void> {
    const gc = (this.drawingImages as any)?.garbageCollectAsync as ((keep: Iterable<string>) => Promise<void>) | undefined;
    if (typeof gc !== "function") return;

    const keep = new Set<string>();

    const docAny = this.document as any;
    const sheetIds: string[] = (() => {
      if (typeof docAny.getSheetIds === "function") {
        try {
          const ids = docAny.getSheetIds();
          return Array.isArray(ids) ? ids.map((id: any) => String(id ?? "")).filter(Boolean) : [];
        } catch {
          return [];
        }
      }
      return [];
    })();

    const scanDrawings = (drawings: unknown) => {
      if (!Array.isArray(drawings)) return;
      for (const raw of drawings) {
        const kind = (raw as any)?.kind;
        const type = typeof kind?.type === "string" ? kind.type : "";
        if (type !== "image") continue;
        const id = typeof kind?.imageId === "string" ? kind.imageId : typeof kind?.image_id === "string" ? kind.image_id : "";
        if (id) keep.add(id);
      }
    };

    for (const sheetId of sheetIds.length > 0 ? sheetIds : [this.sheetId]) {
      try {
        scanDrawings(docAny.getSheetDrawings?.(sheetId));
      } catch {
        // ignore
      }
    }

    // Include locally-cached drawings (e.g. drag/resize interactions) that may not yet be
    // reflected in the DocumentController snapshot.
    for (const obj of this.sheetDrawings) {
      if (obj.kind.type === "image") keep.add(obj.kind.imageId);
    }
    const cachedObjects = this.drawingObjectsCache;
    if (cachedObjects && cachedObjects.sheetId === this.sheetId) {
      for (const obj of cachedObjects.objects) {
        if (obj.kind.type === "image") keep.add(obj.kind.imageId);
      }
    }

    try {
      await gc.call(this.drawingImages, keep);
    } catch {
      // Best-effort: ignore persistence failures.
    }
  }

  private createDrawingChartRendererStore(): ChartRendererStore {
    return {
      getChartModel: (chartId) => {
        // In canvas-charts mode, ChartStore charts are rendered as drawing objects using
        // their ChartStore ids. Imported workbook charts use their own ids (e.g.
        // `${sheetId}:${drawingObjectId}`) and continue to read from `formulaChartModelStore`.
        if (this.getChartRecordById(chartId)) return this.chartCanvasStoreAdapter.getChartModel(chartId);
        return this.formulaChartModelStore.getChartModel(chartId);
      },
      getChartData: (chartId) => {
        if (this.getChartRecordById(chartId)) return undefined;
        return this.formulaChartModelStore.getChartData(chartId);
      },
      getChartTheme: (chartId) => {
        if (this.getChartRecordById(chartId)) return this.chartCanvasStoreAdapter.getChartTheme(chartId);
        return this.formulaChartModelStore.getChartTheme(chartId);
      },
      getChartRevision: (chartId) => {
        // Return NaN for chart ids we don't own so ChartRendererAdapter falls back to
        // model/theme identity checks rather than freezing cached surfaces.
        if (this.getChartRecordById(chartId)) return this.chartCanvasStoreAdapter.getChartRevision(chartId);
        return Number.NaN;
      },
    };
  }

  /**
   * Chart renderer used by drawings overlays (primary + split-view secondary pane).
   *
   * In `?canvasCharts=1` mode this also supports ChartStore charts rendered as drawing objects.
   */
  getDrawingChartRenderer(): ChartRendererAdapter {
    return new ChartRendererAdapter(this.createDrawingChartRendererStore());
  }

  /**
   * Best-effort selected drawing id (used to render selection handles in the drawings overlay).
   *
   * Prefers the explicit drawing selection state, but falls back to chart selection
   * in canvas-charts mode (so chart selection continues to show handles in split view).
   */
  getSelectedDrawingId(): number | null {
    if (this.selectedDrawingId != null) return this.selectedDrawingId;
    if (!this.useCanvasCharts) return null;
    if (!this.selectedChartId) return null;
    return chartStoreIdToDrawingId(this.selectedChartId);
  }

  getGridLimits(): GridLimits {
    return { ...this.limits };
  }

  /**
   * Callbacks that allow a shared-grid instance (e.g. split-view secondary pane) to drive the
   * formula bar range-selection UX.
   */
  getSharedGridRangeSelectionCallbacks(): Pick<
    DesktopSharedGridCallbacks,
    "onRangeSelectionStart" | "onRangeSelectionChange" | "onRangeSelectionEnd"
  > {
    return {
      onRangeSelectionStart: (range) => this.onSharedRangeSelectionStart(range),
      onRangeSelectionChange: (range) => this.onSharedRangeSelectionChange(range),
      onRangeSelectionEnd: () => this.onSharedRangeSelectionEnd(),
    };
  }

  getSelectionRanges(): Range[] {
    return this.selection.ranges;
  }

  /**
   * Hide or unhide rows (0-based indices).
   */
  setRowsHidden(rows: number[] | null | undefined, hidden: boolean): void {
    if (this.isReadOnly()) return;
    if (this.isEditing()) return;
    if (!Array.isArray(rows) || rows.length === 0) return;

    const outline = this.getOutlineForSheet(this.sheetId);
    let changed = false;
    for (const raw of rows) {
      if (!Number.isFinite(raw)) continue;
      const row = Math.trunc(raw);
      if (row < 0 || row >= this.limits.maxRows) continue;
      const entry = outline.rows.entryMut(row + 1);
      if (entry.hidden.user !== hidden) {
        entry.hidden.user = hidden;
        changed = true;
      }
    }

    if (!changed) return;
    this.onOutlineUpdated();
  }

  /**
   * Hide or unhide columns (0-based indices).
   */
  setColsHidden(cols: number[] | null | undefined, hidden: boolean): void {
    if (this.isReadOnly()) return;
    if (this.isEditing()) return;
    if (!Array.isArray(cols) || cols.length === 0) return;

    const outline = this.getOutlineForSheet(this.sheetId);
    let changed = false;
    for (const raw of cols) {
      if (!Number.isFinite(raw)) continue;
      const col = Math.trunc(raw);
      if (col < 0 || col >= this.limits.maxCols) continue;
      const entry = outline.cols.entryMut(col + 1);
      if (entry.hidden.user !== hidden) {
        entry.hidden.user = hidden;
        changed = true;
      }
    }

    if (!changed) return;
    this.onOutlineUpdated();
  }

  /**
   * Hide or unhide rows (0-based indices) due to an AutoFilter.
   *
   * This sets `outline.rows.*.hidden.filter` (not `.hidden.user`) so that clearing
   * filters does not clobber user-hidden rows.
   *
   * Note: This is only supported in the legacy renderer. Shared-grid mode intentionally does not
   * currently implement outline-based hidden rows/cols.
   */
  setRowsFilteredHidden(rows: number[] | null | undefined, hidden: boolean): void {
    if (this.gridMode !== "legacy") {
      showToast("Filter is not supported in shared grid mode yet.", "info");
      // Preserve keyboard workflows even when the action is unsupported.
      this.focus();
      return;
    }
    if (!Array.isArray(rows) || rows.length === 0) return;

    const outline = this.getOutlineForSheet(this.sheetId);
    let changed = false;
    for (const raw of rows) {
      if (!Number.isFinite(raw)) continue;
      const row = Math.trunc(raw);
      if (row < 0 || row >= this.limits.maxRows) continue;
      const entry = outline.rows.entryMut(row + 1);
      if (entry.hidden.filter !== hidden) {
        entry.hidden.filter = hidden;
        changed = true;
      }
    }

    if (!changed) return;
    this.onOutlineUpdated();
  }

  /**
   * Clear `.hidden.filter` flags for any outline rows in the given (inclusive) 0-based range.
   *
   * This is a best-effort operation used by the ribbon AutoFilter MVP; it does not delete outline
   * entries (so grouping metadata is preserved).
   */
  clearFilteredHiddenRowsInRange(startRow: number, endRow: number): void {
    if (this.gridMode !== "legacy") {
      showToast("Filter is not supported in shared grid mode yet.", "info");
      this.focus();
      return;
    }

    const start = Math.max(0, Math.min(Math.trunc(startRow), Math.trunc(endRow)));
    const end = Math.max(Math.trunc(startRow), Math.trunc(endRow));

    const outline = this.getOutlineForSheet(this.sheetId);
    let changed = false;
    for (const [index1, entry] of outline.rows.entries.entries()) {
      const row = index1 - 1;
      if (row < start || row > end) continue;
      if (entry?.hidden?.filter) {
        entry.hidden.filter = false;
        changed = true;
      }
    }

    if (!changed) return;
    this.onOutlineUpdated();
  }

  /**
   * Clear `.hidden.filter` flags for outline rows across *all* sheets in this app instance.
   *
   * Used by the ribbon AutoFilter MVP to avoid leaking view-local filter hidden state across
   * `restoreDocumentState()` calls (workbook open / version restore).
   */
  clearAllFilteredHiddenRows(): void {
    if (this.gridMode !== "legacy") {
      showToast("Filter is not supported in shared grid mode yet.", "info");
      this.focus();
      return;
    }

    let changed = false;
    for (const outline of this.outlinesBySheet.values()) {
      for (const entry of outline.rows.entries.values()) {
        if (!entry.hidden.filter) continue;
        entry.hidden.filter = false;
        changed = true;
      }
    }

    if (!changed) return;
    this.onOutlineUpdated();
  }

  hideRows(rows: number[] | null | undefined): void {
    this.setRowsHidden(rows, true);
  }

  unhideRows(rows: number[] | null | undefined): void {
    this.setRowsHidden(rows, false);
  }

  hideCols(cols: number[] | null | undefined): void {
    this.setColsHidden(cols, true);
  }

  unhideCols(cols: number[] | null | undefined): void {
    this.setColsHidden(cols, false);
  }

  /**
   * Compute Excel-like status bar stats (Sum / Average / Count) for the current selection.
   *
   * Performance note:
   * - For small selections, we scan the selection coordinates directly.
   * - For large selections (e.g. select-all), we iterate only the *stored* cells in the
   *   DocumentController's sparse cell map (not every coordinate in the rectangular ranges).
   */
  getSelectionSummary(): SpreadsheetSelectionSummary {
    const sheetId = this.sheetId;
    const sheetContentVersion = this.document.getSheetContentVersion(sheetId);
    const workbookContentVersion = this.document.contentVersion;
    // `getCellComputedValue` only consults the WASM engine cache when a single sheet is present.
    // When multiple sheets exist we evaluate formulas in-process.
    const sheetCount = (this.document as any)?.model?.sheets?.size;
    const actualSheetCount = typeof sheetCount === "number" ? sheetCount : this.document.getSheetIds().length;
    const isMultiSheet = actualSheetCount > 1;
    const useEngineCache = actualSheetCount <= 1;

    const cached = this.selectionSummaryCache;
    if (cached && cached.sheetId === sheetId && cached.rangesKey.length === this.selection.ranges.length * 4) {
      // If the cached selection includes formulas in a multi-sheet workbook, those formulas can
      // reference *other* sheets. In that case we must key the cache on the workbook-level content
      // version. Otherwise, we can key it on the active sheet's content version.
      const versionOk =
        cached.selectionHasFormula && isMultiSheet
          ? cached.workbookContentVersion === workbookContentVersion
          : cached.sheetContentVersion === sheetContentVersion;
      if (versionOk) {
        // Only key on engine computed-value churn when the engine cache is used *and* the selection
        // includes formulas (value-only selections do not depend on computed values).
        const computedValuesVersionKey = useEngineCache && cached.selectionHasFormula ? this.computedValuesVersion : 0;
        if (cached.computedValuesVersion === computedValuesVersionKey) {
          let sameRanges = true;
          for (let rangeIdx = 0; rangeIdx < this.selection.ranges.length; rangeIdx += 1) {
            const r = this.selection.ranges[rangeIdx]!;
            const startRow = Math.min(r.startRow, r.endRow);
            const endRow = Math.max(r.startRow, r.endRow);
            const startCol = Math.min(r.startCol, r.endCol);
            const endCol = Math.max(r.startCol, r.endCol);
            const keyIdx = rangeIdx * 4;
            if (
              cached.rangesKey[keyIdx] !== startRow ||
              cached.rangesKey[keyIdx + 1] !== endRow ||
              cached.rangesKey[keyIdx + 2] !== startCol ||
              cached.rangesKey[keyIdx + 3] !== endCol
            ) {
              sameRanges = false;
              break;
            }
          }
          if (sameRanges) return cached.summary;
        }
      }
    }

    const SELECTION_AREA_SCAN_THRESHOLD = 10_000;

    // Encode the selection ranges as a compact numeric key:
    // [startRow, endRow, startCol, endCol, ...] (normalized to start<=end).
    //
    // We keep this separate from `ranges` so we can store it in `selectionSummaryCache`
    // without retaining references to the mutable selection objects.
    const rangesKey: number[] = [];
    const ranges: Array<{ startRow: number; endRow: number; startCol: number; endCol: number }> = [];
    let selectionArea = 0;
    for (const r of this.selection.ranges) {
      const startRow = Math.min(r.startRow, r.endRow);
      const endRow = Math.max(r.startRow, r.endRow);
      const startCol = Math.min(r.startCol, r.endCol);
      const endCol = Math.max(r.startCol, r.endCol);
      rangesKey.push(startRow, endRow, startCol, endCol);
      ranges.push({ startRow, endRow, startCol, endCol });

      if (selectionArea <= SELECTION_AREA_SCAN_THRESHOLD) {
        const rows = Math.max(0, endRow - startRow + 1);
        const cols = Math.max(0, endCol - startCol + 1);
        selectionArea += rows * cols;
        if (selectionArea > SELECTION_AREA_SCAN_THRESHOLD) {
          // We only need to distinguish "small" vs "large"; clamp to avoid pointless arithmetic
          // once we've crossed the threshold.
          selectionArea = SELECTION_AREA_SCAN_THRESHOLD + 1;
        }
      }
    }

    // Fast path: if the sheet has no stored content (value/formula), the summary is always empty
    // regardless of selection area. Avoid scanning potentially thousands of blank coordinates on
    // new/empty sheets.
    const sheetModel: any = (this.document as any)?.model?.sheets?.get?.(sheetId) ?? null;
    if (!sheetModel) {
      const summary: SpreadsheetSelectionSummary = {
        sum: null,
        average: null,
        count: 0,
        numericCount: 0,
        countNonEmpty: 0,
      };
      this.selectionSummaryCache = {
        sheetId,
        sheetContentVersion,
        workbookContentVersion,
        computedValuesVersion: 0,
        selectionHasFormula: false,
        rangesKey,
        summary,
      };
      return summary;
    }
    const contentCellCount = typeof sheetModel.contentCellCount === "number" ? sheetModel.contentCellCount : null;
    if (contentCellCount === 0) {
      const summary: SpreadsheetSelectionSummary = {
        sum: null,
        average: null,
        count: 0,
        numericCount: 0,
        countNonEmpty: 0,
      };
      this.selectionSummaryCache = {
        sheetId,
        sheetContentVersion,
        workbookContentVersion,
        computedValuesVersion: 0,
        selectionHasFormula: false,
        rangesKey,
        summary,
      };
      return summary;
    }

    // Heuristic: for sparse sheets, iterating only stored cells can be faster even when the
    // selection area is below the scan threshold (it avoids `getCell` calls for blank coords).
    const storedCellCount = typeof sheetModel?.cells?.size === "number" ? sheetModel.cells.size : null;
    const useSparseIteration =
      selectionArea > SELECTION_AREA_SCAN_THRESHOLD ||
      (selectionArea <= SELECTION_AREA_SCAN_THRESHOLD &&
        storedCellCount != null &&
        // `storedCellCount` includes format-only cells, so this errs on the side of falling back
        // to coordinate scanning.
        storedCellCount <= selectionArea);
    let countNonEmpty = 0;
    let numericCount = 0;
    let numericSum = 0;
    let selectionHasFormula = false;

    const inSelection =
      ranges.length === 1
        ? (() => {
            const r0 = ranges[0];
            if (!r0) return () => false;
            return (row: number, col: number): boolean =>
              row >= r0.startRow && row <= r0.endRow && col >= r0.startCol && col <= r0.endCol;
          })()
        : (row: number, col: number): boolean => {
            for (const r of ranges) {
              if (row < r.startRow || row > r.endRow) continue;
              if (col < r.startCol || col > r.endCol) continue;
              return true;
            }
            return false;
          };

    // Reuse a single coord object while scanning selection cells to avoid allocating
    // `{row,col}` objects for every visited coordinate.
    const coordScratch = { row: 0, col: 0 };
    // When the engine cache is disabled (multi-sheet) *or* when computed values are missing,
    // we fall back to the in-process evaluator. Sharing a memo/stack across formula cells lets
    // dependency formulas be reused across the selection summary calculation (rather than
    // rebuilding the memo for every root cell).
    let formulaMemo: Map<string, Map<number, SpreadsheetValue>> | null = null;
    let formulaStack: Map<string, Set<number>> | null = null;
    const evalOptions = { useEngineCache };
    const engineSheetCache = useEngineCache ? this.getComputedValuesByCoordForSheet(sheetId) : null;
    const getFormulaValue = (): SpreadsheetValue => {
      if (!formulaMemo) {
        formulaMemo = new Map();
        formulaStack = new Map();
      }
      return this.computeCellValue(sheetId, coordScratch, formulaMemo, formulaStack!, evalOptions);
    };

    if (!useSparseIteration) {
      const visited = ranges.length > 1 ? new Set<number>() : null;
      for (const r of ranges) {
        for (let row = r.startRow; row <= r.endRow; row += 1) {
          for (let col = r.startCol; col <= r.endCol; col += 1) {
            if (visited) {
              const key = row * COMPUTED_COORD_COL_STRIDE + col;
              if (visited.has(key)) continue;
              visited.add(key);
            }

            coordScratch.row = row;
            coordScratch.col = col;
            // Avoid `DocumentController.getCell` clones/allocations while scanning coordinates:
            // we only need read-only access to the stored cell state.
            const cell = sheetModel.cells.get(`${row},${col}`);
            if (!cell) continue;
            // Ignore format-only cells (styleId-only).
            const hasContent = cell.value != null || cell.formula != null;
            if (!hasContent) continue;

            countNonEmpty += 1;

            // Sum/average operate on numeric values only (computed values for formulas).
            if (cell.formula != null) {
              selectionHasFormula = true;
              let computed: SpreadsheetValue | undefined;
              if (engineSheetCache && coordScratch.col >= 0 && coordScratch.col < COMPUTED_COORD_COL_STRIDE && coordScratch.row >= 0) {
                const key = coordScratch.row * COMPUTED_COORD_COL_STRIDE + coordScratch.col;
                computed = engineSheetCache.get(key);
              }
              if (computed === undefined) {
                computed = getFormulaValue();
              }
              if (typeof computed === "number" && Number.isFinite(computed)) {
                numericCount += 1;
                numericSum += computed;
              }
            } else if (typeof cell.value === "number" && Number.isFinite(cell.value)) {
              numericCount += 1;
              numericSum += cell.value;
            }
          }
        }
      }
    } else {
      // Iterate only stored cells (value/formula/format-only), then filter by selection.
      this.document.forEachCellInSheet(this.sheetId, ({ row, col, cell }: any) => {
        if (!inSelection(row, col)) return;

        // Ignore format-only cells (styleId-only).
        const hasContent = cell.value != null || cell.formula != null;
        if (!hasContent) return;

        countNonEmpty += 1;

        // Sum/average operate on numeric values only (computed values for formulas).
        if (cell.formula != null) {
          selectionHasFormula = true;
          coordScratch.row = row;
          coordScratch.col = col;
          let computed: SpreadsheetValue | undefined;
          if (engineSheetCache && coordScratch.col >= 0 && coordScratch.col < COMPUTED_COORD_COL_STRIDE && coordScratch.row >= 0) {
            const key = coordScratch.row * COMPUTED_COORD_COL_STRIDE + coordScratch.col;
            computed = engineSheetCache.get(key);
          }
          if (computed === undefined) {
            computed = getFormulaValue();
          }
          if (typeof computed === "number" && Number.isFinite(computed)) {
            numericCount += 1;
            numericSum += computed;
          }
          return;
        }

        if (typeof cell.value === "number" && Number.isFinite(cell.value)) {
          numericCount += 1;
          numericSum += cell.value;
        }
      });
    }

    const sum = numericCount > 0 ? numericSum : null;
    const average = numericCount > 0 ? numericSum / numericCount : null;

    const summary: SpreadsheetSelectionSummary = {
      sum,
      average,
      count: countNonEmpty,
      numericCount,
      countNonEmpty,
    };

    const computedValuesVersionKey = useEngineCache && selectionHasFormula ? this.computedValuesVersion : 0;
    this.selectionSummaryCache = {
      sheetId,
      sheetContentVersion,
      workbookContentVersion,
      computedValuesVersion: computedValuesVersionKey,
      selectionHasFormula,
      rangesKey,
      summary,
    };

    return summary;
  }

  getActiveCell(): CellCoord {
    return { ...this.selection.active };
  }

  /**
   * Return the active cell's bounding box in viewport coordinates.
   *
   * This is useful for anchoring floating UI (e.g. context menus) to the active cell
   * without changing selection.
   */
  getActiveCellRect(): { x: number; y: number; width: number; height: number } | null {
    if (!this.sharedGrid) {
      this.ensureViewportMappingCurrent();
    }

    const rect = this.getCellRect(this.selection.active);
    if (!rect) return null;

    const rootRect = (this.sharedGrid ? this.selectionCanvas : this.root).getBoundingClientRect();
    return {
      x: rootRect.left + rect.x,
      y: rootRect.top + rect.y,
      width: rect.width,
      height: rect.height,
    };
  }

  /**
   * Hit-test the grid and return the document cell (0-based row/col) under the
   * provided client (viewport) coordinates.
   *
   * This is intentionally limited to sheet cells for now (it returns null when
   * the point lands on row/column headers).
   */
  pickCellAtClientPoint(clientX: number, clientY: number): CellCoord | null {
    const rootRect = this.root.getBoundingClientRect();
    const x = clientX - rootRect.left;
    const y = clientY - rootRect.top;
    if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
    if (x < 0 || y < 0 || x > rootRect.width || y > rootRect.height) return null;

    // Only treat hits in the cell grid (exclude row/col headers).
    if (!this.sharedGrid) {
      if (x < this.rowHeaderWidth || y < this.colHeaderHeight) return null;
      return this.cellFromPoint(x, y);
    }

    // Shared grid uses its own internal coordinate space anchored on the
    // selection canvas.
    const canvasRect = this.selectionCanvas.getBoundingClientRect();
    const vx = clientX - canvasRect.left;
    const vy = clientY - canvasRect.top;
    if (!Number.isFinite(vx) || !Number.isFinite(vy)) return null;
    if (vx < 0 || vy < 0 || vx > canvasRect.width || vy > canvasRect.height) return null;

    const picked = this.sharedGrid.renderer.pickCellAt(vx, vy);
    if (!picked) return null;

    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    if (picked.row < headerRows || picked.col < headerCols) return null;

    return { row: picked.row - headerRows, col: picked.col - headerCols };
  }

  /**
   * Hit-test the grid and return the area under the provided client coordinates.
   *
   * Unlike `pickCellAtClientPoint`, this also reports hits on the row/column headers.
   *
   * Returns `area:"cell"` with null row/col when the point cannot be resolved.
   */
  hitTestGridAreaAtClientPoint(
    clientX: number,
    clientY: number,
  ): { area: "cell" | "rowHeader" | "colHeader" | "corner"; row: number | null; col: number | null } {
    const rootRect = this.root.getBoundingClientRect();
    const x = clientX - rootRect.left;
    const y = clientY - rootRect.top;
    if (!Number.isFinite(x) || !Number.isFinite(y)) return { area: "cell", row: null, col: null };

    if (this.sharedGrid) {
      // Shared grid uses its own internal coordinate space anchored on the selection canvas.
      const canvasRect = this.selectionCanvas.getBoundingClientRect();
      const vx = clientX - canvasRect.left;
      const vy = clientY - canvasRect.top;
      if (!Number.isFinite(vx) || !Number.isFinite(vy)) return { area: "cell", row: null, col: null };
      if (vx < 0 || vy < 0 || vx > canvasRect.width || vy > canvasRect.height) {
        return { area: "cell", row: null, col: null };
      }

      const picked = this.sharedGrid.renderer.pickCellAt(vx, vy);
      if (!picked) return { area: "cell", row: null, col: null };

      const headerRows = this.sharedHeaderRows();
      const headerCols = this.sharedHeaderCols();

      if (picked.row < headerRows && picked.col < headerCols) return { area: "corner", row: null, col: null };
      if (picked.col < headerCols) return { area: "rowHeader", row: picked.row - headerRows, col: null };
      if (picked.row < headerRows) return { area: "colHeader", row: null, col: picked.col - headerCols };
      return { area: "cell", row: picked.row - headerRows, col: picked.col - headerCols };
    }

    const inRowHeader = x < this.rowHeaderWidth;
    const inColHeader = y < this.colHeaderHeight;
    if (inRowHeader && inColHeader) return { area: "corner", row: null, col: null };

    if (inRowHeader || inColHeader) {
      const cell = this.cellFromPoint(x, y);
      if (inRowHeader) return { area: "rowHeader", row: cell.row, col: null };
      return { area: "colHeader", row: null, col: cell.col };
    }

    const picked = this.pickCellAtClientPoint(clientX, clientY);
    if (picked) return { area: "cell", row: picked.row, col: picked.col };

    return { area: "cell", row: null, col: null };
  }

  /**
   * Drawing overlay state helpers (selection, commands).
   *
   * These are public primarily so external integration points (context menus, keyboard shortcuts,
   * e2e harnesses) can interact with drawing objects without needing direct state access.
   */
  pickDrawingAtClientPoint(clientX: number, clientY: number): number | null {
    const objects = this.listDrawingObjectsForSheet();
    if (objects.length === 0) return null;

    // Shared-grid mode uses its own internal coordinate space anchored on the
    // selection canvas.
    const canvasRect = this.selectionCanvas.getBoundingClientRect();
    const x = clientX - canvasRect.left;
    const y = clientY - canvasRect.top;
    if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
    if (x < 0 || y < 0 || x > canvasRect.width || y > canvasRect.height) return null;

    const viewport = this.getDrawingInteractionViewport();
    const index = this.getDrawingHitTestIndex(objects);
    const hit = hitTestDrawings(index, viewport, x, y);
    return hit?.object.id ?? null;
  }

  selectDrawing(id: number | null): void {
    const prev = this.selectedDrawingId;
    this.selectedDrawingId = id;
    // Keep all drawing interaction controllers in sync so keyboard-driven selection changes
    // (e.g. Escape to deselect) don't leave pointer interactions thinking an object is still
    // selected (which could enable invisible resize/rotate handles).
    this.drawingInteractionController?.setSelectedId(id);
    this.drawingOverlay.setSelectedId(id);
    if (prev !== id) {
      this.dispatchDrawingSelectionChanged();
    }
    this.renderDrawings();
  }

  deleteSelectedDrawing(): void {
    const drawingId = this.selectedDrawingId;
    if (drawingId == null) return;

    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;

    const sheetId = this.sheetId;
    const selected = this.listDrawingObjectsForSheet(sheetId).find((obj) => obj.id === drawingId) ?? null;
    const imageId = selected?.kind.type === "image" ? selected.kind.imageId : null;

    const docAny: any = this.document as any;
    const deleteDrawing =
      typeof docAny.deleteDrawing === "function"
        ? (docAny.deleteDrawing as (sheetId: string, drawingId: string | number, options?: unknown) => void)
        : null;
    const getSheetDrawings =
      typeof docAny.getSheetDrawings === "function" ? (docAny.getSheetDrawings as (sheetId: string) => unknown) : null;
    const deleteImage =
      typeof docAny.deleteImage === "function" ? (docAny.deleteImage as (imageId: string, options?: unknown) => void) : null;

    // `DrawingObject.id` is numeric in the UI layer, but DocumentController drawings can be stored with
    // string ids (and we map them through a stable numeric hash). To delete the correct raw entry,
    // scan the raw drawings list and compare via the adapter layer.
    const rawIdsToDelete = new Set<string | number>();
    if (getSheetDrawings) {
      let raw: unknown = null;
      try {
        raw = getSheetDrawings.call(docAny, sheetId);
      } catch {
        raw = null;
      }
      if (Array.isArray(raw)) {
        for (const entry of raw) {
          if (!entry || typeof entry !== "object") continue;
          let uiId: number | null = null;
          try {
            uiId = convertDocumentSheetDrawingsToUiDrawingObjects([entry], { sheetId })[0]?.id ?? null;
          } catch {
            uiId = null;
          }
          if (uiId !== drawingId) continue;
          const rawId = (entry as any).id;
          if (typeof rawId === "string") {
            const trimmed = rawId.trim();
            if (trimmed) rawIdsToDelete.add(trimmed);
          } else if (typeof rawId === "number" && Number.isFinite(rawId)) {
            rawIdsToDelete.add(rawId);
          }
        }
      }
    }
    if (rawIdsToDelete.size === 0) rawIdsToDelete.add(drawingId);

    let batchStarted = false;
    try {
      this.document.beginBatch({ label: "Delete Drawing" });
      batchStarted = true;
    } catch {
      batchStarted = false;
    }

    if (deleteDrawing) {
      for (const rawId of rawIdsToDelete) {
        try {
          deleteDrawing.call(docAny, sheetId, rawId, { label: "Delete Drawing" });
        } catch {
          // ignore
        }
      }
    }

    if (imageId && deleteImage && !this.isImageReferencedByAnyDrawing(imageId)) {
      try {
        deleteImage.call(docAny, imageId, { label: "Delete Drawing" });
      } catch {
        // ignore
      }
      this.drawingOverlay.invalidateImage(imageId);
    }

    if (batchStarted) {
      try {
        this.document.endBatch();
      } catch {
        // ignore
      }
    }

    this.selectedDrawingId = null;
    this.dispatchDrawingSelectionChanged();
    this.drawingOverlay.setSelectedId(null);
    this.drawingInteractionController?.setSelectedId(null);
    this.drawingObjectsCache = null;
    this.drawingHitTestIndex = null;
    this.drawingHitTestIndexObjects = null;
    this.renderDrawings(this.sharedGrid ? this.sharedGrid.renderer.scroll.getViewportState() : undefined);
  }

  duplicateSelectedDrawing(): void {
    const selected = this.selectedDrawingId;
    if (selected == null) return;
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;

    const result = duplicateDrawingSelected(this.listDrawingObjectsForSheet(), selected);
    if (!result) return;

    this.document.beginBatch({ label: "Duplicate Drawing" });
    try {
      this.document.setSheetDrawings(this.sheetId, result.objects, { source: "drawings" });
      this.document.endBatch();
    } catch (err) {
      this.document.cancelBatch();
      throw err;
    }

    // Clear the caches so hit testing + overlay re-render see the committed state.
    this.drawingObjectsCache = null;
    this.drawingHitTestIndex = null;
    this.drawingHitTestIndexObjects = null;
    this.selectDrawing(result.duplicatedId);
  }

  bringSelectedDrawingForward(): void {
    this.arrangeSelectedDrawing("forward");
  }

  sendSelectedDrawingBackward(): void {
    this.arrangeSelectedDrawing("backward");
  }

  bringSelectedDrawingToFront(): void {
    this.arrangeSelectedDrawing("front");
  }

  sendSelectedDrawingToBack(): void {
    this.arrangeSelectedDrawing("back");
  }

  fillDown(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.applyFillShortcut("down", "formulas");
  }

  fillUp(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.applyFillShortcut("up", "formulas");
  }

  fillRight(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.applyFillShortcut("right", "formulas");
  }

  fillLeft(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.applyFillShortcut("left", "formulas");
  }

  fillSeries(direction: "down" | "right" | "up" | "left"): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.applyFillShortcut(direction, "series");
  }

  async insertRows(row: number, count: number): Promise<void> {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;

    const sheetId = this.sheetId;
    const row0 = Math.trunc(row);
    const count0 = Math.trunc(count);
    if (!Number.isInteger(row0) || row0 < 0) return;
    if (!Number.isInteger(count0) || count0 <= 0) return;

    // Best-effort fallback when the engine is unavailable (e.g. restricted environments).
    if (!this.wasmEngine) {
      this.document.insertRows(sheetId, row0, count0, { label: "Insert Rows", source: "ribbon" });
      this.refresh();
      this.focus();
      return;
    }

    const op: EditOp = { type: "InsertRows", sheet: sheetId, row: row0, count: count0 };
    await this.applyStructuralEdit(op, (result) => {
      this.document.insertRows(sheetId, row0, count0, {
        label: "Insert Rows",
        source: "ribbon",
        formulaRewrites: result.formulaRewrites,
      });
    }, { label: "Insert Rows" });

    this.refresh();
    this.focus();
  }

  async deleteRows(row: number, count: number): Promise<void> {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;

    const sheetId = this.sheetId;
    const row0 = Math.trunc(row);
    const count0 = Math.trunc(count);
    if (!Number.isInteger(row0) || row0 < 0) return;
    if (!Number.isInteger(count0) || count0 <= 0) return;

    if (!this.wasmEngine) {
      this.document.deleteRows(sheetId, row0, count0, { label: "Delete Rows", source: "ribbon" });
      this.refresh();
      this.focus();
      return;
    }

    const op: EditOp = { type: "DeleteRows", sheet: sheetId, row: row0, count: count0 };
    await this.applyStructuralEdit(op, (result) => {
      this.document.deleteRows(sheetId, row0, count0, {
        label: "Delete Rows",
        source: "ribbon",
        formulaRewrites: result.formulaRewrites,
      });
    }, { label: "Delete Rows" });

    this.refresh();
    this.focus();
  }

  async insertCols(col: number, count: number): Promise<void> {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;

    const sheetId = this.sheetId;
    const col0 = Math.trunc(col);
    const count0 = Math.trunc(count);
    if (!Number.isInteger(col0) || col0 < 0) return;
    if (!Number.isInteger(count0) || count0 <= 0) return;

    if (!this.wasmEngine) {
      this.document.insertCols(sheetId, col0, count0, { label: "Insert Columns", source: "ribbon" });
      this.refresh();
      this.focus();
      return;
    }

    const op: EditOp = { type: "InsertCols", sheet: sheetId, col: col0, count: count0 };
    await this.applyStructuralEdit(op, (result) => {
      this.document.insertCols(sheetId, col0, count0, {
        label: "Insert Columns",
        source: "ribbon",
        formulaRewrites: result.formulaRewrites,
      });
    }, { label: "Insert Columns" });

    this.refresh();
    this.focus();
  }

  async deleteCols(col: number, count: number): Promise<void> {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;

    const sheetId = this.sheetId;
    const col0 = Math.trunc(col);
    const count0 = Math.trunc(count);
    if (!Number.isInteger(col0) || col0 < 0) return;
    if (!Number.isInteger(count0) || count0 <= 0) return;

    if (!this.wasmEngine) {
      this.document.deleteCols(sheetId, col0, count0, { label: "Delete Columns", source: "ribbon" });
      this.refresh();
      this.focus();
      return;
    }

    const op: EditOp = { type: "DeleteCols", sheet: sheetId, col: col0, count: count0 };
    await this.applyStructuralEdit(op, (result) => {
      this.document.deleteCols(sheetId, col0, count0, {
        label: "Delete Columns",
        source: "ribbon",
        formulaRewrites: result.formulaRewrites,
      });
    }, { label: "Delete Columns" });

    this.refresh();
    this.focus();
  }

  async insertCells(range: Range, direction: "right" | "down"): Promise<void> {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    const sheetId = this.sheetId;
    const a1 = rangeToA1(range);

    const op: EditOp =
      direction === "right"
        ? { type: "InsertCellsShiftRight", sheet: sheetId, range: a1 }
        : { type: "InsertCellsShiftDown", sheet: sheetId, range: a1 };

    await this.applyStructuralEdit(op, (result) => {
      if (direction === "right") {
        this.document.insertCellsShiftRight(sheetId, range, { label: "Insert Cells", formulaRewrites: result.formulaRewrites });
      } else {
        this.document.insertCellsShiftDown(sheetId, range, { label: "Insert Cells", formulaRewrites: result.formulaRewrites });
      }
    }, { label: "Insert Cells" });

    this.refresh();
    this.focus();
  }

  async deleteCells(range: Range, direction: "left" | "up"): Promise<void> {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    const sheetId = this.sheetId;
    const a1 = rangeToA1(range);

    const op: EditOp =
      direction === "left"
        ? { type: "DeleteCellsShiftLeft", sheet: sheetId, range: a1 }
        : { type: "DeleteCellsShiftUp", sheet: sheetId, range: a1 };

    await this.applyStructuralEdit(op, (result) => {
      if (direction === "left") {
        this.document.deleteCellsShiftLeft(sheetId, range, { label: "Delete Cells", formulaRewrites: result.formulaRewrites });
      } else {
        this.document.deleteCellsShiftUp(sheetId, range, { label: "Delete Cells", formulaRewrites: result.formulaRewrites });
      }
    }, { label: "Delete Cells" });

    this.refresh();
    this.focus();
  }

  private async applyStructuralEdit(
    op: EditOp,
    applyToDocument: (result: EditResult) => void,
    options: { label: string },
  ): Promise<void> {
    const engine = this.wasmEngine;
    if (!engine) {
      showToast("This command requires the WASM engine.");
      return;
    }

    this.wasmSyncSuspended = true;
    try {
      // Ensure the engine has processed all prior deltas before we apply a structural op.
      await this.wasmSyncPromise.catch(() => {});

      let result: EditResult | null = null;
      let opError: unknown = null;
      let docError: unknown = null;

      await this.enqueueWasmSync(async (worker) => {
        try {
          result = await worker.applyOperation(op);
        } catch (err) {
          opError = err;
          throw err;
        }
      });

      if (opError || !result) {
        const message = opError instanceof Error ? opError.message : String(opError ?? "unknown error");
        showToast(`Failed to apply ${options.label}: ${message}`, "error");
        return;
      }

      try {
        applyToDocument(result);
      } catch (err) {
        console.error("[formula][desktop] Failed to apply structural edit to document:", err);
        docError = err;
        const message = err instanceof Error ? err.message : String(err);
        showToast(`Failed to apply ${options.label}: ${message}`, "error");
      }

      // Re-hydrate the WASM engine from the DocumentController to avoid incremental delta loops.
      this.clearComputedValuesByCoord();
      let hydrateError: unknown = null;
      await this.enqueueWasmSync(async (worker) => {
        try {
          const changes = await engineHydrateFromDocument(worker, this.document);
          this.applyComputedChanges(changes);
        } catch (err) {
          hydrateError = err;
          throw err;
        }
      });
      if (hydrateError) {
        const message = hydrateError instanceof Error ? hydrateError.message : String(hydrateError);
        showToast(`Failed to sync ${options.label} to engine: ${message}`, "error");
      }

      if (docError) return;
    } finally {
      this.wasmSyncSuspended = false;
    }
  }

  selectCurrentRegion(): void {
    if (this.inlineEditController.isOpen()) return;
    if (this.editor.isOpen()) return;
    if (this.formulaBar?.isEditing() || this.formulaEditCell) return;

    const active = { ...this.selection.active };
    const range = computeCurrentRegionRange(active, this.usedRangeProvider(), this.limits);
    this.selection = buildSelection({ ranges: [range], active, anchor: active, activeRangeIndex: 0 }, this.limits);
    if (this.sharedGrid) this.syncSharedGridSelectionFromState();
    this.renderSelection();
    this.updateStatus();
    this.focus();
  }

  insertDate(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.insertCurrentDateTimeIntoSelection("date");
    this.refresh();
    this.focus();
  }

  insertTime(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.insertCurrentDateTimeIntoSelection("time");
    this.refresh();
    this.focus();
  }

  autoSumAverage(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.autoSumSelection("AVERAGE");
    this.refresh();
    this.focus();
  }

  autoSumCountNumbers(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.autoSumSelection("COUNT");
    this.refresh();
    this.focus();
  }

  autoSumMax(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.autoSumSelection("MAX");
    this.refresh();
    this.focus();
  }

  autoSumMin(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.autoSumSelection("MIN");
    this.refresh();
    this.focus();
  }

  insertImageFromLocalFile(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    if (typeof document === "undefined") return;

    // Desktop/Tauri: prefer native dialog + backend reads (avoids `<input type=file>` sandbox quirks).
    // Web fallback below retains the persistent `<input>` element so unit tests can interact with it.
    const tauriDialogOpenAvailable = getTauriDialogOpenOrNull() != null;
    const tauriInvokeAvailable = typeof (globalThis as any).__TAURI__?.core?.invoke === "function";
    if (tauriDialogOpenAvailable && tauriInvokeAvailable) {
      void (async () => {
        const files = await pickLocalImageFiles({ multiple: false });
        const file = files[0] ?? null;
        if (!file) {
          this.focus();
          return;
        }
        await this.insertImageFromPickedFile(file);
      })().catch((err) => {
        console.warn("Insert image failed", err);
        this.focus();
      });
      return;
    }

    const input = this.ensureInsertImageInput();
    // Allow selecting the same file repeatedly.
    input.value = "";
    input.onchange = () => {
      const file = input.files?.[0] ?? null;
      if (!file) return;
      void this.insertImageFromPickedFile(file).catch((err) => {
        console.warn("Insert image failed", err);
      });
    };

    try {
      input.click();
    } catch {
      // Best-effort; some environments (tests) may not support programmatic clicks.
    }
  }

  autoSum(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.autoSumSelection("SUM");
    this.refresh();
    this.focus();
  }

  openCellEditorAtActiveCell(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.inlineEditController.isOpen()) return;
    if (this.editor.isOpen()) return;
    if (this.formulaBar?.isEditing() || this.formulaEditCell) return;

    const cell = this.selection.active;
    const bounds = this.getCellRect(cell);
    if (!bounds) return;
    const initialValue = this.getCellInputText(cell);
    this.editor.open(cell, bounds, initialValue, { cursor: "end" });
    this.updateEditState();
  }

  private ensureInsertImageInput(): HTMLInputElement {
    const existing = this.insertImageInput;
    if (existing && existing.isConnected) return existing;

    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    input.style.display = "none";
    input.dataset.testid = "insert-image-input";

    // Keep the element mounted so subsequent insertions reuse it (and tests can find it).
    this.root.appendChild(input);
    this.insertImageInput = input;
    return input;
  }

  private async insertImageFromPickedFile(file: File): Promise<void> {
    if (this.isReadOnly()) return;
    const size = typeof (file as any)?.size === "number" ? (file as any).size : null;
    if (size == null || size > MAX_INSERT_IMAGE_BYTES) {
      const mb = Math.round(MAX_INSERT_IMAGE_BYTES / 1024 / 1024);
      try {
        showToast(`Image too large (>${mb}MB). Choose a smaller file.`, "warning");
      } catch {
        // `showToast` requires a #toast-root; unit tests don't always include it.
      }
      return;
    }

    const imageId = (() => {
      const randomUUID = (globalThis as any)?.crypto?.randomUUID;
      if (typeof randomUUID === "function") {
        try {
          return String(randomUUID.call((globalThis as any).crypto));
        } catch {
          // Fall back to monotonic ids below.
        }
      }
      return `image_${this.nextDrawingImageId++}`;
    })();

    const active = this.selection.active;
    const anchor: DrawingAnchor = {
      type: "oneCell",
      from: { cell: { row: active.row, col: active.col }, offset: { xEmu: 0, yEmu: 0 } },
      size: { cx: pxToEmu(200), cy: pxToEmu(150) },
    };

    const existingObjects = this.listDrawingObjectsForSheet();
    const maxZOrder = existingObjects.reduce((max, obj) => Math.max(max, obj.zOrder), -1);

    const docAny = this.document as any;
    const drawingsGetter = typeof docAny.getSheetDrawings === "function" ? docAny.getSheetDrawings : null;
    const canInsertDrawing = typeof docAny.insertDrawing === "function";

    if (canInsertDrawing) {
      this.document.beginBatch({ label: "Insert Image" });
    }

    try {
      const { objects: combinedObjects, image } = await insertImageFromFile(file, {
        imageId,
        anchor,
        objects: existingObjects,
        images: this.drawingImages,
      });

      const inserted = combinedObjects[combinedObjects.length - 1];
      if (!inserted) {
        if (canInsertDrawing) this.document.endBatch();
        return;
      }

      // Prefer placing the new object on top of existing drawings.
      inserted.zOrder = maxZOrder + 1;

      if (canInsertDrawing) {
        try {
          docAny.insertDrawing(this.sheetId, inserted);
        } catch {
          // Best-effort: if inserting into the document fails, fall back to the in-memory cache below.
          try {
            this.document.cancelBatch();
          } catch {
            // ignore
          }
        }
      }

      if (canInsertDrawing) {
        this.document.endBatch();
      }

      // Update the cache immediately so the first re-render includes the inserted object even
      // if the DocumentController does not publish drawing changes synchronously.
      this.drawingObjectsCache = { sheetId: this.sheetId, objects: combinedObjects, source: drawingsGetter };

      // Preload the bitmap so the first overlay render can reuse the decode promise.
      void this.drawingOverlay.preloadImage(image).catch(() => {
        // ignore
      });

      try {
        this.imageBytesBinder?.onLocalImageInserted(image);
      } catch {
        // Best-effort: never fail insertion due to collab image propagation.
      }

      const prevSelected = this.selectedDrawingId;
      this.selectedDrawingId = inserted.id;
      this.drawingOverlay.setSelectedId(inserted.id);
      if (this.gridMode === "shared") {
        this.ensureDrawingInteractionController().setSelectedId(inserted.id);
      }
      if (prevSelected !== inserted.id) {
        this.dispatchDrawingSelectionChanged();
      }
      this.renderDrawings();
      this.renderSelection();
      this.focus();
    } catch (err) {
      if (canInsertDrawing) {
        try {
          this.document.cancelBatch();
        } catch {
          // ignore
        }
      }
      throw err;
    }
  }

  private ensureDrawingInteractionController(): DrawingInteractionController {
    const existing = this.drawingInteractionController;
    if (existing) return existing;

    const callbacks: DrawingInteractionCallbacks = {
      getViewport: () => this.getDrawingInteractionViewport(this.sharedGrid?.renderer.scroll.getViewportState()),
      getObjects: () => this.listDrawingObjectsForSheet(),
      setObjects: (next) => {
        this.setDrawingObjectsForSheet(next);
        this.renderDrawings();
        const selected = this.selectedDrawingId != null ? next.find((obj) => obj.id === this.selectedDrawingId) : undefined;
        if (selected?.kind.type === "chart") {
          // Best-effort: keep chart overlays aligned when moving/resizing chart drawings.
          this.renderCharts(false);
        }
      },
      shouldHandlePointerDown: () => !this.formulaBar?.isFormulaEditing(),
      onPointerDownHit: () => {
        if (this.editor.isOpen()) {
          this.editor.commit("command");
        }
      },
      onSelectionChange: (selectedId) => {
        const prev = this.selectedDrawingId;
        this.selectedDrawingId = selectedId;
        this.drawingOverlay.setSelectedId(selectedId);
        if (prev !== selectedId) {
          this.dispatchDrawingSelectionChanged();
        }
        if (selectedId != null && this.selectedChartId != null) {
          // Drawings and charts share a single selection model; selecting one should clear the other.
          this.setSelectedChartId(null);
        }
        this.renderDrawings();
        this.focus();
      },
      requestFocus: () => this.focus(),
    };
    this.drawingInteractionCallbacks = callbacks;
    const interactionElement = this.gridMode === "shared" ? this.selectionCanvas : this.root;
    const controller = new DrawingInteractionController(interactionElement, this.drawingGeom, callbacks, {
      capture: this.gridMode === "shared",
    });
    this.drawingInteractionController = controller;
    return controller;
  }

  subscribeSelection(listener: (selection: SelectionState) => void): () => void {
    this.selectionListeners.add(listener);
    listener(this.selection);
    return () => this.selectionListeners.delete(listener);
  }

  private sharedHeaderRows(): number {
    return this.sharedGrid ? 1 : 0;
  }

  private sharedHeaderCols(): number {
    return this.sharedGrid ? 1 : 0;
  }

  private docCellFromGridCell(cell: { row: number; col: number }): CellCoord {
    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    return { row: Math.max(0, cell.row - headerRows), col: Math.max(0, cell.col - headerCols) };
  }

  private gridCellFromDocCell(cell: CellCoord): { row: number; col: number } {
    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    return { row: cell.row + headerRows, col: cell.col + headerCols };
  }

  private gridRangeFromDocRange(range: Range): GridCellRange {
    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    return {
      startRow: range.startRow + headerRows,
      endRow: range.endRow + headerRows + 1,
      startCol: range.startCol + headerCols,
      endCol: range.endCol + headerCols + 1
    };
  }

  private docRangeFromGridRange(range: GridCellRange): Range {
    // Grid ranges used for formula reference insertion always come from a DesktopSharedGrid instance,
    // which includes 1 frozen header row/col (row/col labels). This remains true even when the primary
    // pane uses the legacy renderer (eg `/?grid=legacy`): the split-view secondary pane is still a
    // shared-grid instance.
    const headerRows = 1;
    const headerCols = 1;
    return {
      startRow: Math.max(0, range.startRow - headerRows),
      endRow: Math.max(0, range.endRow - headerRows - 1),
      startCol: Math.max(0, range.startCol - headerCols),
      endCol: Math.max(0, range.endCol - headerCols - 1)
    };
  }

  private syncSelectionFromSharedGrid(): void {
    if (!this.sharedGrid) return;
    const selection = this.sharedGrid.renderer.getSelection();
    const ranges = this.sharedGrid.renderer.getSelectionRanges();
    const activeIndex = this.sharedGrid.renderer.getActiveSelectionIndex();

    if (!selection || ranges.length === 0) {
      this.selection = createSelection({ row: 0, col: 0 }, this.limits);
      return;
    }

    const docActive = this.docCellFromGridCell(selection);
    const docRanges = ranges.map((r) => this.docRangeFromGridRange(r));
    const activeRangeIndex = Math.max(0, Math.min(docRanges.length - 1, activeIndex));
    const anchor = { ...docActive };

    this.selection = buildSelection(
      {
        ranges: docRanges,
        active: docActive,
        anchor,
        activeRangeIndex
      },
      this.limits
    );
  }

  private syncSharedGridSelectionFromState(options?: { scrollIntoView?: boolean }): void {
    if (!this.sharedGrid) return;
    const gridRanges = this.selection.ranges.map((r) => this.gridRangeFromDocRange(r));
    const gridActive = this.gridCellFromDocCell(this.selection.active);
    this.sharedGridSelectionSyncInProgress = true;
    try {
      this.sharedGrid.setSelectionRanges(gridRanges, {
        activeIndex: this.selection.activeRangeIndex,
        activeCell: gridActive,
        scrollIntoView: options?.scrollIntoView,
      });
      this.scrollX = this.sharedGrid.getScroll().x;
      this.scrollY = this.sharedGrid.getScroll().y;
    } finally {
      this.sharedGridSelectionSyncInProgress = false;
    }
  }

  private openEditorFromSharedGrid(request: { row: number; col: number; initialKey?: string }): void {
    if (!this.sharedGrid) return;
    if (this.isReadOnly()) {
      const headerRows = this.sharedHeaderRows();
      const headerCols = this.sharedHeaderCols();
      if (request.row >= headerRows && request.col >= headerCols) {
        const docCell = this.docCellFromGridCell({ row: request.row, col: request.col });
        showCollabEditRejectedToast([
          { sheetId: this.sheetId, row: docCell.row, col: docCell.col, rejectionKind: "cell", rejectionReason: "permission" },
        ]);
      }
      return;
    }
    if (this.editor.isOpen()) return;
    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    if (request.row < headerRows || request.col < headerCols) return;
    const docCell = this.docCellFromGridCell({ row: request.row, col: request.col });
    const rect = this.sharedGrid.getCellRect(request.row, request.col);
    if (!rect) return;
    const initialValue = request.initialKey ?? this.getCellInputText(docCell);
    this.editor.open(docCell, rect, initialValue, { cursor: "end" });
    this.updateEditState();
  }

  private onSharedGridAxisSizeChange(change: GridAxisSizeChange): void {
    if (!this.sharedGrid) return;
    this.clearSharedHoverCellCache();
    this.hideCommentTooltip();
    // Keep drawings spatial indices in sync with axis size changes (row/col resize,
    // auto-fit, etc). The drawing geometry is backed by live shared-grid scroll
    // state, so cached sheet-space bounds must be recomputed.
    const drawingOverlay = (this as any).drawingOverlay as DrawingOverlay | undefined;
    drawingOverlay?.invalidateSpatialIndex();

    // Do not allow row/col resize/auto-fit to mutate the sheet while the user is actively editing
    // (cell editor, formula bar, inline edit). This keeps edit state isolated from unrelated
    // document mutations (Excel-like "editing mode" safety).
    //
    // Note: DesktopSharedGrid updates the renderer sizes interactively during the drag, so when we
    // no-op the document mutation we must also restore the renderer to its previous size.
    if (this.isEditing()) {
      const renderer = this.sharedGrid.renderer;
      const EPS = 1e-6;
      if (change.kind === "col") {
        const prev = change.previousSize;
        const isDefault = Math.abs(prev - change.defaultSize) < EPS;
        if (isDefault) renderer.resetColWidth(change.index);
        else renderer.setColWidth(change.index, prev);
      } else {
        const prev = change.previousSize;
        const isDefault = Math.abs(prev - change.defaultSize) < EPS;
        if (isDefault) renderer.resetRowHeight(change.index);
        else renderer.setRowHeight(change.index, prev);
      }

      // Restore focus to the active editing surface so the user can continue typing.
      if (this.editor.isOpen()) {
        try {
          (this.editor.element as any).focus?.({ preventScroll: true });
        } catch {
          this.editor.element.focus();
        }
      } else if (this.formulaBar?.isEditing() || this.formulaEditCell) {
        this.formulaBar?.focus();
      }

      // Even though we no-op the underlying document mutation, the shared-grid renderer may have
      // updated axis sizes interactively during the drag. Ensure drawings/pictures overlays are
      // redrawn to stay aligned with the restored renderer geometry.
      this.scheduleDrawingsRender();
      return;
    }

    // Tag sheet-view mutations originating from the primary shared grid so the document `change`
    // listener can avoid redundantly re-syncing axis sizes back into the same renderer (which is
    // already updated interactively during the drag). Other panes (e.g. split view) will still
    // observe the change and sync their own renderer instances.
    const source = "sharedGridAxis";

    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();
    const baseSize = change.size / change.zoom;
    const baseDefault = change.defaultSize / change.zoom;
    const isDefault = Math.abs(baseSize - baseDefault) < 1e-6;

    if (change.kind === "col") {
      const docCol = change.index - headerCols;
      if (docCol < 0) return;
      const label = change.source === "autoFit" ? "Autofit Column Width" : "Resize Column";
      if (isDefault) {
        this.document.resetColWidth(this.sheetId, docCol, { label, source });
      } else {
        this.document.setColWidth(this.sheetId, docCol, baseSize, { label, source });
      }
      // Shared-grid axis resize updates the CanvasGridRenderer directly during the drag; SpreadsheetApp
      // intentionally skips `syncFrozenPanes()` for these source-tagged sheetView deltas. Ensure any
      // geometry-dependent overlays (drawings/pictures canvas) are still redrawn.
      this.scheduleDrawingsRender();
      return;
    }

    const docRow = change.index - headerRows;
    if (docRow < 0) return;
    const label = change.source === "autoFit" ? "Autofit Row Height" : "Resize Row";
    if (isDefault) {
      this.document.resetRowHeight(this.sheetId, docRow, { label, source });
    } else {
      this.document.setRowHeight(this.sheetId, docRow, baseSize, { label, source });
    }
    this.scheduleDrawingsRender();
  }

  private syncSharedGridInteractionMode(): void {
    const mode = this.formulaBar?.isFormulaEditing() ? "rangeSelection" : "default";
    this.sharedGrid?.setInteractionMode(mode);
  }

  private syncSharedGridReferenceHighlights(): void {
    if (!this.sharedGrid) return;

    if (this.referenceHighlights.length === 0) {
      this.sharedGrid.renderer.setReferenceHighlights(null);
      return;
    }

    const headerRows = this.sharedHeaderRows();
    const headerCols = this.sharedHeaderCols();

    const gridHighlights = this.referenceHighlights.map((highlight) => {
      const startRow = Math.min(highlight.start.row, highlight.end.row);
      const endRow = Math.max(highlight.start.row, highlight.end.row);
      const startCol = Math.min(highlight.start.col, highlight.end.col);
      const endCol = Math.max(highlight.start.col, highlight.end.col);

      const range: GridCellRange = {
        startRow: startRow + headerRows,
        endRow: endRow + headerRows + 1,
        startCol: startCol + headerCols,
        endCol: endCol + headerCols + 1
      };

      return { range, color: highlight.color, active: highlight.active };
    });

    this.sharedGrid.renderer.setReferenceHighlights(gridHighlights);
  }

  private onSharedRangeSelectionStart(range: GridCellRange): void {
    if (!this.formulaBar) return;
    this.syncSharedGridInteractionMode();
    const docRange = this.docRangeFromGridRange(range);
    const rangeSheetId = this.formulaEditCell && this.formulaEditCell.sheetId !== this.sheetId ? this.sheetId : undefined;
    const rangeSheetName = rangeSheetId ? this.resolveSheetDisplayNameById(rangeSheetId) : undefined;
    this.formulaBar.beginRangeSelection({
      start: { row: docRange.startRow, col: docRange.startCol },
      end: { row: docRange.endRow, col: docRange.endCol }
    }, rangeSheetName);
    this.updateEditState();
  }

  private onSharedRangeSelectionChange(range: GridCellRange): void {
    if (!this.formulaBar) return;
    this.syncSharedGridInteractionMode();
    const docRange = this.docRangeFromGridRange(range);
    const rangeSheetId = this.formulaEditCell && this.formulaEditCell.sheetId !== this.sheetId ? this.sheetId : undefined;
    const rangeSheetName = rangeSheetId ? this.resolveSheetDisplayNameById(rangeSheetId) : undefined;
    this.formulaBar.updateRangeSelection({
      start: { row: docRange.startRow, col: docRange.startCol },
      end: { row: docRange.endRow, col: docRange.endCol }
    }, rangeSheetName);
  }

  private onSharedRangeSelectionEnd(): void {
    if (!this.formulaBar) return;
    this.formulaBar.endRangeSelection();
    this.formulaBar.focus();
  }

  private onSharedPointerMove(e: PointerEvent): void {
    if (!this.sharedGrid) return;
    // Shared-grid pointermove listener exists only for hover tooltips. Touch pointers
    // do not have a hover state, so skip all work to avoid unnecessary overhead
    // during touch scroll/pan interactions.
    if (e.pointerType === "touch") {
      if (this.lastHoveredCommentCellKey != null) this.hideCommentTooltip();
      return;
    }
    const hasDrawings = this.drawingObjects.length !== 0;
    const hasComments = this.commentMetaByCoord.size !== 0;
    const hasCharts = this.chartStore.listCharts().some((chart) => chart.sheetId === this.sheetId);

    // Fast path: if the active sheet has no hover-relevant overlays, skip all work.
    if (!hasComments && !hasDrawings && !hasCharts) {
      if (this.sharedHoverCellKey != null || this.sharedHoverCellRect != null) {
        this.clearSharedHoverCellCache();
      }
      if (this.lastHoveredCommentCellKey != null) this.hideCommentTooltip();
      if (this.root.style.cursor) this.root.style.cursor = "";
      return;
    }

    const target = e.target as HTMLElement | null;
    // In both shared and legacy grid modes, pointermoves in the sheet body almost always target
    // the selection canvas. Skip expensive DOM tree walks (`contains`) in that common case.
    if (target && target !== this.selectionCanvas) {
      if (this.vScrollbarTrack.contains(target) || this.hScrollbarTrack.contains(target) || this.outlineLayer.contains(target)) {
        this.clearSharedHoverCellCache();
        this.hideCommentTooltip();
        if (this.root.style.cursor) this.root.style.cursor = "";
        return;
      }
    }
    if (this.commentsPanelVisible) {
      this.hideCommentTooltip();
      if (this.root.style.cursor) this.root.style.cursor = "";
      return;
    }
    if (this.editor.isOpen()) {
      this.hideCommentTooltip();
      if (this.root.style.cursor) this.root.style.cursor = "";
      return;
    }
    if (e.buttons) {
      this.hideCommentTooltip();
      return;
    }

    const useOffsetCoords =
      target === this.root ||
      target === this.selectionCanvas ||
      target === this.gridCanvas ||
      target === this.referenceCanvas ||
      target === this.auditingCanvas ||
      target === this.presenceCanvas;
    // Only refresh cached root position when we need to fall back to client-relative coords.
    // This mirrors the legacy grid path and avoids per-move layout reads for the common case
    // where pointermove targets a full-viewport canvas overlay (selection/grid/etc).
    if (!useOffsetCoords) {
      this.maybeRefreshRootPosition();
    }
    const x = useOffsetCoords ? e.offsetX : e.clientX - this.rootLeft;
    const y = useOffsetCoords ? e.offsetY : e.clientY - this.rootTop;
    if (x < 0 || y < 0 || x > this.width || y > this.height) {
      this.hideCommentTooltip();
      if (this.root.style.cursor) this.root.style.cursor = "";
      return;
    }

    const chartCursor = this.chartCursorAtPoint(x, y);
    const drawingCursor = this.drawingCursorAtPoint(x, y);
    const nextCursor = chartCursor ?? drawingCursor ?? "";
    if (this.root.style.cursor !== nextCursor) {
      this.root.style.cursor = nextCursor;
    }
    // In shared-grid mode, the selection canvas sets its own cursor value, so apply drawing
    // cursor feedback there as well when the pointermove targets the canvas surface.
    const cursorOverride = chartCursor ?? drawingCursor;
    if (cursorOverride && this.selectionCanvas.style.cursor !== cursorOverride) {
      this.selectionCanvas.style.cursor = cursorOverride;
    }

    if (chartCursor) {
      // Charts sit above cell content; suppress comment tooltips while hovering chart bounds.
      this.clearSharedHoverCellCache();
      this.hideCommentTooltip();
      return;
    }

    if (!hasComments) {
      if (this.sharedHoverCellKey != null || this.sharedHoverCellRect != null) {
        this.clearSharedHoverCellCache();
      }
      if (this.lastHoveredCommentCellKey != null) this.hideCommentTooltip();
      return;
    }

    const cachedRect = this.sharedHoverCellRect;
    if (
      cachedRect &&
      x >= cachedRect.x &&
      y >= cachedRect.y &&
      x < cachedRect.x + cachedRect.width &&
      y < cachedRect.y + cachedRect.height
    ) {
      const cellKey = this.sharedHoverCellKey;
      if (cellKey == null) {
        this.hideCommentTooltip();
        return;
      }

      let previewOverride: string | undefined = undefined;
      if (this.sharedHoverCellCommentIndexVersion !== this.commentIndexVersion) {
        previewOverride = this.commentPreviewByCoord.get(cellKey);
        this.sharedHoverCellHasComment = previewOverride !== undefined;
        this.sharedHoverCellCommentIndexVersion = this.commentIndexVersion;
      }

      if (!this.sharedHoverCellHasComment) {
        if (this.lastHoveredCommentCellKey != null) this.hideCommentTooltip();
        return;
      }

      if (
        this.lastHoveredCommentCellKey === cellKey &&
        this.lastHoveredCommentIndexVersion === this.commentIndexVersion
      ) {
        return;
      }

      const preview = previewOverride ?? (this.commentPreviewByCoord.get(cellKey) ?? "");

      this.lastHoveredCommentCellKey = cellKey;
      this.lastHoveredCommentIndexVersion = this.commentIndexVersion;
      this.commentTooltip.textContent = preview;
      this.commentTooltip.style.setProperty("--comment-tooltip-x", `${x + 12}px`);
      this.commentTooltip.style.setProperty("--comment-tooltip-y", `${y + 12}px`);
      this.commentTooltipVisible = true;
      this.commentTooltip.classList.add("comment-tooltip--visible");
      return;
    }

    // Comment tooltips only apply to the sheet body (not row/col headers).
    // Prefer the renderer sizes in case header indices are ever resized.
    const headerWidth = this.sharedGrid.renderer.getColWidth(0);
    const headerHeight = this.sharedGrid.renderer.getRowHeight(0);
    if (x < headerWidth || y < headerHeight) {
      if (this.lastHoveredCommentCellKey != null) this.hideCommentTooltip();
      return;
    }

    const picked = this.sharedGrid.renderer.pickCellAt(x, y);
    if (!picked) {
      this.hideCommentTooltip();
      this.clearSharedHoverCellCache();
      return;
    }
    const cellRect = this.sharedGrid.getCellRect(picked.row, picked.col);
    if (!cellRect) {
      this.hideCommentTooltip();
      this.clearSharedHoverCellCache();
      return;
    }
    this.sharedHoverCellRect = cellRect;
    const headerRows = 1;
    const headerCols = 1;
    if (picked.row < headerRows || picked.col < headerCols) {
      this.sharedHoverCellKey = null;
      this.sharedHoverCellHasComment = false;
      this.hideCommentTooltip();
      return;
    }

    const docRow = picked.row - headerRows;
    const docCol = picked.col - headerCols;
    const cellKey = docRow * COMMENT_COORD_COL_STRIDE + docCol;
    this.sharedHoverCellKey = cellKey;
    this.sharedHoverCellCommentIndexVersion = this.commentIndexVersion;
    const preview = this.commentPreviewByCoord.get(cellKey);
    this.sharedHoverCellHasComment = preview !== undefined;
    if (preview === undefined) {
      this.hideCommentTooltip();
      return;
    }

    if (
      this.lastHoveredCommentCellKey === cellKey &&
      this.lastHoveredCommentIndexVersion === this.commentIndexVersion
    ) {
      return;
    }
    this.lastHoveredCommentCellKey = cellKey;
    this.lastHoveredCommentIndexVersion = this.commentIndexVersion;
    this.commentTooltip.textContent = preview;
    this.commentTooltip.style.setProperty("--comment-tooltip-x", `${x + 12}px`);
    this.commentTooltip.style.setProperty("--comment-tooltip-y", `${y + 12}px`);
    this.commentTooltipVisible = true;
    this.commentTooltip.classList.add("comment-tooltip--visible");
  }

  goTo(reference: string): boolean {
    try {
      const trimmed = reference.trim();
      // `parseGoTo` needs a sheet token it can round-trip through the workbook lookup.
      // Use the stable id here so unqualified A1 references ("A1") still work even if
      // display-name -> id resolution is temporarily unavailable (e.g. during sheet
      // rename propagation).
      const currentSheetName = this.sheetId;
      const { sheetName: qualifiedSheetName, ref: rawRef } = splitSheetQualifier(trimmed);

      // Excel-style row/column range go-to (e.g. "A:A", "A:C", "1:1", "1:10").
      // The search package's `parseGoTo` intentionally only parses A1-style cell refs, so handle
      // these shorthands at the app layer where we have sheet limits.
      const ref = rawRef.trim();
      const targetSheetId = qualifiedSheetName ? this.resolveSheetIdByName(qualifiedSheetName) : this.sheetId;
      if (!targetSheetId) return false;

      const colRange = /^(\$?[A-Za-z]{1,3})\s*:\s*(\$?[A-Za-z]{1,3})$/.exec(ref);
      if (colRange) {
        const parseCol = (token: string): number => fromA1A1(`${token}1`).col0;
        const a = parseCol(colRange[1]!);
        const b = parseCol(colRange[2]!);
        const startCol = Math.min(a, b);
        const endCol = Math.max(a, b);
        this.selectRange({
          sheetId: targetSheetId,
          range: { startRow: 0, endRow: this.limits.maxRows - 1, startCol, endCol },
        });
        return true;
      }

      const rowRange = /^(\$?\d+)\s*:\s*(\$?\d+)$/.exec(ref);
      if (rowRange) {
        const parseRow = (token: string): number => {
          const rawRow = Number.parseInt(token.replaceAll("$", ""), 10);
          if (!Number.isFinite(rawRow) || rawRow < 1) return Number.NaN;
          return rawRow - 1;
        };
        const a = parseRow(rowRange[1]!);
        const b = parseRow(rowRange[2]!);
        if (!Number.isFinite(a) || !Number.isFinite(b)) return false;
        const startRow = Math.min(a, b);
        const endRow = Math.max(a, b);
        this.selectRange({
          sheetId: targetSheetId,
          range: { startRow, endRow, startCol: 0, endCol: this.limits.maxCols - 1 },
        });
        return true;
      }

      // When multiple disjoint ranges are selected, the Name Box shows a stable label like
      // "2 ranges". Treat pressing Enter on that label as a no-op navigation (Excel behavior is
      // effectively a focus-return, not an "invalid reference" error).
      const multiRangesLabel = `${this.selection.ranges.length} ranges`;
      if (!qualifiedSheetName && this.selection.ranges.length > 1 && ref === multiRangesLabel) {
        this.focus();
        return true;
      }

      const parsed = parseGoTo(trimmed, { workbook: this.searchWorkbook, currentSheetName });
      if (parsed.type !== "range") return false;

      const { range } = parsed;
      // For unqualified A1 references (e.g. "A1" or "A1:B2"), always treat navigation as relative
      // to the *current* stable sheet id (not the display name), even if the sheet metadata
      // resolver is temporarily unavailable/out-of-date.
      // `targetSheetId` is resolved above for row/column refs; recompute for A1/table/name refs.
      const targetSheetIdForParsed =
        parsed.source === "a1" && !qualifiedSheetName ? this.sheetId : this.resolveSheetIdByName(parsed.sheetName);
      if (!targetSheetIdForParsed) return false;
      if (range.startRow === range.endRow && range.startCol === range.endCol) {
        this.activateCell({ sheetId: targetSheetIdForParsed, row: range.startRow, col: range.startCol });
      } else {
        this.selectRange({ sheetId: targetSheetIdForParsed, range });
      }
      return true;
    } catch {
      // Invalid Go To inputs should not throw; signal failure so the name box can
      // present "invalid reference" feedback instead of failing silently.
      return false;
    }
  }

  private async openNameBoxMenu(): Promise<void> {
    const formatSheetPrefix = (sheetName: string): string => {
      const token = formatSheetNameForA1(sheetName);
      return token ? `${token}!` : "";
    };

    const normalizeDocRange = (range: any): Range | null => {
      if (!range) return null;
      const { startRow, endRow, startCol, endCol } = range as any;
      if (
        !Number.isInteger(startRow) ||
        !Number.isInteger(endRow) ||
        !Number.isInteger(startCol) ||
        !Number.isInteger(endCol) ||
        startRow < 0 ||
        endRow < 0 ||
        startCol < 0 ||
        endCol < 0
      ) {
        return null;
      }

      return {
        startRow: Math.min(startRow, endRow),
        endRow: Math.max(startRow, endRow),
        startCol: Math.min(startCol, endCol),
        endCol: Math.max(startCol, endCol),
      };
    };

    const items: Array<{ label: string; value: string; description?: string }> = [];

    for (const entry of this.searchWorkbook.names.values()) {
      const name = typeof (entry as any)?.name === "string" ? ((entry as any).name as string).trim() : "";
      if (!name) continue;

      const range = normalizeDocRange((entry as any)?.range);
      const sheetName = typeof (entry as any)?.sheetName === "string" ? ((entry as any).sheetName as string).trim() : "";
      const a1 =
        range ? rangeToA1(range) : null;

      items.push({
        label: name,
        value: name,
        description: a1 ? (sheetName ? `${sheetName}!${a1}` : a1) : undefined,
      });
    }

    for (const table of this.searchWorkbook.tables.values()) {
      const name = typeof (table as any)?.name === "string" ? ((table as any).name as string).trim() : "";
      if (!name) continue;

      const sheetName = typeof (table as any)?.sheetName === "string" ? ((table as any).sheetName as string).trim() : "";
      const range = normalizeDocRange({
        startRow: (table as any)?.startRow,
        startCol: (table as any)?.startCol,
        endRow: (table as any)?.endRow,
        endCol: (table as any)?.endCol,
      });
      const a1 = range ? rangeToA1(range) : null;

      const structuredOk = /^[A-Za-z_][A-Za-z0-9_]*$/.test(name);
      // `parseGoTo` only understands structured refs for tables with identifier-like names.
      // Fall back to sheet-qualified A1 when we have valid bounds.
      const value = structuredOk ? `${name}[#All]` : sheetName && a1 ? `${formatSheetPrefix(sheetName)}${a1}` : null;
      if (!value) continue;

      items.push({
        label: name,
        value,
        description: a1 ? (sheetName ? `${sheetName}!${a1}` : a1) : undefined,
      });
    }

    items.sort((a, b) => a.label.localeCompare(b.label, undefined, { sensitivity: "base" }));

    if (items.length === 0) return;

    const selected = await showQuickPick(items, { placeHolder: "Go to…" });
    if (typeof selected !== "string" || selected.trim() === "") return;
    this.goTo(selected);
  }

  private resolveRemoteUserLabel(userId: string): string {
    const id = String(userId ?? "");
    if (!id) return id;
    const presence = this.collabSession?.presence;
    if (!presence) return id;

    try {
      const presences = presence.getRemotePresences({ includeOtherSheets: true }) as any[];
      for (const entry of presences) {
        if (String(entry?.id ?? "") !== id) continue;
        const name = String(entry?.name ?? "");
        if (name) return name;
      }
    } catch {
      // Best-effort: presence snapshots can fail if the collab session is tearing down.
    }

    return id;
  }

  private navigateToConflictCell(cellRef: { sheetId: string; row: number; col: number }): void {
    const sheetId = String(cellRef?.sheetId ?? "");
    const row = Number(cellRef?.row);
    const col = Number(cellRef?.col);
    if (!sheetId) return;
    if (!Number.isFinite(row) || !Number.isFinite(col)) return;

    const safeRow = Math.max(0, Math.min(this.limits.maxRows - 1, Math.trunc(row)));
    const safeCol = Math.max(0, Math.min(this.limits.maxCols - 1, Math.trunc(col)));

    // Avoid stealing focus from the conflict dialog itself.
    this.activateCell({ sheetId, row: safeRow, col: safeCol }, { scrollIntoView: false, focus: false });
    // In legacy mode, the target cell may be hidden by outline state. Remap to the nearest
    // visible cell so the user doesn't "jump" to an invisible selection.
    this.ensureActiveCellVisible();

    const target = { ...this.selection.active };
    const didScroll = this.scrollCellToCenter(target);
    // Legacy renderer needs an explicit redraw for scroll changes.
    if (!this.sharedGrid && didScroll) {
      this.refresh("scroll");
    } else {
      // Even when we didn't scroll, `ensureActiveCellVisible` may have remapped the active cell.
      this.renderSelection();
    }
    this.updateStatus();
  }

  async getCellValueA1(a1: string): Promise<string> {
    await this.whenIdle();
    const cell = parseA1(a1);
    return this.getCellDisplayValue(cell);
  }

  getFillHandleRect(): { x: number; y: number; width: number; height: number } | null {
    if (this.formulaBar?.isFormulaEditing()) return null;
    if (this.sharedGrid) {
      return this.sharedGrid.renderer.getFillHandleRect();
    }
    this.ensureViewportMappingCurrent();
    return this.selectionRenderer.getFillHandleRect(
      this.selection,
      {
        getCellRect: (cell) => this.getCellRect(cell),
        visibleRows: this.visibleRows,
        visibleCols: this.visibleCols,
      },
      {
        clipRect: {
          x: this.rowHeaderWidth,
          y: this.colHeaderHeight,
          width: this.viewportWidth(),
          height: this.viewportHeight(),
        },
      }
    );
  }

  private getDrawingHitTestIndex(objects: readonly DrawingObject[]): HitTestIndex {
    const zoom = this.getZoom();
    const cached = this.drawingHitTestIndex;
    if (cached && this.drawingHitTestIndexObjects === objects && Math.abs(cached.zoom - zoom) < 1e-6) return cached;
    const index = buildHitTestIndex(objects, this.drawingGeom, { zoom });
    this.drawingHitTestIndex = index;
    this.drawingHitTestIndexObjects = objects;
    return index;
  }

  private drawingCursorAtPoint(x: number, y: number): string | null {
    const objects = this.drawingObjects;
    if (objects.length === 0) return null;

    const viewport = this.getDrawingInteractionViewport();

    const selectedId = this.selectedDrawingId;
    if (selectedId != null) {
      const selected = objects.find((obj) => obj.id === selectedId) ?? null;
      if (selected) {
        const bounds = drawingObjectToViewportRect(selected, viewport, this.drawingGeom);
        if (hitTestRotationHandle(bounds, x, y, selected.transform)) return cursorForRotationHandle(false);
        // When selected, handles can extend slightly outside the untransformed bounds. Check the selected
        // object explicitly so hover feedback works even when the cursor lies just beyond the anchor rect.
        const handle = hitTestResizeHandle(bounds, x, y, selected.transform);
        if (handle) return cursorForResizeHandleWithTransform(handle, selected.transform);
      }
    }

    const index = this.getDrawingHitTestIndex(objects);
    const hit = hitTestDrawings(index, viewport, x, y);
    if (!hit) return null;
    const handle = hitTestResizeHandle(hit.bounds, x, y, hit.object.transform);
    if (handle) return cursorForResizeHandleWithTransform(handle, hit.object.transform);
    return "move";
  }

  private chartCursorAtPoint(x: number, y: number): string | null {
    const charts = this.chartStore.listCharts().filter((chart) => chart.sheetId === this.sheetId);
    if (charts.length === 0) return null;

    const layout = this.chartOverlayLayout(this.sharedGrid ? this.sharedGrid.renderer.scroll.getViewportState() : undefined);
    const px = x - layout.originX;
    const py = y - layout.originY;
    if (!Number.isFinite(px) || !Number.isFinite(py)) return null;
    if (px < 0 || py < 0) return null;

    const intersect = (
      a: { left: number; top: number; width: number; height: number },
      b: { left: number; top: number; width: number; height: number },
    ): { left: number; top: number; width: number; height: number } | null => {
      const left = Math.max(a.left, b.left);
      const top = Math.max(a.top, b.top);
      const right = Math.min(a.left + a.width, b.left + b.width);
      const bottom = Math.min(a.top + a.height, b.top + b.height);
      const width = right - left;
      const height = bottom - top;
      if (width <= 0 || height <= 0) return null;
      return { left, top, width, height };
    };

    const { frozenRows, frozenCols } = this.getFrozen();
    const selectedId = this.selectedChartId;

    for (let i = charts.length - 1; i >= 0; i -= 1) {
      const chart = charts[i]!;
      const rect = this.chartAnchorToViewportRect(chart.anchor);
      if (!rect) continue;

      const fromRow = chart.anchor.kind === "oneCell" || chart.anchor.kind === "twoCell" ? chart.anchor.fromRow : Number.POSITIVE_INFINITY;
      const fromCol = chart.anchor.kind === "oneCell" || chart.anchor.kind === "twoCell" ? chart.anchor.fromCol : Number.POSITIVE_INFINITY;
      const inFrozenRows = fromRow < frozenRows;
      const inFrozenCols = fromCol < frozenCols;
      const paneKey: "topLeft" | "topRight" | "bottomLeft" | "bottomRight" =
        inFrozenRows && inFrozenCols
          ? "topLeft"
          : inFrozenRows && !inFrozenCols
            ? "topRight"
            : !inFrozenRows && inFrozenCols
              ? "bottomLeft"
              : "bottomRight";
      const paneRect = layout.paneRects[paneKey];
      const visible = intersect(rect, { left: paneRect.x, top: paneRect.y, width: paneRect.width, height: paneRect.height });
      if (!visible) continue;
      if (px < visible.left || px > visible.left + visible.width) continue;
      if (py < visible.top || py > visible.top + visible.height) continue;

      if (selectedId === chart.id) {
        const handle = hitTestResizeHandle(
          { x: rect.left, y: rect.top, width: rect.width, height: rect.height },
          px,
          py,
        );
        if (handle) return cursorForResizeHandle(handle);
      }

      return "move";
    }

    return null;
  }

  async getCellDisplayValueA1(a1: string): Promise<string> {
    return this.getCellValueA1(a1);
  }

  async getCellDisplayTextForRenderA1(a1: string): Promise<string> {
    await this.whenIdle();
    const cell = parseA1(a1);
    const state = this.document.getCell(this.sheetId, cell) as { value: unknown; formula: string | null };
    if (!state) return "";

    if (state.formula != null) {
      if (this.showFormulas) return state.formula;
      const computed = this.getCellComputedValue(cell);
      return computed == null ? "" : this.formatCellValueForDisplay(cell, computed);
    }

    if (isRichTextValue(state.value)) return state.value.text;
    if (state.value != null) return this.formatCellValueForDisplay(cell, state.value as any);
    return "";
  }

  getLastSelectionDrawn(): unknown {
    if (this.sharedGrid) {
      // Shared-grid mode renders selection via CanvasGridRenderer, not the legacy SelectionRenderer.
      // Preserve the existing "selection debug" API for e2e tests by returning an equivalent
      // bounding-rect representation in viewport coordinates.
      const ranges = this.selection.ranges
        .map((range) => {
          const start = this.getCellRect({ row: range.startRow, col: range.startCol });
          const end = this.getCellRect({ row: range.endRow, col: range.endCol });
          if (!start || !end) return null;
          const x = Math.min(start.x, end.x);
          const y = Math.min(start.y, end.y);
          const width = Math.max(start.x + start.width, end.x + end.width) - x;
          const height = Math.max(start.y + start.height, end.y + end.height) - y;
          if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return null;
          return {
            range: { ...range },
            rect: { x, y, width, height },
            edges: { top: true, right: true, bottom: true, left: true }
          };
        })
        .filter((r): r is NonNullable<typeof r> => r !== null);

      return {
        ranges,
        activeCellRect: this.getCellRect(this.selection.active),
      };
    }

    return this.selectionRenderer.getLastDebug();
  }

  getReferenceHighlightCount(): number {
    return this.referenceHighlights.length;
  }

  getAuditingHighlights(): {
    mode: "off" | AuditingMode;
    transitive: boolean;
    precedents: string[];
    dependents: string[];
    errors: { precedents: string | null; dependents: string | null };
  } {
    return {
      mode: this.auditingMode,
      transitive: this.auditingTransitive,
      precedents: Array.from(this.auditingHighlights.precedents).sort(),
      dependents: Array.from(this.auditingHighlights.dependents).sort(),
      errors: { ...this.auditingErrors },
    };
  }

  toggleAuditingPrecedents(): void {
    this.toggleAuditingComponent("precedents");
  }

  toggleAuditingDependents(): void {
    this.toggleAuditingComponent("dependents");
  }

  toggleAuditingTransitive(): void {
    this.auditingTransitive = !this.auditingTransitive;
    this.updateAuditingLegend();
    this.scheduleAuditingUpdate();
  }

  clearAuditing(): void {
    this.auditingMode = "off";
    this.auditingHighlights = { precedents: new Set(), dependents: new Set() };
    this.auditingErrors = { precedents: null, dependents: null };
    this.updateAuditingLegend();
    this.renderAuditing();
  }

  isCommentsPanelVisible(): boolean {
    return this.commentsPanelVisible;
  }

  /**
   * Opens the comments panel (idempotent) and focuses the new-comment input.
   */
  openCommentsPanel(): void {
    if (!this.commentsPanelVisible) {
      this.toggleCommentsPanel();
      return;
    }
    this.focusNewCommentInput();
  }

  /**
   * Closes the comments panel (idempotent).
   */
  closeCommentsPanel(): void {
    if (!this.commentsPanelVisible) return;
    this.toggleCommentsPanel();
  }

  /**
   * Best-effort focus for the "new comment" input.
   */
  focusNewCommentInput(): void {
    // Viewer roles can open the comments panel to read, but should not be routed
    // into disabled composer UI.
    if (!this.canUserComment()) return;
    try {
      const input =
        this.root.querySelector<HTMLInputElement>('[data-testid="new-comment-input"]') ??
        (this.newCommentInput as HTMLInputElement | undefined);
      input?.focus();
    } catch {
      // ignore
    }
  }

  /**
   * Returns true if the active cell currently has at least one comment thread.
   */
  activeCellHasComment(): boolean {
    const cell = this.selection.active;
    return Boolean(this.commentMetaByCoord.get(cell.row * COMMENT_COORD_COL_STRIDE + cell.col));
  }

  openInlineAiEdit(): void {
    // Match the Cmd/Ctrl+K guard behavior (see `onKeyDown`).
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.inlineEditController.isOpen()) return;
    if (this.editor.isOpen()) return;
    // Inline edit should not trigger while the formula bar is actively editing.
    if (this.formulaBar?.isEditing() || this.formulaEditCell) return;
    this.inlineEditController.open();
    this.updateEditState();
  }

  toggleCommentsPanel(): void {
    this.commentsPanelVisible = !this.commentsPanelVisible;
    this.commentsPanel.classList.toggle("comments-panel--visible", this.commentsPanelVisible);
    if (this.commentsPanelVisible) {
      this.renderCommentsPanel();
      this.focusNewCommentInput();
    } else {
      this.focus();
    }

    // Used by the Ribbon to sync pressed state for view toggles when opened/closed
    // outside of Ribbon interactions (e.g. keyboard shortcuts, debug buttons).
    this.dispatchViewChanged();

    // Broadcast for consumers like extension context keys (best-effort).
    if (typeof window !== "undefined") {
      window.dispatchEvent(
        new CustomEvent("formula:comments-panel-visibility-changed", { detail: { visible: this.commentsPanelVisible } }),
      );
    }
  }

  private createCommentsPanel(): HTMLDivElement {
    const panel = document.createElement("div");
    panel.dataset.testid = "comments-panel";
    panel.className = "comments-panel";

    panel.addEventListener("pointerdown", (e) => e.stopPropagation());
    panel.addEventListener("dblclick", (e) => e.stopPropagation());
    panel.addEventListener("keydown", (e) => e.stopPropagation());

    const header = document.createElement("div");
    header.className = "comments-panel__header";

    const title = document.createElement("div");
    title.textContent = t("comments.title");
    title.className = "comments-panel__title";

    const closeButton = document.createElement("button");
    closeButton.textContent = "×";
    closeButton.type = "button";
    closeButton.className = "comments-panel__close-button";
    closeButton.setAttribute("aria-label", t("comments.closePanel"));
    closeButton.addEventListener("click", () => this.toggleCommentsPanel());

    header.appendChild(title);
    header.appendChild(closeButton);
    panel.appendChild(header);

    this.commentsPanelCell = document.createElement("div");
    this.commentsPanelCell.dataset.testid = "comments-active-cell";
    this.commentsPanelCell.className = "comments-panel__active-cell";
    panel.appendChild(this.commentsPanelCell);

    this.commentsPanelThreads = document.createElement("div");
    this.commentsPanelThreads.className = "comments-panel__threads";
    panel.appendChild(this.commentsPanelThreads);

    const footer = document.createElement("div");
    footer.className = "comments-panel__footer";

    this.commentsPanelReadOnlyHint = document.createElement("div");
    this.commentsPanelReadOnlyHint.dataset.testid = "comments-readonly-hint";
    this.commentsPanelReadOnlyHint.className = "comments-panel__readonly-hint";
    this.commentsPanelReadOnlyHint.textContent = t("comments.readOnlyHint");
    this.commentsPanelReadOnlyHint.hidden = true;
    footer.appendChild(this.commentsPanelReadOnlyHint);

    const footerRow = document.createElement("div");
    footerRow.className = "comments-panel__footer-row";

    this.newCommentInput = document.createElement("input");
    this.newCommentInput.dataset.testid = "new-comment-input";
    this.newCommentInput.type = "text";
    this.newCommentInput.placeholder = t("comments.new.placeholder");
    this.newCommentInput.className = "comments-panel__new-comment-input";

    this.newCommentSubmitButton = document.createElement("button");
    this.newCommentSubmitButton.dataset.testid = "submit-comment";
    this.newCommentSubmitButton.textContent = t("comments.new.submit");
    this.newCommentSubmitButton.type = "button";
    this.newCommentSubmitButton.className = "comments-panel__submit-button";
    this.newCommentSubmitButton.addEventListener("click", () => this.submitNewComment());

    footerRow.appendChild(this.newCommentInput);
    footerRow.appendChild(this.newCommentSubmitButton);
    footer.appendChild(footerRow);
    panel.appendChild(footer);

    return panel;
  }

  private createCommentTooltip(): HTMLDivElement {
    const tooltip = document.createElement("div");
    tooltip.dataset.testid = "comment-tooltip";
    tooltip.className = "comment-tooltip";
    return tooltip;
  }

  private createFormulaRangePreviewTooltip(): HTMLDivElement {
    const tooltip = document.createElement("div");
    tooltip.dataset.testid = "formula-range-preview-tooltip";
    tooltip.className = "formula-range-preview-tooltip";
    tooltip.hidden = true;
    tooltip.setAttribute("role", "tooltip");
    tooltip.setAttribute("aria-hidden", "true");
    tooltip.id = nextFormulaRangePreviewTooltipId();
    return tooltip;
  }

  private createAuditingLegend(): HTMLDivElement {
    const legend = document.createElement("div");
    legend.dataset.testid = "auditing-legend";
    legend.className = "auditing-legend";
    legend.hidden = true;
    return legend;
  }

  private updateAuditingLegend(): void {
    const legend = this.auditingLegend;

    if (this.auditingMode === "off") {
      legend.hidden = true;
      legend.replaceChildren();
      return;
    }

    const wantsPrecedents = this.auditingMode === "precedents" || this.auditingMode === "both";
    const wantsDependents = this.auditingMode === "dependents" || this.auditingMode === "both";

    legend.replaceChildren();

    const makeModeItem = (opts: { kind: "precedents" | "dependents"; label: string }): HTMLElement => {
      const item = document.createElement("span");
      item.className = "auditing-legend__item";

      const swatch = document.createElement("span");
      swatch.className = `auditing-legend__swatch auditing-legend__swatch--${opts.kind}`;

      const text = document.createElement("span");
      text.textContent = opts.label;

      item.appendChild(swatch);
      item.appendChild(text);
      return item;
    };

    const transitive = this.auditingTransitive ? "Transitive" : "Direct";

    const errors: string[] = [];
    if (wantsPrecedents && this.auditingErrors.precedents) errors.push(`Precedents: ${this.auditingErrors.precedents}`);
    if (wantsDependents && this.auditingErrors.dependents) errors.push(`Dependents: ${this.auditingErrors.dependents}`);

    const row = document.createElement("div");
    row.className = "auditing-legend__row";

    const modeRow = document.createElement("div");
    modeRow.className = "auditing-legend__modes";
    if (wantsPrecedents) modeRow.appendChild(makeModeItem({ kind: "precedents", label: "Precedents" }));
    if (wantsDependents) modeRow.appendChild(makeModeItem({ kind: "dependents", label: "Dependents" }));
    row.appendChild(modeRow);

    const transitiveEl = document.createElement("span");
    transitiveEl.className = "auditing-legend__transitive";
    transitiveEl.textContent = `(${transitive})`;
    row.appendChild(transitiveEl);

    legend.appendChild(row);

    if (errors.length > 0) {
      const errorsEl = document.createElement("div");
      errorsEl.className = "auditing-legend__errors";
      errorsEl.textContent = errors.join("\n");
      legend.appendChild(errorsEl);
    }

    legend.hidden = false;
  }

  private hideCommentTooltip(): void {
    // Some unit tests construct a SpreadsheetApp instance by `Object.create(SpreadsheetApp.prototype)`
    // and only populate a subset of fields. Be defensive so helper methods can run without the full
    // DOM scaffold created by the constructor.
    if (!this.commentTooltip) {
      this.lastHoveredCommentCellKey = null;
      this.lastHoveredCommentIndexVersion = -1;
      this.commentTooltipVisible = false;
      return;
    }
    if (
      this.lastHoveredCommentCellKey == null &&
      !this.commentTooltipVisible
    ) {
      return;
    }

    this.lastHoveredCommentCellKey = null;
    this.lastHoveredCommentIndexVersion = -1;
    this.commentTooltipVisible = false;
    this.commentTooltip.classList.remove("comment-tooltip--visible");
  }

  private hideFormulaRangePreviewTooltip(): void {
    const tooltip = this.formulaRangePreviewTooltip;
    if (!tooltip) {
      this.formulaRangePreviewTooltipVisible = false;
      this.formulaRangePreviewTooltipLastKey = null;
      return;
    }
    if (!this.formulaRangePreviewTooltipVisible && tooltip.hidden) return;
    this.formulaRangePreviewTooltipVisible = false;
    tooltip.hidden = true;
    tooltip.setAttribute("aria-hidden", "true");
    tooltip.classList.remove("formula-range-preview-tooltip--visible");
    this.syncFormulaRangePreviewTooltipDescribedBy(false);
  }

  private syncFormulaRangePreviewTooltipDescribedBy(visible: boolean): void {
    const tooltip = this.formulaRangePreviewTooltip;
    if (!tooltip?.id) return;
    const textarea = this.formulaBar?.textarea;
    if (!textarea) return;

    const describedBy = textarea.getAttribute("aria-describedby") ?? "";
    const tokens = describedBy
      .split(/\s+/)
      .map((t) => t.trim())
      .filter(Boolean);

    const idx = tokens.indexOf(tooltip.id);
    if (visible) {
      if (idx >= 0) return;
      tokens.push(tooltip.id);
      textarea.setAttribute("aria-describedby", tokens.join(" "));
      return;
    }

    if (idx < 0) return;
    tokens.splice(idx, 1);
    const next = tokens.join(" ");
    if (next) textarea.setAttribute("aria-describedby", next);
    else textarea.removeAttribute("aria-describedby");
  }

  private resolveFormulaRangePreviewTargetSheet(refText: string | null): { explicit: boolean; sheetId: string | null } {
    const rawText = typeof refText === "string" ? refText.trim() : "";
    if (!rawText) return { explicit: false, sheetId: null };

    const { sheetName } = splitSheetQualifier(rawText);
    if (sheetName) {
      return { explicit: true, sheetId: this.resolveSheetIdByName(sheetName) };
    }

    // Named ranges can point at a specific sheet (even though the identifier itself is unqualified).
    // When possible, resolve the name so we can avoid previewing the wrong sheet.
    const entry: any = this.searchWorkbook.getName(rawText);
    const nameSheet = typeof entry?.sheetName === "string" ? entry.sheetName.trim() : "";
    if (nameSheet) {
      return { explicit: true, sheetId: this.resolveSheetIdByName(nameSheet) };
    }

    // Structured table refs (e.g. `Table1[Amount]`) can also belong to a specific sheet.
    const bracket = rawText.indexOf("[");
    if (bracket > 0) {
      const tableName = rawText.slice(0, bracket).trim();
      if (tableName) {
        const table: any = this.searchWorkbook.getTable(tableName);
        const tableSheet =
          typeof table?.sheetName === "string"
            ? table.sheetName.trim()
            : typeof table?.sheet === "string"
              ? table.sheet.trim()
              : "";
        if (tableSheet) {
          return { explicit: true, sheetId: this.resolveSheetIdByName(tableSheet) };
        }
      }
    }

    // Unqualified A1 references (e.g. `A1:B2`) are relative to the sheet containing the formula.
    // When the formula bar is editing a cell on another sheet (Excel-style range selection mode),
    // avoid previewing/highlighting them on the *currently active* sheet.
    if (this.formulaEditCell?.sheetId) {
      return { explicit: true, sheetId: this.formulaEditCell.sheetId };
    }

    return { explicit: false, sheetId: null };
  }

  private isFormulaRangePreviewAllowed(refText: string | null): boolean {
    const { explicit, sheetId } = this.resolveFormulaRangePreviewTargetSheet(refText);
    if (!explicit) return true;
    if (!sheetId) return false;
    return sheetId.toLowerCase() === this.sheetId.toLowerCase();
  }

  private syncFormulaBarHoverRangeOverlays(): void {
    const bar = this.formulaBar;
    if (!bar || !bar.isEditing()) return;

    const range = bar.model.hoveredReference();
    const refText = bar.model.hoveredReferenceText();
    const allowed = this.isFormulaRangePreviewAllowed(refText);

    if (!allowed) {
      this.hideFormulaRangePreviewTooltip();
      this.referencePreview = null;
      if (this.sharedGrid) {
        this.sharedGrid.clearRangeSelection();
      } else {
        this.renderReferencePreview();
      }
      return;
    }

    this.updateFormulaRangePreviewTooltip(range, refText);

    this.referencePreview = range
      ? {
          start: { row: range.start.row, col: range.start.col },
          end: { row: range.end.row, col: range.end.col },
        }
      : null;

    if (this.sharedGrid) {
      if (range) {
        const gridRange = this.gridRangeFromDocRange({
          startRow: range.start.row,
          endRow: range.end.row,
          startCol: range.start.col,
          endCol: range.end.col,
        });
        this.sharedGrid.setRangeSelection(gridRange);
      } else {
        this.sharedGrid.clearRangeSelection();
      }
      return;
    }

    this.renderReferencePreview();
  }

  private updateFormulaRangePreviewTooltip(range: A1RangeAddress | null, refText: string | null): void {
    const tooltip = this.formulaRangePreviewTooltip;
    if (!tooltip) return;

    if (!range) {
      this.hideFormulaRangePreviewTooltip();
      return;
    }

    if (!this.isFormulaRangePreviewAllowed(refText)) {
      this.hideFormulaRangePreviewTooltip();
      return;
    }

    const rawText = typeof refText === "string" ? refText.trim() : "";
    const { sheetName } = splitSheetQualifier(rawText);

    const startRow = Math.min(range.start.row, range.end.row);
    const endRow = Math.max(range.start.row, range.end.row);
    const startCol = Math.min(range.start.col, range.end.col);
    const endCol = Math.max(range.start.col, range.end.col);
    const rowCount = endRow - startRow + 1;
    const colCount = endCol - startCol + 1;
    if (rowCount <= 0 || colCount <= 0) {
      this.hideFormulaRangePreviewTooltip();
      return;
    }

    const totalCells = rowCount * colCount;
    const tooLarge = totalCells > MAX_FORMULA_RANGE_PREVIEW_CELLS;

    const label = (() => {
      if (rawText && !sheetName) {
        const entry: any = this.searchWorkbook.getName(rawText);
        const r = entry?.range;
        if (
          r &&
          typeof r.startRow === "number" &&
          typeof r.startCol === "number" &&
          typeof r.endRow === "number" &&
          typeof r.endCol === "number"
        ) {
          return `${rawText} (${rangeToA1(r)})`;
        }

        // Structured table refs: show the resolved A1 address for context (similar to named ranges).
        const bracket = rawText.indexOf("[");
        if (bracket > 0) {
          const tableName = rawText.slice(0, bracket).trim();
          if (tableName && this.searchWorkbook.getTable(tableName)) {
            return `${rawText} (${rangeToA1({ startRow, endRow, startCol, endCol })})`;
          }
        }
      }
      return (
        rawText ||
        rangeToA1({
          startRow,
          endRow,
          startCol,
          endCol,
        })
      );
    })();

    // Cache key: avoid re-reading cells on every mousemove while hovering the same reference span.
    // Include a monotonic document version so the tooltip can refresh after any workbook update
    // (values/formulas/formatting/view changes, including edits on other sheets that may affect
    // computed values in the active sheet).
    const docUpdateVersion = this.document.updateVersion;
    const key = `${this.sheetId}:${docUpdateVersion}:${startRow},${startCol}:${endRow},${endCol}:${label}:${tooLarge ? "L" : "S"}`;
    if (this.formulaRangePreviewTooltipVisible && this.formulaRangePreviewTooltipLastKey === key) {
      return;
    }
    this.formulaRangePreviewTooltipLastKey = key;

    tooltip.replaceChildren();

    const header = document.createElement("div");
    header.className = "formula-range-preview-tooltip__header";
    header.textContent = label;
    tooltip.appendChild(header);

    const sampleRows = Math.min(rowCount, FORMULA_RANGE_PREVIEW_SAMPLE_ROWS);
    const sampleCols = Math.min(colCount, FORMULA_RANGE_PREVIEW_SAMPLE_COLS);

    // Materialize a small sample grid of *displayed* values (formatted numbers, etc).
    const sampleValues = Array.from({ length: sampleRows }, () => Array<string>(sampleCols).fill(""));

    const table = document.createElement("table");
    table.className = "formula-range-preview-tooltip__grid";
    const tbody = document.createElement("tbody");
    table.appendChild(tbody);

    const coordScratch = { row: 0, col: 0 };
    const summary = document.createElement("div");
    summary.className = "formula-range-preview-tooltip__summary";

    const formatDisplay = (computed: SpreadsheetValue): string => {
      if (computed == null) return "";
      return this.formatCellValueForDisplay(coordScratch, computed);
    };

    // Keep read cost bounded:
    // - For large ranges, only read the small sample.
    // - For small ranges (<= MAX_FORMULA_RANGE_PREVIEW_CELLS), enumerate each cell at most once,
    //   collecting both sample display strings and numeric summary stats.
    if (tooLarge) {
      const formatter =
        this.selectionStatsFormatter ??
        (this.selectionStatsFormatter = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 }));
      for (let r = 0; r < sampleRows; r += 1) {
        for (let c = 0; c < sampleCols; c += 1) {
          coordScratch.row = startRow + r;
          coordScratch.col = startCol + c;
          const computed = this.getCellComputedValue(coordScratch);
          sampleValues[r]![c] = formatDisplay(computed);
        }
      }
      summary.textContent = `(range too large: ${formatter.format(totalCells)} cells)`;
    } else {
      let sum = 0;
      let numericCount = 0;
      for (let row = startRow; row <= endRow; row += 1) {
        const sampleRowIdx = row - startRow;
        const inSampleRow = sampleRowIdx >= 0 && sampleRowIdx < sampleRows;
        for (let col = startCol; col <= endCol; col += 1) {
          coordScratch.row = row;
          coordScratch.col = col;
          const computed = this.getCellComputedValue(coordScratch);

          if (inSampleRow) {
            const sampleColIdx = col - startCol;
            if (sampleColIdx >= 0 && sampleColIdx < sampleCols) {
              sampleValues[sampleRowIdx]![sampleColIdx] = formatDisplay(computed);
            }
          }

          // Match Excel/status-bar semantics: only treat actual numbers as numeric values.
          if (typeof computed === "number" && Number.isFinite(computed)) {
            sum += computed;
            numericCount += 1;
          }
        }
      }

      if (numericCount === 0) {
        summary.textContent = "No numeric values";
      } else {
        const formatter =
          this.selectionStatsFormatter ??
          (this.selectionStatsFormatter = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 }));
        summary.textContent = `Sum: ${formatter.format(sum)} · Count: ${formatter.format(numericCount)}`;
      }
    }

    for (let r = 0; r < sampleRows; r += 1) {
      const tr = document.createElement("tr");
      for (let c = 0; c < sampleCols; c += 1) {
        const td = document.createElement("td");
        td.textContent = sampleValues[r]![c] ?? "";
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }

    tooltip.appendChild(table);
    tooltip.appendChild(summary);

    this.formulaRangePreviewTooltipVisible = true;
    tooltip.hidden = false;
    tooltip.setAttribute("aria-hidden", "false");
    tooltip.classList.add("formula-range-preview-tooltip--visible");
    this.syncFormulaRangePreviewTooltipDescribedBy(true);
  }

  private clearSharedHoverCellCache(): void {
    this.sharedHoverCellKey = null;
    this.sharedHoverCellRect = null;
    this.sharedHoverCellHasComment = false;
    this.sharedHoverCellCommentIndexVersion = -1;
  }

  private maybeRefreshRootPosition(opts: { force?: boolean } = {}): void {
    const root = this.root as unknown as HTMLElement | undefined;
    if (!root) return;
    if (typeof root.getBoundingClientRect !== "function") return;

    const now =
      typeof performance !== "undefined" && typeof performance.now === "function" ? performance.now() : Date.now();
    const last = this.rootPosLastMeasuredAtMs;
    const force = opts.force ?? false;
    // Root position changes (e.g. window scroll / layout shifts) are rare compared to pointermove.
    // Refresh at most ~once per second to avoid reintroducing per-move layout reads.
    if (!force && now - last < 1_000) return;

    const rect = root.getBoundingClientRect();
    this.rootLeft = rect.left;
    this.rootTop = rect.top;
    this.rootPosLastMeasuredAtMs = now;
  }

  private commentCellRefFromA1(sheetId: string, a1: string): string {
    return this.collabMode ? `${sheetId}!${a1}` : a1;
  }

  private commentCellRef(cell: CellCoord, sheetId: string = this.sheetId): string {
    return this.commentCellRefFromA1(sheetId, cellToA1(cell));
  }

  private commentCellLabelFromA1(sheetId: string, a1: string): string {
    if (!this.collabMode) return a1;
    const sheetName = this.resolveSheetDisplayNameById(sheetId);
    const prefix = formatSheetNameForA1(sheetName || sheetId);
    return prefix ? `${prefix}!${a1}` : a1;
  }

  private commentCellLabel(cell: CellCoord, sheetId: string = this.sheetId): string {
    return this.commentCellLabelFromA1(sheetId, cellToA1(cell));
  }

  private ensureCommentsUndoScope(root?: ReturnType<typeof getCommentsRoot> | null, opts: { allowCreateBeforeSync?: boolean } = {}): void {
    if (this.commentsUndoScopeAdded) return;
    const undoService = this.collabUndoService;
    if (!undoService) return;
    const localOrigins = undoService.localOrigins;
    if (!localOrigins) return;

    let undoManager: Y.UndoManager | null = null;
    const isYUndoManager = (value: unknown): value is Y.UndoManager => {
      if (value instanceof Y.UndoManager) return true;
      if (!value || typeof value !== "object") return false;
      const maybe = value as any;
      // Bundlers can rename constructors and pnpm workspaces can load multiple `yjs`
      // module instances (ESM + CJS). Avoid relying on `constructor.name`; prefer a
      // structural check instead.
      return (
        typeof maybe.addToScope === "function" &&
        typeof maybe.undo === "function" &&
        typeof maybe.redo === "function" &&
        typeof maybe.stopCapturing === "function"
      );
    };
    for (const origin of localOrigins) {
      if (isYUndoManager(origin)) {
        undoManager = origin as Y.UndoManager;
        break;
      }
    }
    if (!undoManager) return;

    let resolvedRoot = root ?? null;
    if (!resolvedRoot) {
      const session = this.collabSession;
      if (!session) return;

      // Avoid instantiating the `comments` root pre-hydration by default; older
      // documents may still use an Array-backed schema and would be clobbered by
      // `doc.getMap`.
      const provider = session.provider;
      const providerSynced =
        provider && typeof (provider as any).on === "function" ? Boolean((provider as any).synced) : true;
      const hasRoot = Boolean(session.doc.share.get("comments"));
      if (!providerSynced && !hasRoot && !opts.allowCreateBeforeSync) return;

      try {
        resolvedRoot = getCommentsRoot(session.doc);
      } catch {
        return;
      }
    }

    try {
      undoManager.addToScope(resolvedRoot.kind === "map" ? resolvedRoot.map : resolvedRoot.array);
      this.commentsUndoScopeAdded = true;
    } catch {
      // Best-effort; never block comment usage on undo wiring.
    }
  }

  private maybeRefreshCommentsUiForLocalEdit(): void {
    // In collab mode we normally observe the comments root and refresh the UI from
    // observer callbacks. If the user edits comments before the provider reports
    // `sync=true`, the observer may not be attached yet; refresh eagerly so the
    // change is visible immediately.
    if (this.disposed) return;
    if (!this.collabSession) return;
    if (this.stopCommentsRootObserver) return;
    this.reindexCommentCells();
    this.renderCommentsPanel();
    this.refresh();
    this.dispatchCommentsChanged();
  }

  private reindexCommentCells(): void {
    // Comment updates can happen while the user is hovering. Clear any shared-grid
    // hover caches so the next pointermove re-evaluates comment presence and preview
    // content (otherwise we may "stick" to stale hover state).
    this.clearSharedHoverCellCache();
    this.hideCommentTooltip();

    this.commentMetaByCoord.clear();
    this.commentPreviewByCoord.clear();
    this.commentThreadsByCellRef.clear();

    // Collab-mode safety: do not eagerly instantiate the `comments` root before the collab
    // provider has hydrated the shared Y.Doc. Older documents may still use a legacy
    // Array-backed schema, and calling `doc.getMap("comments")` too early can clobber it.
    //
    // `CommentManager.listAll()` calls `getCommentsRoot()` under the hood, so we guard here.
    if (this.collabMode) {
      const session = this.collabSession;
      const provider = session?.provider;
      const providerSynced =
        provider && typeof (provider as any).on === "function" ? Boolean((provider as any).synced) : true;
      const hasRoot = Boolean(session?.doc?.share?.get?.("comments"));
      if (!providerSynced && !hasRoot) {
        this.commentIndexVersion += 1;
        this.sharedProvider?.invalidateAll();
        return;
      }
    }

    // Back-compat: older collaboration docs stored comments under unqualified A1 refs
    // (e.g. "A1"). In collab mode we now require sheet-qualified refs ("sheetId!A1").
    //
    // We cannot infer the original sheet for legacy comments once multiple sheets exist, so
    // we conservatively attach them to the default sheet id to keep them visible without
    // reintroducing cross-sheet collisions.
    const legacyUnqualifiedSheetId = (() => {
      if (!this.collabMode) return null;
      // Some unit tests call `reindexCommentCells` on `Object.create(SpreadsheetApp.prototype)`
      // without running the constructor/field initializers, so `this.document` may be undefined.
      const ids = this.document?.getSheetIds?.() ?? [];
      if (ids.includes("Sheet1")) return "Sheet1";
      // When mixing `??` and `||`, parenthesize explicitly (required by JS syntax).
      return ids[0] ?? (this.sheetId || "Sheet1");
    })();

    for (const comment of this.commentManager.listAll()) {
      const rawCellRef = comment.cellRef;
      // Use the last `!` so sheet-qualified refs with `!` in the sheet id (unlikely but possible)
      // still parse correctly.
      const bang = rawCellRef.lastIndexOf("!");
      const sheetIdFromRef = bang >= 0 ? rawCellRef.slice(0, bang) : null;
      const a1Raw = bang >= 0 ? rawCellRef.slice(bang + 1) : rawCellRef;
      const a1IsPlain = A1_CELL_REF_RE.test(a1Raw);
      const normalizedA1 = a1IsPlain ? a1Raw.replaceAll("$", "").toUpperCase() : a1Raw;
      const sheetId = sheetIdFromRef ?? (a1IsPlain ? legacyUnqualifiedSheetId : null);

      // Normalize A1 refs so `$A$1`, `a1`, etc map to the same cell key.
      // For collab-mode sheet-qualified refs, normalize only the A1 portion.
      const cellRef = sheetId ? `${sheetId}!${normalizedA1}` : a1IsPlain ? normalizedA1 : rawCellRef;
      const resolved = Boolean(comment.resolved);

      const threads = this.commentThreadsByCellRef.get(cellRef);
      if (threads) threads.push(comment);
      else this.commentThreadsByCellRef.set(cellRef, [comment]);

      // In collab mode we key comments by `sheetId!A1`. `commentMetaByCoord` and
      // `commentPreviewByCoord` are used for fast per-cell lookups (indicators/tooltips) and
      // must not collide across sheets, so we only index comments for the active sheet.
      if (this.collabMode) {
        if (!sheetId) continue;
        if (sheetId !== this.sheetId) continue;
      } else {
        // In non-collab mode comments are unqualified ("A1") for back-compat; ignore any
        // qualified refs to keep the UI consistent.
        if (sheetId) continue;
      }

      // Only populate coord-keyed maps when the stored ref looks like a plain A1 address.
      // This prevents corrupt/non-canonical refs from being mis-indexed into A1 (0,0).
      if (!a1IsPlain) continue;

      const coord = parseA1(a1Raw);
      const coordKey = coord.row * COMMENT_COORD_COL_STRIDE + coord.col;
      const existingCoord = this.commentMetaByCoord.get(coordKey);
      if (!existingCoord) this.commentMetaByCoord.set(coordKey, { resolved });
      else existingCoord.resolved = existingCoord.resolved && resolved;

      // Tooltip previews show the first (oldest) thread content for a cell, matching the
      // previous behavior of taking `commentManager.listForCell(...)[0]`.
      if (!this.commentPreviewByCoord.has(coordKey)) {
        this.commentPreviewByCoord.set(coordKey, comment.content ?? "");
      }
    }

    this.commentIndexVersion += 1;

    // The shared renderer caches cell metadata, so comment indicator updates require a provider invalidation.
    this.sharedProvider?.invalidateAll();
  }

  private renderCommentsPanel(): void {
    if (!this.commentsPanelVisible) return;

    this.syncCommentsPanelPermissions();

    const cellRef = this.commentCellRef(this.selection.active);
    const cellLabel = this.commentCellLabel(this.selection.active);
    this.commentsPanelCell.textContent = tWithVars("comments.cellLabel", { cellRef: cellLabel });

    const threads = this.commentThreadsByCellRef.get(cellRef) ?? [];
    this.commentsPanelThreads.replaceChildren();

    if (threads.length === 0) {
      const empty = document.createElement("div");
      empty.textContent = t("comments.none");
      empty.className = "comments-panel__empty";
      this.commentsPanelThreads.appendChild(empty);
      return;
    }

    for (const comment of threads) {
      this.commentsPanelThreads.appendChild(this.renderCommentThread(comment));
    }
  }

  private renderCommentThread(comment: Comment): HTMLElement {
    const canComment = this.canUserComment();
    const container = document.createElement("div");
    container.dataset.testid = "comment-thread";
    container.dataset.commentId = comment.id;
    container.dataset.resolved = comment.resolved ? "true" : "false";
    container.className = "comment-thread";

    const header = document.createElement("div");
    header.className = "comment-thread__header";

    const author = document.createElement("div");
    author.textContent = comment.author.name || t("presence.anonymous");
    author.className = "comment-thread__author";

    const resolve = document.createElement("button");
    resolve.dataset.testid = "resolve-comment";
    resolve.textContent = comment.resolved ? t("comments.unresolve") : t("comments.resolve");
    resolve.type = "button";
    resolve.className = "comment-thread__resolve-button";
    resolve.disabled = !canComment;
    resolve.addEventListener("click", () => {
      if (!this.canUserComment()) return;
      this.commentManager.setResolved({
        commentId: comment.id,
        resolved: !comment.resolved,
      });
      this.maybeRefreshCommentsUiForLocalEdit();
    });

    header.appendChild(author);
    header.appendChild(resolve);
    container.appendChild(header);

    const body = document.createElement("div");
    body.textContent = comment.content;
    body.className = "comment-thread__body";
    container.appendChild(body);

    for (const reply of comment.replies) {
      const replyEl = document.createElement("div");
      replyEl.className = "comment-thread__reply";

      const replyAuthor = document.createElement("div");
      replyAuthor.textContent = reply.author.name || t("presence.anonymous");
      replyAuthor.className = "comment-thread__reply-author";

      const replyBody = document.createElement("div");
      replyBody.textContent = reply.content;
      replyBody.className = "comment-thread__reply-body";

      replyEl.appendChild(replyAuthor);
      replyEl.appendChild(replyBody);
      container.appendChild(replyEl);
    }

    const replyRow = document.createElement("div");
    replyRow.className = "comment-thread__reply-row";

    const replyInput = document.createElement("input");
    replyInput.dataset.testid = "reply-input";
    replyInput.type = "text";
    replyInput.placeholder = t("comments.reply.placeholder");
    replyInput.className = "comment-thread__reply-input";
    replyInput.disabled = !canComment;

    const submitReply = document.createElement("button");
    submitReply.dataset.testid = "submit-reply";
    submitReply.textContent = t("comments.reply.send");
    submitReply.type = "button";
    submitReply.className = "comment-thread__submit-reply-button";
    submitReply.disabled = !canComment;
    submitReply.addEventListener("click", () => {
      if (!this.canUserComment()) return;
      const content = replyInput.value.trim();
      if (!content) return;
      this.commentManager.addReply({
        commentId: comment.id,
        content,
        author: this.currentUser,
      });
      replyInput.value = "";
      this.maybeRefreshCommentsUiForLocalEdit();
    });

    replyRow.appendChild(replyInput);
    replyRow.appendChild(submitReply);
    container.appendChild(replyRow);

    return container;
  }

  private submitNewComment(): void {
    if (!this.canUserComment()) return;
    const content = this.newCommentInput.value.trim();
    if (!content) return;
    const cellRef = this.commentCellRef(this.selection.active);

    this.commentManager.addComment({
      cellRef,
      kind: "threaded",
      content,
      author: this.currentUser,
    });

    this.newCommentInput.value = "";
    this.maybeRefreshCommentsUiForLocalEdit();
  }

  private canUserComment(): boolean {
    return this.collabMode ? (this.collabSession?.canComment() ?? true) : true;
  }

  private syncCommentsPanelPermissions(): void {
    // Some unit tests call comment helpers on `Object.create(SpreadsheetApp.prototype)`
    // without running the constructor.
    if (!this.newCommentInput || !this.newCommentSubmitButton || !this.commentsPanelReadOnlyHint) return;
    const canComment = this.canUserComment();
    this.newCommentInput.disabled = !canComment;
    this.newCommentSubmitButton.disabled = !canComment;
    this.commentsPanelReadOnlyHint.hidden = canComment;
  }

  private onResize(): void {
    const rect = this.root.getBoundingClientRect();
    this.width = rect.width;
    this.height = rect.height;
    this.rootLeft = rect.left;
    this.rootTop = rect.top;
    this.rootPosLastMeasuredAtMs =
      typeof performance !== "undefined" && typeof performance.now === "function" ? performance.now() : Date.now();
    this.clearSharedHoverCellCache();
    this.dpr = window.devicePixelRatio || 1;

    if (this.sharedGrid) {
      // The shared grid owns the main canvas layers, but we still render auditing overlays
      // and chart overlays on separate canvases.
      this.auditingCanvas.width = Math.floor(this.width * this.dpr);
      this.auditingCanvas.height = Math.floor(this.height * this.dpr);
      this.auditingCanvas.style.width = `${this.width}px`;
      this.auditingCanvas.style.height = `${this.height}px`;
      this.auditingCtx.setTransform(1, 0, 0, 1, 0, 0);
      this.auditingCtx.scale(this.dpr, this.dpr);

      this.chartCanvas.width = Math.floor(this.width * this.dpr);
      this.chartCanvas.height = Math.floor(this.height * this.dpr);
      this.chartCanvas.style.width = `${this.width}px`;
      this.chartCanvas.style.height = `${this.height}px`;
      this.chartCtx.setTransform(1, 0, 0, 1, 0, 0);
      this.chartCtx.scale(this.dpr, this.dpr);

      // If auditing is currently off, avoid repeatedly clearing this canvas on scroll.
      // Track that it was resized so `renderAuditing()` can do a one-time clear if needed.
      this.auditingNeedsClear = true;

      this.sharedGrid.resize(this.width, this.height, this.dpr);
      const viewport = this.sharedGrid.renderer.scroll.getViewportState();

      // Keep our legacy scroll coordinates in sync for chart positioning helpers.
      const scroll = this.sharedGrid.getScroll();
      this.scrollX = scroll.x;
      this.scrollY = scroll.y;

      this.renderDrawings(viewport);
      if (!this.useCanvasCharts) {
        this.renderCharts(false);
      }
      this.renderAuditing();
      this.renderSelection();
      this.updateStatus();
      return;
    }

    const legacyCanvases: HTMLCanvasElement[] = [
      this.gridCanvas,
      this.chartCanvas,
      this.referenceCanvas,
      this.auditingCanvas,
      ...(this.presenceCanvas ? [this.presenceCanvas] : []),
      this.selectionCanvas,
    ];
    for (const canvas of legacyCanvases) {
      canvas.width = Math.floor(this.width * this.dpr);
      canvas.height = Math.floor(this.height * this.dpr);
      canvas.style.width = `${this.width}px`;
      canvas.style.height = `${this.height}px`;
    }

    // Reset transforms and apply DPR scaling so drawing code uses CSS pixels.
    const legacyContexts: CanvasRenderingContext2D[] = [
      this.gridCtx,
      this.chartCtx,
      this.referenceCtx,
      this.auditingCtx,
      ...(this.presenceCtx ? [this.presenceCtx] : []),
      this.selectionCtx,
    ];
    for (const ctx of legacyContexts) {
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.scale(this.dpr, this.dpr);
    }
    this.auditingNeedsClear = true;

    const didClamp = this.clampScroll();
    if (didClamp) this.hideCommentTooltip();
    this.syncScrollbars();
    if (didClamp) this.notifyScrollListeners();

    this.renderDrawings();
    this.renderGrid();
    if (!this.useCanvasCharts) {
      this.renderCharts(false);
    }
    this.renderReferencePreview();
    this.renderAuditing();
    this.renderPresence();
    this.renderSelection();
    this.updateStatus();
  }

  private renderGrid(): void {
    if (this.sharedGrid) {
      // Shared-grid mode renders via CanvasGridRenderer. The legacy renderer must
      // not clear or paint over the canvases.
      return;
    }

    this.ensureViewportMappingCurrent();

    const ctx = this.gridCtx;
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, this.gridCanvas.width, this.gridCanvas.height);
    ctx.restore();

    ctx.save();
    ctx.fillStyle = resolveCssVar("--formula-grid-bg", { fallback: resolveCssVar("--bg-primary", { fallback: "Canvas" }) });
    ctx.fillRect(0, 0, this.width, this.height);

    const originX = this.rowHeaderWidth;
    const originY = this.colHeaderHeight;
    const viewportWidth = this.viewportWidth();
    const viewportHeight = this.viewportHeight();

    const frozenWidth = Math.min(viewportWidth, this.frozenWidth);
    const frozenHeight = Math.min(viewportHeight, this.frozenHeight);
    const scrollableWidth = Math.max(0, viewportWidth - frozenWidth);
    const scrollableHeight = Math.max(0, viewportHeight - frozenHeight);

    const frozenRows = this.frozenVisibleRows;
    const frozenCols = this.frozenVisibleCols;
    const scrollRows = this.visibleRows;
    const scrollCols = this.visibleCols;

    const startXScroll = originX + this.visibleColStart * this.cellWidth - this.scrollX;
    const startYScroll = originY + this.visibleRowStart * this.cellHeight - this.scrollY;

    const backgroundBitmap = this.activeSheetBackgroundBitmap;
    if (backgroundBitmap) {
      const pattern = ctx.createPattern(backgroundBitmap, "repeat");
      if (pattern) {
        const fillPattern = (options: {
          clipX: number;
          clipY: number;
          clipWidth: number;
          clipHeight: number;
          translateX: number;
          translateY: number;
        }) => {
          const { clipX, clipY, clipWidth, clipHeight, translateX, translateY } = options;
          if (clipWidth <= 0 || clipHeight <= 0) return;
          ctx.save();
          ctx.beginPath();
          ctx.rect(clipX, clipY, clipWidth, clipHeight);
          ctx.clip();
          ctx.translate(translateX, translateY);
          ctx.fillStyle = pattern;
          ctx.fillRect(clipX - translateX, clipY - translateY, clipWidth, clipHeight);
          ctx.restore();
        };

        // Render beneath the cell area (not underneath row/col headers). For frozen panes, draw per quadrant
        // so pinned regions don't scroll their background pattern.
        fillPattern({
          clipX: originX,
          clipY: originY,
          clipWidth: frozenWidth,
          clipHeight: frozenHeight,
          translateX: originX,
          translateY: originY
        });
        fillPattern({
          clipX: originX + frozenWidth,
          clipY: originY,
          clipWidth: scrollableWidth,
          clipHeight: frozenHeight,
          translateX: originX - this.scrollX,
          translateY: originY
        });
        fillPattern({
          clipX: originX,
          clipY: originY + frozenHeight,
          clipWidth: frozenWidth,
          clipHeight: scrollableHeight,
          translateX: originX,
          translateY: originY - this.scrollY
        });
        fillPattern({
          clipX: originX + frozenWidth,
          clipY: originY + frozenHeight,
          clipWidth: scrollableWidth,
          clipHeight: scrollableHeight,
          translateX: originX - this.scrollX,
          translateY: originY - this.scrollY
        });
      }
    }

    ctx.strokeStyle = resolveCssVar("--formula-grid-line", { fallback: resolveCssVar("--grid-line", { fallback: "CanvasText" }) });
    ctx.lineWidth = 1;

    // Header backgrounds.
    ctx.fillStyle = resolveCssVar("--formula-grid-header-bg", { fallback: resolveCssVar("--grid-header-bg", { fallback: "Canvas" }) });
    ctx.fillRect(0, 0, this.width, this.colHeaderHeight);
    ctx.fillRect(0, 0, this.rowHeaderWidth, this.height);

    // Corner cell.
    ctx.fillStyle = resolveCssVar("--formula-grid-header-bg", { fallback: resolveCssVar("--grid-header-bg", { fallback: "Canvas" }) });
    ctx.fillRect(0, 0, this.rowHeaderWidth, this.colHeaderHeight);

    // Header separator lines.
    ctx.beginPath();
    ctx.moveTo(originX + 0.5, 0);
    ctx.lineTo(originX + 0.5, this.height);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(0, originY + 0.5);
    ctx.lineTo(this.width, originY + 0.5);
    ctx.stroke();

    const cellFontFamily = resolveCssVar("--font-mono", { fallback: DEFAULT_GRID_MONOSPACE_FONT_FAMILY });
    const headerFontFamily = resolveCssVar("--font-sans", { fallback: DEFAULT_GRID_FONT_FAMILY });
    const fontSizePx = 14;
    const defaultTextColor = resolveCssVar("--formula-grid-cell-text", { fallback: resolveCssVar("--text-primary", { fallback: "CanvasText" }) });
    const errorTextColor = resolveCssVar("--formula-grid-error-text", { fallback: resolveCssVar("--error", { fallback: defaultTextColor }) });
    const linkTextColor = resolveCssVar("--formula-grid-link", { fallback: resolveCssVar("--link", { fallback: defaultTextColor }) });
    const commentIndicatorColor = resolveCssVar("--formula-grid-comment-indicator", { fallback: resolveCssVar("--warning", { fallback: "CanvasText" }) });
    const commentIndicatorResolvedColor = resolveCssVar("--formula-grid-comment-indicator-resolved", {
      fallback: resolveCssVar("--text-secondary", { fallback: commentIndicatorColor }),
    });

    // Avoid allocating per-cell `{row,col}` objects in the legacy renderer.
    const coordScratch = { row: 0, col: 0 };

    const renderCellRegion = (options: {
      clipX: number;
      clipY: number;
      clipWidth: number;
      clipHeight: number;
      rows: number[];
      cols: number[];
      startX: number;
      startY: number;
    }): void => {
      const { clipX, clipY, clipWidth, clipHeight, rows, cols, startX, startY } = options;
      if (clipWidth <= 0 || clipHeight <= 0) return;
      if (rows.length === 0 || cols.length === 0) return;

      const endX = startX + cols.length * this.cellWidth;
      const endY = startY + rows.length * this.cellHeight;

      ctx.save();
      ctx.beginPath();
      ctx.rect(clipX, clipY, clipWidth, clipHeight);
      ctx.clip();

      // Grid lines for this quadrant.
      for (let r = 0; r <= rows.length; r++) {
        const y = startY + r * this.cellHeight + 0.5;
        ctx.beginPath();
        ctx.moveTo(startX, y);
        ctx.lineTo(endX, y);
        ctx.stroke();
      }

      for (let c = 0; c <= cols.length; c++) {
        const x = startX + c * this.cellWidth + 0.5;
        ctx.beginPath();
        ctx.moveTo(x, startY);
        ctx.lineTo(x, endY);
        ctx.stroke();
      }

      for (let visualRow = 0; visualRow < rows.length; visualRow++) {
        const row = rows[visualRow]!;
        for (let visualCol = 0; visualCol < cols.length; visualCol++) {
          const col = cols[visualCol]!;
          coordScratch.row = row;
          coordScratch.col = col;
          const state = this.document.getCell(this.sheetId, coordScratch) as {
            value: unknown;
            formula: string | null;
          };
          if (!state) continue;

          let rich: { text: string; runs?: Array<{ start: number; end: number; style?: Record<string, unknown> }> } | null =
            null;
          let color = defaultTextColor;

          if (state.formula != null) {
            if (this.showFormulas) {
              rich = { text: state.formula, runs: [] };
            } else {
              const computed = this.getCellComputedValue(coordScratch);
              if (computed != null) {
                rich = { text: this.formatCellValueForDisplay(coordScratch, computed), runs: [] };
                if (typeof computed === "string" && computed.startsWith("#")) {
                  color = errorTextColor;
                }
              }
            }
          } else if (isRichTextValue(state.value)) {
            rich = state.value;
            } else if (state.value != null) {
              rich = { text: this.formatCellValueForDisplay(coordScratch, state.value as any), runs: [] };
            }

            // Apply default hyperlink styling for URL-like strings when the underlying cell value is not
            // already rich text (rich text runs may carry explicit formatting).
            if (rich && typeof rich.text === "string" && looksLikeExternalHyperlink(rich.text) && !isRichTextValue(state.value)) {
              color = linkTextColor;
              // `renderRichText` draws underlines based on run style.
              if (Array.isArray(rich.runs) && rich.runs.length === 0) {
                rich.runs = [{ start: 0, end: rich.text.length, style: { underline: true } }];
              }
            }

            if (!rich || rich.text === "") continue;

          renderRichText(
            ctx,
            rich,
            {
              x: startX + visualCol * this.cellWidth,
              y: startY + visualRow * this.cellHeight,
              width: this.cellWidth,
              height: this.cellHeight
            },
            {
              padding: 4,
              align: "start",
              verticalAlign: "middle",
              fontFamily: cellFontFamily,
              fontSizePx,
              color
            }
          );
        }
      }

      // Comment indicators.
      for (let visualRow = 0; visualRow < rows.length; visualRow++) {
        const row = rows[visualRow]!;
        for (let visualCol = 0; visualCol < cols.length; visualCol++) {
          const col = cols[visualCol]!;
          const meta = this.commentMetaByCoord.get(row * COMMENT_COORD_COL_STRIDE + col);
          if (!meta) continue;
          const resolved = meta.resolved ?? false;
          drawCommentIndicator(ctx, {
            x: startX + visualCol * this.cellWidth,
            y: startY + visualRow * this.cellHeight,
            width: this.cellWidth,
            height: this.cellHeight,
          }, {
            color: resolved ? commentIndicatorResolvedColor : commentIndicatorColor,
          });
        }
      }

      ctx.restore();
    };

    // --- Data quadrants ---
    renderCellRegion({
      clipX: originX,
      clipY: originY,
      clipWidth: frozenWidth,
      clipHeight: frozenHeight,
      rows: frozenRows,
      cols: frozenCols,
      startX: originX,
      startY: originY
    });

    renderCellRegion({
      clipX: originX + frozenWidth,
      clipY: originY,
      clipWidth: scrollableWidth,
      clipHeight: frozenHeight,
      rows: frozenRows,
      cols: scrollCols,
      startX: startXScroll,
      startY: originY
    });

    renderCellRegion({
      clipX: originX,
      clipY: originY + frozenHeight,
      clipWidth: frozenWidth,
      clipHeight: scrollableHeight,
      rows: scrollRows,
      cols: frozenCols,
      startX: originX,
      startY: startYScroll
    });

    renderCellRegion({
      clipX: originX + frozenWidth,
      clipY: originY + frozenHeight,
      clipWidth: scrollableWidth,
      clipHeight: scrollableHeight,
      rows: scrollRows,
      cols: scrollCols,
      startX: startXScroll,
      startY: startYScroll
    });

    // Freeze lines.
    ctx.save();
    ctx.strokeStyle = resolveCssVar("--formula-grid-freeze-line", {
      fallback: resolveCssVar("--grid-line", { fallback: "CanvasText" })
    });
    ctx.lineWidth = 2;
    if (frozenWidth > 0 && frozenWidth < viewportWidth) {
      const x = originX + frozenWidth + 0.5;
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, this.height);
      ctx.stroke();
    }
    if (frozenHeight > 0 && frozenHeight < viewportHeight) {
      const y = originY + frozenHeight + 0.5;
      ctx.beginPath();
      ctx.moveTo(0, y);
      ctx.lineTo(this.width, y);
      ctx.stroke();
    }
    ctx.restore();

    // Header labels.
    ctx.fillStyle = resolveCssVar("--formula-grid-header-text", { fallback: resolveCssVar("--text-primary", { fallback: "CanvasText" }) });
    ctx.font = `12px ${headerFontFamily}`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    // Column header labels: frozen columns (pinned).
    ctx.save();
    ctx.beginPath();
    ctx.rect(originX, 0, frozenWidth, originY);
    ctx.clip();
    for (let visualCol = 0; visualCol < frozenCols.length; visualCol++) {
      const colIndex = frozenCols[visualCol]!;
      ctx.fillText(colToName(colIndex), originX + visualCol * this.cellWidth + this.cellWidth / 2, originY / 2);
    }
    ctx.restore();

    // Column header labels: scrollable columns.
    ctx.save();
    ctx.beginPath();
    ctx.rect(originX + frozenWidth, 0, scrollableWidth, originY);
    ctx.clip();
    for (let visualCol = 0; visualCol < scrollCols.length; visualCol++) {
      const colIndex = scrollCols[visualCol]!;
      ctx.fillText(colToName(colIndex), startXScroll + visualCol * this.cellWidth + this.cellWidth / 2, originY / 2);
    }
    ctx.restore();

    // Row header labels: frozen rows (pinned).
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, originY, originX, frozenHeight);
    ctx.clip();
    for (let visualRow = 0; visualRow < frozenRows.length; visualRow++) {
      const rowIndex = frozenRows[visualRow]!;
      ctx.fillText(String(rowIndex + 1), originX / 2, originY + visualRow * this.cellHeight + this.cellHeight / 2);
    }
    ctx.restore();

    // Row header labels: scrollable rows.
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, originY + frozenHeight, originX, scrollableHeight);
    ctx.clip();
    for (let visualRow = 0; visualRow < scrollRows.length; visualRow++) {
      const rowIndex = scrollRows[visualRow]!;
      ctx.fillText(
        String(rowIndex + 1),
        originX / 2,
        startYScroll + visualRow * this.cellHeight + this.cellHeight / 2
      );
    }
    ctx.restore();

    ctx.restore();

    this.renderOutlineControls();
  }

  private lowerBound(values: number[], target: number): number {
    let lo = 0;
    let hi = values.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if ((values[mid] ?? 0) < target) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  private visualIndexForRow(row: number): number {
    const direct = this.rowToVisual.get(row);
    if (direct !== undefined) return direct;
    return this.lowerBound(this.rowIndexByVisual, row);
  }

  private visualIndexForCol(col: number): number {
    const direct = this.colToVisual.get(col);
    if (direct !== undefined) return direct;
    return this.lowerBound(this.colIndexByVisual, col);
  }

  private chartAnchorToViewportRect(anchor: ChartRecord["anchor"]): { left: number; top: number; width: number; height: number } | null {
    if (!anchor || !("kind" in anchor)) return null;

    const zoom = this.getZoom();
    const z = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;

    let left = 0;
    let top = 0;
    let width = 0;
    let height = 0;

    if (anchor.kind === "absolute") {
      left = emuToPx(anchor.xEmu) * z;
      top = emuToPx(anchor.yEmu) * z;
      width = emuToPx(anchor.cxEmu) * z;
      height = emuToPx(anchor.cyEmu) * z;
    } else if (anchor.kind === "oneCell") {
      if (this.sharedGrid) {
        const headerRows = this.sharedHeaderRows();
        const headerCols = this.sharedHeaderCols();
        const headerWidth = headerCols > 0 ? this.sharedGrid.renderer.scroll.cols.totalSize(headerCols) : 0;
        const headerHeight = headerRows > 0 ? this.sharedGrid.renderer.scroll.rows.totalSize(headerRows) : 0;
        const gridRow = anchor.fromRow + headerRows;
        const gridCol = anchor.fromCol + headerCols;
        const { rowCount, colCount } = this.sharedGrid.renderer.scroll.getCounts();
        if (gridRow < 0 || gridRow >= rowCount || gridCol < 0 || gridCol >= colCount) return null;

        left =
          this.sharedGrid.renderer.scroll.cols.positionOf(gridCol) -
          headerWidth +
          emuToPx(anchor.fromColOffEmu ?? 0) * z;
        top =
          this.sharedGrid.renderer.scroll.rows.positionOf(gridRow) -
          headerHeight +
          emuToPx(anchor.fromRowOffEmu ?? 0) * z;
        width = emuToPx(anchor.cxEmu ?? 0) * z;
        height = emuToPx(anchor.cyEmu ?? 0) * z;
      } else {
        left = this.visualIndexForCol(anchor.fromCol) * this.cellWidth + emuToPx(anchor.fromColOffEmu) * z;
        top = this.visualIndexForRow(anchor.fromRow) * this.cellHeight + emuToPx(anchor.fromRowOffEmu) * z;
        width = emuToPx(anchor.cxEmu) * z;
        height = emuToPx(anchor.cyEmu) * z;
      }
    } else if (anchor.kind === "twoCell") {
      if (this.sharedGrid) {
        const headerRows = this.sharedHeaderRows();
        const headerCols = this.sharedHeaderCols();
        const headerWidth = headerCols > 0 ? this.sharedGrid.renderer.scroll.cols.totalSize(headerCols) : 0;
        const headerHeight = headerRows > 0 ? this.sharedGrid.renderer.scroll.rows.totalSize(headerRows) : 0;
        const fromRow = anchor.fromRow + headerRows;
        const fromCol = anchor.fromCol + headerCols;
        const toRow = anchor.toRow + headerRows;
        const toCol = anchor.toCol + headerCols;
        const { rowCount, colCount } = this.sharedGrid.renderer.scroll.getCounts();
        if (fromRow < 0 || fromRow >= rowCount || toRow < 0 || toRow >= rowCount) return null;
        if (fromCol < 0 || fromCol >= colCount || toCol < 0 || toCol >= colCount) return null;

        left =
          this.sharedGrid.renderer.scroll.cols.positionOf(fromCol) -
          headerWidth +
          emuToPx(anchor.fromColOffEmu ?? 0) * z;
        top =
          this.sharedGrid.renderer.scroll.rows.positionOf(fromRow) -
          headerHeight +
          emuToPx(anchor.fromRowOffEmu ?? 0) * z;
        const right =
          this.sharedGrid.renderer.scroll.cols.positionOf(toCol) -
          headerWidth +
          emuToPx(anchor.toColOffEmu ?? 0) * z;
        const bottom =
          this.sharedGrid.renderer.scroll.rows.positionOf(toRow) -
          headerHeight +
          emuToPx(anchor.toRowOffEmu ?? 0) * z;

        width = Math.max(0, right - left);
        height = Math.max(0, bottom - top);
      } else {
        left = this.visualIndexForCol(anchor.fromCol) * this.cellWidth + emuToPx(anchor.fromColOffEmu) * z;
        top = this.visualIndexForRow(anchor.fromRow) * this.cellHeight + emuToPx(anchor.fromRowOffEmu) * z;
        const right = this.visualIndexForCol(anchor.toCol) * this.cellWidth + emuToPx(anchor.toColOffEmu) * z;
        const bottom = this.visualIndexForRow(anchor.toRow) * this.cellHeight + emuToPx(anchor.toRowOffEmu) * z;
        width = Math.max(0, right - left);
        height = Math.max(0, bottom - top);
      }
    } else {
      return null;
    }

    if (width <= 0 || height <= 0) return null;

    const { frozenRows, frozenCols } = this.getFrozen();
    const scrollX = anchor.kind === "absolute" ? this.scrollX : (anchor.fromCol < frozenCols ? 0 : this.scrollX);
    const scrollY = anchor.kind === "absolute" ? this.scrollY : (anchor.fromRow < frozenRows ? 0 : this.scrollY);

    return {
      left: left - scrollX,
      top: top - scrollY,
      width,
      height
    };
  }

  private chartCellOriginPx(cell: { row: number; col: number }): { x: number; y: number } {
    if (this.sharedGrid) {
      const headerRows = this.sharedHeaderRows();
      const headerCols = this.sharedHeaderCols();
      const gridRow = cell.row + headerRows;
      const gridCol = cell.col + headerCols;

      const headerWidth = headerCols > 0 ? this.sharedGrid.renderer.scroll.cols.totalSize(headerCols) : 0;
      const headerHeight = headerRows > 0 ? this.sharedGrid.renderer.scroll.rows.totalSize(headerRows) : 0;
      return {
        x: this.sharedGrid.renderer.scroll.cols.positionOf(gridCol) - headerWidth,
        y: this.sharedGrid.renderer.scroll.rows.positionOf(gridRow) - headerHeight,
      };
    }

    return {
      x: this.visualIndexForCol(cell.col) * this.cellWidth,
      y: this.visualIndexForRow(cell.row) * this.cellHeight,
    };
  }

  private chartCellSizePx(cell: { row: number; col: number }): { width: number; height: number } {
    if (this.sharedGrid) {
      const headerRows = this.sharedHeaderRows();
      const headerCols = this.sharedHeaderCols();
      const gridRow = cell.row + headerRows;
      const gridCol = cell.col + headerCols;
      return {
        width: this.sharedGrid.renderer.getColWidth(gridCol),
        height: this.sharedGrid.renderer.getRowHeight(gridRow),
      };
    }

    return { width: this.cellWidth, height: this.cellHeight };
  }

  private setSelectedChartId(id: string | null): void {
    const next = id && String(id).trim() !== "" ? String(id) : null;
    // Selecting a chart should clear any drawing selection so selection handles don't
    // double-render and split-view panes can mirror a single "active object" selection.
    if (next != null && this.selectedDrawingId != null) {
      this.selectedDrawingId = null;
      this.dispatchDrawingSelectionChanged();
      this.renderDrawings();
    }
    if (next === this.selectedChartId) return;
    this.selectedChartId = next;
    this.renderChartSelectionOverlay();
  }

  private chartIdToDrawingId(chartId: string): number {
    const match = /^chart_(\d+)$/.exec(chartId);
    if (match) {
      const parsed = Number(match[1]);
      if (Number.isFinite(parsed) && parsed > 0) return parsed;
    }
    let hash = 0;
    for (let i = 0; i < chartId.length; i += 1) {
      hash = (hash * 31 + chartId.charCodeAt(i)) | 0;
    }
    return Math.max(1, Math.abs(hash));
  }

  private chartOverlayLayout(sharedViewport?: GridViewportState): {
    originX: number;
    originY: number;
    frozenBoundaryX: number;
    frozenBoundaryY: number;
    paneRects: Record<"topLeft" | "topRight" | "bottomLeft" | "bottomRight", { x: number; y: number; width: number; height: number }>;
  } {
    if (this.sharedGrid) {
      const viewport = sharedViewport ?? this.sharedGrid.renderer.scroll.getViewportState();
      const headerRows = this.sharedHeaderRows();
      const headerCols = this.sharedHeaderCols();
      const headerWidth = headerCols > 0 ? this.sharedGrid.renderer.scroll.cols.totalSize(headerCols) : 0;
      const headerHeight = headerRows > 0 ? this.sharedGrid.renderer.scroll.rows.totalSize(headerRows) : 0;
      const headerWidthClamped = Math.min(headerWidth, viewport.width);
      const headerHeightClamped = Math.min(headerHeight, viewport.height);

      const frozenWidthClamped = Math.min(viewport.frozenWidth, viewport.width);
      const frozenHeightClamped = Math.min(viewport.frozenHeight, viewport.height);

      const frozenContentWidth = Math.max(0, frozenWidthClamped - headerWidthClamped);
      const frozenContentHeight = Math.max(0, frozenHeightClamped - headerHeightClamped);

      const cellAreaWidth = Math.max(0, viewport.width - headerWidthClamped);
      const cellAreaHeight = Math.max(0, viewport.height - headerHeightClamped);

      const scrollableWidth = Math.max(0, cellAreaWidth - frozenContentWidth);
      const scrollableHeight = Math.max(0, cellAreaHeight - frozenContentHeight);

      return {
        originX: headerWidthClamped,
        originY: headerHeightClamped,
        frozenBoundaryX: frozenWidthClamped,
        frozenBoundaryY: frozenHeightClamped,
        paneRects: {
          topLeft: { x: 0, y: 0, width: frozenContentWidth, height: frozenContentHeight },
          topRight: { x: frozenContentWidth, y: 0, width: scrollableWidth, height: frozenContentHeight },
          bottomLeft: { x: 0, y: frozenContentHeight, width: frozenContentWidth, height: scrollableHeight },
          bottomRight: { x: frozenContentWidth, y: frozenContentHeight, width: scrollableWidth, height: scrollableHeight },
        },
      };
    }

    const originX = this.rowHeaderWidth;
    const originY = this.colHeaderHeight;
    const cellAreaWidth = Math.max(0, this.width - originX);
    const cellAreaHeight = Math.max(0, this.height - originY);
    const frozenContentWidth = Math.min(cellAreaWidth, this.frozenWidth);
    const frozenContentHeight = Math.min(cellAreaHeight, this.frozenHeight);
    const scrollableWidth = Math.max(0, cellAreaWidth - frozenContentWidth);
    const scrollableHeight = Math.max(0, cellAreaHeight - frozenContentHeight);
    return {
      originX,
      originY,
      frozenBoundaryX: originX + frozenContentWidth,
      frozenBoundaryY: originY + frozenContentHeight,
      paneRects: {
        topLeft: { x: 0, y: 0, width: frozenContentWidth, height: frozenContentHeight },
        topRight: { x: frozenContentWidth, y: 0, width: scrollableWidth, height: frozenContentHeight },
        bottomLeft: { x: 0, y: frozenContentHeight, width: frozenContentWidth, height: scrollableHeight },
        bottomRight: { x: frozenContentWidth, y: frozenContentHeight, width: scrollableWidth, height: scrollableHeight },
      },
    };
  }

  private chartPointPxToAnchorPoint(point: { x: number; y: number }): {
    col: number;
    row: number;
    colOffEmu: number;
    rowOffEmu: number;
  } {
    const zoom = this.getZoom();
    const z = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;

    const x = Math.max(0, point.x);
    const y = Math.max(0, point.y);

    if (this.sharedGrid) {
      const renderer = this.sharedGrid.renderer;
      const cols = renderer.scroll.cols;
      const rows = renderer.scroll.rows;
      const headerRows = this.sharedHeaderRows();
      const headerCols = this.sharedHeaderCols();
      const headerWidth = headerCols > 0 ? cols.totalSize(headerCols) : 0;
      const headerHeight = headerRows > 0 ? rows.totalSize(headerRows) : 0;
      const counts = renderer.scroll.getCounts();
      const maxGridCol = Math.max(0, counts.colCount - 1);
      const maxGridRow = Math.max(0, counts.rowCount - 1);

      const gridCol = cols.indexAt(x + headerWidth, { min: headerCols, maxInclusive: maxGridCol });
      const gridRow = rows.indexAt(y + headerHeight, { min: headerRows, maxInclusive: maxGridRow });

      const col = Math.max(0, gridCol - headerCols);
      const row = Math.max(0, gridRow - headerRows);

      const originX = cols.positionOf(gridCol) - headerWidth;
      const originY = rows.positionOf(gridRow) - headerHeight;
      return {
        col,
        row,
        colOffEmu: Math.round(pxToEmu((x - originX) / z)),
        rowOffEmu: Math.round(pxToEmu((y - originY) / z)),
      };
    }

    const colVisual = Math.floor(x / this.cellWidth);
    const rowVisual = Math.floor(y / this.cellHeight);

    const hasCols = this.colIndexByVisual.length > 0;
    const hasRows = this.rowIndexByVisual.length > 0;

    const safeColVisual = hasCols ? Math.max(0, Math.min(this.colIndexByVisual.length - 1, colVisual)) : Math.max(0, colVisual);
    const safeRowVisual = hasRows ? Math.max(0, Math.min(this.rowIndexByVisual.length - 1, rowVisual)) : Math.max(0, rowVisual);

    const col = (hasCols ? this.colIndexByVisual[safeColVisual] : safeColVisual) ?? safeColVisual;
    const row = (hasRows ? this.rowIndexByVisual[safeRowVisual] : safeRowVisual) ?? safeRowVisual;

    const originX = safeColVisual * this.cellWidth;
    const originY = safeRowVisual * this.cellHeight;
    return {
      col,
      row,
      colOffEmu: Math.round(pxToEmu((x - originX) / z)),
      rowOffEmu: Math.round(pxToEmu((y - originY) / z)),
    };
  }

  private computeChartAnchorFromRectPx(
    anchorKind: ChartRecord["anchor"]["kind"],
    rect: { x: number; y: number; width: number; height: number },
  ): ChartRecord["anchor"] {
    const zoom = this.getZoom();
    const z = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;

    const x = Math.max(0, rect.x);
    const y = Math.max(0, rect.y);
    const width = Math.max(1, rect.width);
    const height = Math.max(1, rect.height);

    if (anchorKind === "absolute") {
      return {
        kind: "absolute",
        xEmu: Math.round(pxToEmu(x / z)),
        yEmu: Math.round(pxToEmu(y / z)),
        cxEmu: Math.round(pxToEmu(width / z)),
        cyEmu: Math.round(pxToEmu(height / z)),
      };
    }

    if (anchorKind === "oneCell") {
      const from = this.chartPointPxToAnchorPoint({ x, y });
      return {
        kind: "oneCell",
        fromCol: from.col,
        fromRow: from.row,
        fromColOffEmu: from.colOffEmu,
        fromRowOffEmu: from.rowOffEmu,
        cxEmu: Math.round(pxToEmu(width / z)),
        cyEmu: Math.round(pxToEmu(height / z)),
      };
    }

    // twoCell
    const from = this.chartPointPxToAnchorPoint({ x, y });
    const to = this.chartPointPxToAnchorPoint({ x: x + width, y: y + height });
    return {
      kind: "twoCell",
      fromCol: from.col,
      fromRow: from.row,
      fromColOffEmu: from.colOffEmu,
      fromRowOffEmu: from.rowOffEmu,
      toCol: to.col,
      toRow: to.row,
      toColOffEmu: to.colOffEmu,
      toRowOffEmu: to.rowOffEmu,
    };
  }

  private renderChartSelectionOverlay(): void {
    const geom = this.chartOverlayGeom;
    const overlay = this.chartSelectionOverlay;
    if (!geom || !overlay) return;

    const charts = this.chartStore.listCharts().filter((chart) => chart.sheetId === this.sheetId);
    const selected = this.selectedChartId ? charts.find((chart) => chart.id === this.selectedChartId) : null;

    const { frozenRows, frozenCols } = this.getFrozen();
    const layout = this.chartOverlayLayout();
    const viewport: DrawingViewport = {
      scrollX: this.scrollX,
      scrollY: this.scrollY,
      width: this.width,
      height: this.height,
      dpr: this.dpr,
      zoom: this.getZoom(),
      frozenRows,
      frozenCols,
      frozenWidthPx: layout.frozenBoundaryX,
      frozenHeightPx: layout.frozenBoundaryY,
      headerOffsetX: layout.originX,
      headerOffsetY: layout.originY,
    };
    const memo = this.chartSelectionViewportMemo;
    if (!memo || memo.width !== viewport.width || memo.height !== viewport.height || memo.dpr !== viewport.dpr) {
      overlay.resize(viewport);
      this.chartSelectionViewportMemo = { width: viewport.width, height: viewport.height, dpr: viewport.dpr };
    }

    if (!selected) {
      overlay.setSelectedId(null);
      void overlay.render([], viewport, { drawObjects: false }).catch((err) => {
        console.warn("Chart selection overlay render failed", err);
      });
      return;
    }

    const drawingId = this.chartIdToDrawingId(selected.id);
    const obj: DrawingObject = {
      id: drawingId,
      kind: { type: "chart", chartId: selected.id },
      anchor: chartAnchorToDrawingAnchor(selected.anchor),
      zOrder: 0,
    };
    overlay.setSelectedId(drawingId);
    void overlay.render([obj], viewport, { drawObjects: false }).catch((err) => {
      console.warn("Chart selection overlay render failed", err);
    });
  }

  private hitTestChartAtClientPoint(clientX: number, clientY: number): {
    chart: ChartRecord;
    rect: { left: number; top: number; width: number; height: number };
    pane: { key: "topLeft" | "topRight" | "bottomLeft" | "bottomRight"; rect: { x: number; y: number; width: number; height: number } };
    pointInCellArea: { x: number; y: number };
  } | null {
    this.maybeRefreshRootPosition({ force: true });
    const layout = this.chartOverlayLayout(this.sharedGrid ? this.sharedGrid.renderer.scroll.getViewportState() : undefined);
    const x = clientX - this.rootLeft - layout.originX;
    const y = clientY - this.rootTop - layout.originY;
    if (!Number.isFinite(x) || !Number.isFinite(y)) return null;

    const charts = this.chartStore.listCharts().filter((chart) => chart.sheetId === this.sheetId);
    if (charts.length === 0) return null;

    const intersect = (
      a: { left: number; top: number; width: number; height: number },
      b: { left: number; top: number; width: number; height: number },
    ): { left: number; top: number; width: number; height: number } | null => {
      const left = Math.max(a.left, b.left);
      const top = Math.max(a.top, b.top);
      const right = Math.min(a.left + a.width, b.left + b.width);
      const bottom = Math.min(a.top + a.height, b.top + b.height);
      const width = right - left;
      const height = bottom - top;
      if (width <= 0 || height <= 0) return null;
      return { left, top, width, height };
    };

    const { frozenRows, frozenCols } = this.getFrozen();
    for (let i = charts.length - 1; i >= 0; i -= 1) {
      const chart = charts[i]!;
      const rect = this.chartAnchorToViewportRect(chart.anchor);
      if (!rect) continue;
      const fromRow = chart.anchor.kind === "oneCell" || chart.anchor.kind === "twoCell" ? chart.anchor.fromRow : Number.POSITIVE_INFINITY;
      const fromCol = chart.anchor.kind === "oneCell" || chart.anchor.kind === "twoCell" ? chart.anchor.fromCol : Number.POSITIVE_INFINITY;
      const inFrozenRows = fromRow < frozenRows;
      const inFrozenCols = fromCol < frozenCols;
      const paneKey: "topLeft" | "topRight" | "bottomLeft" | "bottomRight" =
        inFrozenRows && inFrozenCols
          ? "topLeft"
          : inFrozenRows && !inFrozenCols
            ? "topRight"
            : !inFrozenRows && inFrozenCols
              ? "bottomLeft"
              : "bottomRight";
      const paneRect = layout.paneRects[paneKey];

      const visible = intersect(rect, { left: paneRect.x, top: paneRect.y, width: paneRect.width, height: paneRect.height });
      if (!visible) continue;
      if (x < visible.left || x > visible.left + visible.width) continue;
      if (y < visible.top || y > visible.top + visible.height) continue;

      return {
        chart,
        rect,
        pane: { key: paneKey, rect: paneRect },
        pointInCellArea: { x, y },
      };
    }

    return null;
  }

  private chartResizeHandleAtPoint(hit: {
    rect: { left: number; top: number; width: number; height: number };
    pointInCellArea: { x: number; y: number };
  }): ResizeHandle | null {
    return hitTestResizeHandle(
      { x: hit.rect.left, y: hit.rect.top, width: hit.rect.width, height: hit.rect.height },
      hit.pointInCellArea.x,
      hit.pointInCellArea.y,
    );
  }

  private onChartPointerDownCapture(e: PointerEvent): void {
    if (this.disposed) return;
    if (e.button !== 0) return;

    const target = e.target as HTMLElement | null;
    // Only treat pointerdown events originating from the grid surface (canvases/root) as
    // chart selection/drags. This avoids interfering with interactive DOM overlays
    // (scrollbars, outline buttons, comments panel, etc) even when a chart happens to
    // extend underneath them.
    const isGridSurface =
      target === this.root ||
      target === this.selectionCanvas ||
      target === this.gridCanvas ||
      target === this.referenceCanvas ||
      target === this.auditingCanvas ||
      target === this.presenceCanvas;
    if (!isGridSurface) return;

    const hit = this.hitTestChartAtClientPoint(e.clientX, e.clientY);
    if (!hit) {
      if (this.selectedChartId != null) this.setSelectedChartId(null);
      return;
    }

    e.preventDefault();
    e.stopPropagation();

    // Chart interactions take precedence over drawing selection. Since we stop propagation
    // here, any previously-selected drawing would otherwise remain selected indefinitely.
    if (this.selectedDrawingId != null) {
      this.selectedDrawingId = null;
      this.drawingOverlay.setSelectedId(null);
      this.renderDrawings();
    }

    const wasSelected = this.selectedChartId === hit.chart.id;
    this.setSelectedChartId(hit.chart.id);
    this.focus();

    const resizeHandle = wasSelected ? this.chartResizeHandleAtPoint(hit) : null;
    const mode = resizeHandle ? "resize" : "move";

    this.chartDragAbort?.abort();
    this.chartDragAbort = new AbortController();

    this.chartDragState = {
      pointerId: e.pointerId,
      chartId: hit.chart.id,
      mode,
      ...(resizeHandle ? { resizeHandle } : {}),
      startClientX: e.clientX,
      startClientY: e.clientY,
      startAnchor: { ...(hit.chart.anchor as any) },
    };

    const signal = this.chartDragAbort.signal;
    const onMove = (ev: PointerEvent) => this.onChartDragPointerMove(ev);
    const onUp = (ev: PointerEvent) => this.onChartDragPointerUp(ev);
    window.addEventListener("pointermove", onMove, { capture: true, passive: false, signal });
    window.addEventListener("pointerup", onUp, { capture: true, passive: false, signal });
    window.addEventListener("pointercancel", onUp, { capture: true, passive: false, signal });
  }

  private onDrawingPointerDownCapture(e: PointerEvent): void {
    if (this.disposed) return;
    if (e.button !== 0) return;
    // When the dedicated DrawingInteractionController is enabled, it owns selection/dragging.
    // Avoid competing with its pointer listeners (especially in legacy mode where it uses
    // bubbling listeners and relies on pointer events not being cancelled in capture phase).
    if (this.drawingInteractionController) return;
    // If another capture listener already claimed the event (e.g. chart interactions),
    // do not compete.
    if (e.cancelBubble) return;
    // When the formula bar is in range-selection mode, drawing hits should not steal the
    // pointerdown; let normal grid range selection continue.
    if (this.formulaBar?.isFormulaEditing()) return;
    const target = e.target as HTMLElement | null;
    // Only treat pointerdown events originating from the grid surface (canvases/root) as
    // drawing selection. This avoids interfering with interactive DOM overlays
    // (scrollbars, outline buttons, comments panel, etc) even when drawings extend underneath them.
    const isGridSurface =
      target === this.root ||
      target === this.selectionCanvas ||
      target === this.gridCanvas ||
      target === this.referenceCanvas ||
      target === this.auditingCanvas ||
      target === this.presenceCanvas;
    if (!isGridSurface) return;
    const objects = this.listDrawingObjectsForSheet();
    const prevSelected = this.selectedDrawingId;
    const editorWasOpen = this.editor.isOpen();

    // If there are no drawings, clear selection on any click in the cell area.
    if (objects.length === 0) {
      if (prevSelected != null) {
        this.selectedDrawingId = null;
        this.dispatchDrawingSelectionChanged();
        this.renderDrawings(this.sharedGrid ? this.sharedGrid.renderer.scroll.getViewportState() : undefined);
      }
      return;
    }

    this.maybeRefreshRootPosition({ force: true });
    const x = e.clientX - this.rootLeft;
    const y = e.clientY - this.rootTop;
    if (!Number.isFinite(x) || !Number.isFinite(y)) return;

    const sharedViewport = this.sharedGrid ? this.sharedGrid.renderer.scroll.getViewportState() : undefined;
    const viewport = this.getDrawingInteractionViewport(sharedViewport);
    const index = this.getDrawingHitTestIndex(objects);
    const hit = hitTestDrawings(index, viewport, x, y, this.drawingGeom);

    if (!hit) {
      // Clicking outside of any drawing clears selection, but still allows normal grid selection.
      if (prevSelected != null) {
        const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
        const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;
        if (x >= headerOffsetX && y >= headerOffsetY) {
          this.selectedDrawingId = null;
          this.dispatchDrawingSelectionChanged();
          this.renderDrawings(sharedViewport);
        }
      }
      return;
    }

    if (editorWasOpen) {
      this.editor.commit("command");
    }

    // In shared-grid mode, pointer events typically target the full-size selection canvas and
    // would otherwise initiate grid selection. Selecting a drawing should behave like Excel
    // (the grid selection should not change), so stop propagation in that mode.
    //
    // In legacy mode, allow the bubbling `onPointerDown` handler to run so it can start the
    // existing drag/resize gesture state machine for drawings.
    if (this.sharedGrid) {
      e.preventDefault();
      e.stopPropagation();
    }

    // Drawings and charts should behave like a single selection model; selecting a drawing
    // clears any chart selection so selection handles don't double-render.
    if (this.selectedChartId != null) {
      this.setSelectedChartId(null);
    }

    this.selectedDrawingId = hit.object.id;
    if (this.selectedDrawingId !== prevSelected) {
      this.dispatchDrawingSelectionChanged();
      this.renderDrawings(sharedViewport);
    }
    this.focus();
  }

  private onChartDragPointerMove(e: PointerEvent): void {
    const state = this.chartDragState;
    if (!state) return;
    if (e.pointerId !== state.pointerId) return;

    e.preventDefault();
    e.stopPropagation();

    const dx = e.clientX - state.startClientX;
    const dy = e.clientY - state.startClientY;

    const startRect = this.chartAnchorToViewportRect(state.startAnchor);
    if (!startRect) return;

    const minSize = 20;
    const nextRect = (() => {
      if (state.mode === "move") {
        return {
          x: startRect.left + dx,
          y: startRect.top + dy,
          width: startRect.width,
          height: startRect.height
        };
      }

      const handle = state.resizeHandle ?? "se";
      let x = startRect.left;
      let y = startRect.top;
      let width = startRect.width;
      let height = startRect.height;

      const fromWest = handle === "w" || handle === "nw" || handle === "sw";
      const fromEast = handle === "e" || handle === "ne" || handle === "se";
      const fromNorth = handle === "n" || handle === "ne" || handle === "nw";
      const fromSouth = handle === "s" || handle === "sw" || handle === "se";

      if (fromEast) {
        width = width + dx;
      }
      if (fromWest) {
        x = x + dx;
        width = width - dx;
      }

      if (fromSouth) {
        height = height + dy;
      }
      if (fromNorth) {
        y = y + dy;
        height = height - dy;
      }

      if (width < minSize) {
        const delta = minSize - width;
        width = minSize;
        if (fromWest) {
          x -= delta;
        }
      }
      if (height < minSize) {
        const delta = minSize - height;
        height = minSize;
        if (fromNorth) {
          y -= delta;
        }
      }

      return { x, y, width, height };
    })();

    const { frozenRows, frozenCols } = this.getFrozen();

    const computeAnchor = (scrollBaseX: number, scrollBaseY: number): ChartRecord["anchor"] =>
      this.computeChartAnchorFromRectPx(state.startAnchor.kind, {
        x: nextRect.x + scrollBaseX,
        y: nextRect.y + scrollBaseY,
        width: nextRect.width,
        height: nextRect.height,
      });

    // Convert from screen-space (viewport/cell-area) back to sheet-space coordinates for anchor updates.
    //
    // The effective scroll offsets are determined by the *resulting* anchor cell quadrant, not by the
    // raw screen-space X/Y values (a chart anchored in the scrollable pane can have a negative `left`
    // when it extends offscreen).
    let scrollBaseX =
      state.startAnchor.kind === "absolute"
        ? this.scrollX
        : state.startAnchor.fromCol < frozenCols
          ? 0
          : this.scrollX;
    let scrollBaseY =
      state.startAnchor.kind === "absolute"
        ? this.scrollY
        : state.startAnchor.fromRow < frozenRows
          ? 0
          : this.scrollY;

    let nextAnchor = computeAnchor(scrollBaseX, scrollBaseY);

    if (nextAnchor.kind !== "absolute") {
      const desiredScrollX = nextAnchor.fromCol < frozenCols ? 0 : this.scrollX;
      const desiredScrollY = nextAnchor.fromRow < frozenRows ? 0 : this.scrollY;
      if (desiredScrollX !== scrollBaseX || desiredScrollY !== scrollBaseY) {
        scrollBaseX = desiredScrollX;
        scrollBaseY = desiredScrollY;
        nextAnchor = computeAnchor(scrollBaseX, scrollBaseY);
      }
    }
    this.chartStore.updateChartAnchor(state.chartId, nextAnchor);
  }

  private onChartDragPointerUp(e: PointerEvent): void {
    const state = this.chartDragState;
    if (!state) return;
    if (e.pointerId !== state.pointerId) return;

    e.preventDefault();
    e.stopPropagation();

    this.chartDragState = null;
    this.chartDragAbort?.abort();
    this.chartDragAbort = null;
  }
  private renderCharts(renderContent: boolean): void {
    if (this.useCanvasCharts) {
      if (renderContent) {
        // Full chart refresh (e.g. sheet switch): invalidate every chart so the adapter rebuilds
        // cached models on the next draw.
        this.invalidateCanvasChartsForActiveSheet();
        this.dirtyChartIds.clear();
      } else if (this.dirtyChartIds.size > 0) {
        // Targeted refresh: invalidate only charts whose underlying data ranges were touched.
        for (const id of this.dirtyChartIds) {
          this.chartCanvasStoreAdapter.invalidate(id);
        }
        this.dirtyChartIds.clear();
      }
      // Charts render as drawing objects on the drawings overlay canvas.
      this.renderDrawings();
      return;
    }

    const charts = this.chartStore.listCharts().filter((chart) => chart.sheetId === this.sheetId);
    const keep = new Set<string>();

    if (!this.sharedGrid) {
      // Keep frozen pane geometry current for quadrant clipping.
      this.ensureViewportMappingCurrent();
    }

    // Reset the chart canvas transform and clear in CSS-pixel coordinates.
    const ctx = this.chartCtx;
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.scale(this.dpr, this.dpr);
    ctx.clearRect(0, 0, this.width, this.height);

    const layout = (() => {
      if (this.sharedGrid) {
        const viewport = this.sharedGrid.renderer.scroll.getViewportState();
        const headerRows = this.sharedHeaderRows();
        const headerCols = this.sharedHeaderCols();
        const headerWidth = headerCols > 0 ? this.sharedGrid.renderer.scroll.cols.totalSize(headerCols) : 0;
        const headerHeight = headerRows > 0 ? this.sharedGrid.renderer.scroll.rows.totalSize(headerRows) : 0;

        const headerWidthClamped = Math.min(headerWidth, viewport.width);
        const headerHeightClamped = Math.min(headerHeight, viewport.height);

        const frozenWidthClamped = Math.min(viewport.frozenWidth, viewport.width);
        const frozenHeightClamped = Math.min(viewport.frozenHeight, viewport.height);

        const frozenContentWidth = Math.max(0, frozenWidthClamped - headerWidthClamped);
        const frozenContentHeight = Math.max(0, frozenHeightClamped - headerHeightClamped);

        const cellAreaWidth = Math.max(0, viewport.width - headerWidthClamped);
        const cellAreaHeight = Math.max(0, viewport.height - headerHeightClamped);

        const scrollableWidth = Math.max(0, cellAreaWidth - frozenContentWidth);
        const scrollableHeight = Math.max(0, cellAreaHeight - frozenContentHeight);

        return {
          originX: headerWidthClamped,
          originY: headerHeightClamped,
          frozenContentWidth,
          frozenContentHeight,
          scrollableWidth,
          scrollableHeight,
          cellAreaWidth,
          cellAreaHeight,
        };
      }

      const cellAreaWidth = Math.max(0, this.width - this.rowHeaderWidth);
      const cellAreaHeight = Math.max(0, this.height - this.colHeaderHeight);
      const frozenWidth = Math.min(cellAreaWidth, this.frozenWidth);
      const frozenHeight = Math.min(cellAreaHeight, this.frozenHeight);
      const scrollableWidth = Math.max(0, cellAreaWidth - frozenWidth);
      const scrollableHeight = Math.max(0, cellAreaHeight - frozenHeight);
      return {
        originX: this.rowHeaderWidth,
        originY: this.colHeaderHeight,
        frozenContentWidth: frozenWidth,
        frozenContentHeight: frozenHeight,
        scrollableWidth,
        scrollableHeight,
        cellAreaWidth,
        cellAreaHeight,
      };
    })();

    const paneRects = {
      topLeft: { x: 0, y: 0, width: layout.frozenContentWidth, height: layout.frozenContentHeight },
      topRight: {
        x: layout.frozenContentWidth,
        y: 0,
        width: layout.scrollableWidth,
        height: layout.frozenContentHeight,
      },
      bottomLeft: {
        x: 0,
        y: layout.frozenContentHeight,
        width: layout.frozenContentWidth,
        height: layout.scrollableHeight,
      },
      bottomRight: {
        x: layout.frozenContentWidth,
        y: layout.frozenContentHeight,
        width: layout.scrollableWidth,
        height: layout.scrollableHeight,
      },
    };

    const { frozenRows, frozenCols } = this.getFrozen();

    const sheetCount = (this.document as any)?.model?.sheets?.size;
    const useEngineCache = (typeof sheetCount === "number" ? sheetCount : this.document.getSheetIds().length) <= 1;
    const hasWasmEngine = Boolean(this.wasmEngine && !this.wasmSyncSuspended);
    const trackFormulaCells = !hasWasmEngine || !useEngineCache;
    if (!trackFormulaCells && this.chartHasFormulaCells.size > 0) {
      // Avoid retaining potentially-stale `false` entries when the engine availability changes.
      // If we later need to fall back to recalc-driven invalidation, missing entries are treated
      // conservatively (we mark charts dirty).
      this.chartHasFormulaCells.clear();
    }

    const createProvider = () => {
      // Avoid allocating a fresh `{row,col}` object for every chart range cell read.
      const coordScratch = { row: 0, col: 0 };
      const memo = new Map<string, Map<number, SpreadsheetValue>>();
      const stack = new Map<string, Set<number>>();
      return {
        getRange: (rangeRef: string, flags?: { sawFormula: boolean }) => {
          const parsed = parseA1Range(rangeRef);
          if (!parsed) return [];
          const sheetId = parsed.sheetName ? this.resolveSheetIdByName(parsed.sheetName) : this.sheetId;
          if (!sheetId) return [];

          const out: unknown[][] = [];
          for (let r = parsed.startRow; r <= parsed.endRow; r += 1) {
            const row: unknown[] = [];
            coordScratch.row = r;
            for (let c = parsed.startCol; c <= parsed.endCol; c += 1) {
              coordScratch.col = c;
              row.push(this.computeCellValue(sheetId, coordScratch, memo, stack, { useEngineCache }, flags));
            }
            out.push(row);
          }
          return out;
        },
      };
    };

    const flatten = (range2d: unknown[][]): unknown[] => {
      if (!Array.isArray(range2d)) return [];
      const out: unknown[] = [];
      for (const row of range2d) {
        if (!Array.isArray(row)) continue;
        for (const value of row) out.push(value);
      }
      return out;
    };

    const toCategoryCache = (values: unknown[]): Array<string | number | null> =>
      values.map((value) => {
        if (value == null) return null;
        if (typeof value === "string" || typeof value === "number") return value;
        if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
        return String(value);
      });

    const toNumberCache = (values: unknown[]): Array<number | string | null> =>
      values.map((value) => {
        if (value == null) return null;
        if (typeof value === "number" || typeof value === "string") return value;
        if (typeof value === "boolean") return value ? 1 : 0;
        return String(value);
      });

    const normalizeChartKind = (kind: ChartRecord["chartType"]["kind"]): ChartModel["chartType"]["kind"] => {
      if (kind === "bar" || kind === "line" || kind === "pie" || kind === "scatter") return kind;
      if (kind === "area") return "line";
      return "unknown";
    };

    const buildBaseModel = (chart: ChartRecord): ChartModel => {
      const kind = normalizeChartKind(chart.chartType.kind);
      const hasNamedSeries = (chart.series ?? []).some((ser) => typeof ser.name === "string" && ser.name.trim() !== "");
      const wantsLegend = kind === "pie" || (chart.series?.length ?? 0) > 1 || hasNamedSeries;

      return {
        chartType: { kind, ...(chart.chartType.name ? { name: chart.chartType.name } : {}) },
        title: chart.title ?? null,
        legend: wantsLegend ? { position: "right", overlay: false } : { position: "none", overlay: false },
        ...(kind === "scatter"
          ? {
              axes: [
                { kind: "value", position: "bottom", formatCode: "0" },
                { kind: "value", position: "left", majorGridlines: true, formatCode: "0" },
              ],
            }
          : kind === "pie"
            ? {}
            : {
                axes: [
                  { kind: "category", position: "bottom" },
                  { kind: "value", position: "left", majorGridlines: true, formatCode: "0" },
                ],
              }),
        series: (chart.series ?? []).map((ser) => ({
          ...(ser.name != null ? { name: ser.name } : {}),
        })),
      };
    };

    const chartDataTooLarge = (chart: ChartRecord): boolean => {
      for (const ser of chart.series ?? []) {
        const refs = [ser.categories, ser.values, ser.xValues, ser.yValues];
        for (const rangeRef of refs) {
          if (typeof rangeRef !== "string" || rangeRef.trim() === "") continue;
          const parsed = parseA1Range(rangeRef);
          if (!parsed) continue;
          const rows = Math.max(0, parsed.endRow - parsed.startRow + 1);
          const cols = Math.max(0, parsed.endCol - parsed.startCol + 1);
          if (rows * cols > MAX_CHART_DATA_CELLS) return true;
        }
      }
      return false;
    };

    const intersects = (a: { x: number; y: number; width: number; height: number }, b: { x: number; y: number; width: number; height: number }): boolean => {
      return !(
        a.x + a.width < b.x ||
        b.x + b.width < a.x ||
        a.y + a.height < b.y ||
        b.y + b.height < a.y
      );
    };

    let provider: ReturnType<typeof createProvider> | null = null;

    ctx.save();
    // Chart anchors are computed in cell-area coordinates; translate under headers once.
    ctx.translate(layout.originX, layout.originY);

    for (const chart of charts) {
      keep.add(chart.id);
      const rect = this.chartAnchorToViewportRect(chart.anchor);
      if (!rect) continue;

      const chartRect = { x: rect.left, y: rect.top, width: rect.width, height: rect.height };

      const fromRow = chart.anchor.kind === "oneCell" || chart.anchor.kind === "twoCell" ? chart.anchor.fromRow : Number.POSITIVE_INFINITY;
      const fromCol = chart.anchor.kind === "oneCell" || chart.anchor.kind === "twoCell" ? chart.anchor.fromCol : Number.POSITIVE_INFINITY;
      const inFrozenRows = fromRow < frozenRows;
      const inFrozenCols = fromCol < frozenCols;

      const pane = inFrozenRows
        ? inFrozenCols
          ? paneRects.topLeft
          : paneRects.topRight
        : inFrozenCols
          ? paneRects.bottomLeft
          : paneRects.bottomRight;

      if (pane.width <= 0 || pane.height <= 0) continue;
      if (!intersects(chartRect, pane)) continue;

      const isDirty = this.dirtyChartIds.has(chart.id);
      const shouldUpdateModel = renderContent || isDirty || !this.chartModels.has(chart.id);
      if (shouldUpdateModel) {
        const base = buildBaseModel(chart);
        if (chartDataTooLarge(chart)) {
          base.options = {
            ...(base.options ?? {}),
            placeholder: `Chart range too large (>${MAX_CHART_DATA_CELLS.toLocaleString()} cells)`,
          };
          this.chartModels.set(chart.id, base);
          this.chartHasFormulaCells.set(chart.id, false);
          this.dirtyChartIds.delete(chart.id);
        } else {
          if (!provider) provider = createProvider();
          const rangeFlags = trackFormulaCells ? { sawFormula: false } : undefined;

          const nextSeries = base.series.map((ser, idx) => {
            const def = chart.series[idx];
            const categories = def.categories ? toCategoryCache(flatten(provider!.getRange(def.categories, rangeFlags))) : [];
            const values = def.values ? toNumberCache(flatten(provider!.getRange(def.values, rangeFlags))) : [];
            const xValues = def.xValues ? toNumberCache(flatten(provider!.getRange(def.xValues, rangeFlags))) : [];
            const yValues = def.yValues ? toNumberCache(flatten(provider!.getRange(def.yValues, rangeFlags))) : [];

            const pieFallback =
              base.chartType.kind === "pie" && categories.length === 0 && values.length > 0
                ? Array.from({ length: values.length }, (_, i) => String(i + 1))
                : null;

            return {
              ...ser,
              ...(pieFallback
                ? { categories: { cache: pieFallback } }
                : categories.length
                  ? { categories: { cache: categories } }
                  : {}),
              ...(values.length ? { values: { cache: values } } : {}),
              ...(xValues.length ? { xValues: { cache: xValues } } : {}),
              ...(yValues.length ? { yValues: { cache: yValues } } : {}),
            };
          });

          this.chartModels.set(chart.id, { ...base, series: nextSeries });
          if (trackFormulaCells) {
            this.chartHasFormulaCells.set(chart.id, Boolean(rangeFlags?.sawFormula));
          } else {
            this.chartHasFormulaCells.delete(chart.id);
          }
          this.dirtyChartIds.delete(chart.id);
        }
      }

      ctx.save();
      ctx.beginPath();
      ctx.rect(pane.x, pane.y, pane.width, pane.height);
      ctx.clip();
      try {
        this.chartRenderer.renderToCanvas(ctx, chart.id, chartRect);
      } catch {
        // Best-effort: ignore chart rendering failures so one bad chart doesn't block overlays.
      }
      ctx.restore();
    }

    ctx.restore();

    // Drop cached offscreen surfaces for charts that no longer exist on the active sheet.
    this.chartRenderer.pruneSurfaces(keep);

    for (const id of this.chartModels.keys()) {
      if (keep.has(id)) continue;
      this.chartModels.delete(id);
    }
    for (const id of this.dirtyChartIds) {
      if (keep.has(id)) continue;
      this.dirtyChartIds.delete(id);
    }
    for (const id of this.chartHasFormulaCells.keys()) {
      if (keep.has(id)) continue;
      this.chartHasFormulaCells.delete(id);
    }
    for (const id of this.chartRangeRectsCache.keys()) {
      if (keep.has(id)) continue;
      this.chartRangeRectsCache.delete(id);
    }

    if (this.selectedChartId != null && !keep.has(this.selectedChartId)) {
      this.selectedChartId = null;
    }
    this.renderChartSelectionOverlay();
  }

  private syncDrawingOverlayViewport(sharedViewport?: GridViewportState): DrawingViewport {
    return this.getDrawingRenderViewport(sharedViewport);
  }

  private computeDrawingViewportLayout(
    sharedViewport?: GridViewportState,
  ): {
    headerOffsetX: number;
    headerOffsetY: number;
    rootWidth: number;
    rootHeight: number;
    cellAreaWidth: number;
    cellAreaHeight: number;
    /** Frozen boundary position in selection/root coordinates (includes header offsets). */
    frozenBoundaryXRoot: number;
    frozenBoundaryYRoot: number;
    /** Frozen boundary position in cell-area coordinates (excludes header offsets). */
    frozenBoundaryXCellArea: number;
    frozenBoundaryYCellArea: number;
    frozenRows: number;
    frozenCols: number;
  } {
    const { frozenRows, frozenCols } = this.getFrozen();

    const clamp = (value: number, min: number, max: number): number => Math.min(max, Math.max(min, value));

    if (this.sharedGrid) {
      const viewport = sharedViewport ?? this.sharedGrid.renderer.scroll.getViewportState();
      const headerRows = this.sharedHeaderRows();
      const headerCols = this.sharedHeaderCols();
      const headerWidth = headerCols > 0 ? this.sharedGrid.renderer.scroll.cols.totalSize(headerCols) : 0;
      const headerHeight = headerRows > 0 ? this.sharedGrid.renderer.scroll.rows.totalSize(headerRows) : 0;

      const headerOffsetX = Math.min(headerWidth, viewport.width);
      const headerOffsetY = Math.min(headerHeight, viewport.height);
      const rootWidth = viewport.width;
      const rootHeight = viewport.height;
      const cellAreaWidth = Math.max(0, rootWidth - headerOffsetX);
      const cellAreaHeight = Math.max(0, rootHeight - headerOffsetY);

      // Shared-grid viewport frozen extents include the synthetic header row/col, but
      // drawings viewports expect sheet-level frozen row/col counts. Keep the pixel
      // boundary positions from the renderer (so hidden/variable row/col sizes stay
      // aligned) while passing sheet-level frozenRows/frozenCols counts separately.
      const frozenBoundaryXRoot = clamp(viewport.frozenWidth, headerOffsetX, rootWidth);
      const frozenBoundaryYRoot = clamp(viewport.frozenHeight, headerOffsetY, rootHeight);
      const frozenBoundaryXCellArea = clamp(frozenBoundaryXRoot - headerOffsetX, 0, cellAreaWidth);
      const frozenBoundaryYCellArea = clamp(frozenBoundaryYRoot - headerOffsetY, 0, cellAreaHeight);

      return {
        headerOffsetX,
        headerOffsetY,
        rootWidth,
        rootHeight,
        cellAreaWidth,
        cellAreaHeight,
        frozenBoundaryXRoot,
        frozenBoundaryYRoot,
        frozenBoundaryXCellArea,
        frozenBoundaryYCellArea,
        frozenRows,
        frozenCols,
      };
    }

    // Legacy renderer: frozen pane extents are derived from visible (non-hidden) rows/cols.
    this.ensureViewportMappingCurrent();

    const headerOffsetX = this.rowHeaderWidth;
    const headerOffsetY = this.colHeaderHeight;
    const rootWidth = this.width;
    const rootHeight = this.height;
    const cellAreaWidth = Math.max(0, rootWidth - headerOffsetX);
    const cellAreaHeight = Math.max(0, rootHeight - headerOffsetY);

    const frozenBoundaryXCellArea = clamp(this.frozenWidth, 0, cellAreaWidth);
    const frozenBoundaryYCellArea = clamp(this.frozenHeight, 0, cellAreaHeight);
    const frozenBoundaryXRoot = clamp(headerOffsetX + frozenBoundaryXCellArea, headerOffsetX, rootWidth);
    const frozenBoundaryYRoot = clamp(headerOffsetY + frozenBoundaryYCellArea, headerOffsetY, rootHeight);

    return {
      headerOffsetX,
      headerOffsetY,
      rootWidth,
      rootHeight,
      cellAreaWidth,
      cellAreaHeight,
      frozenBoundaryXRoot,
      frozenBoundaryYRoot,
      frozenBoundaryXCellArea,
      frozenBoundaryYCellArea,
      frozenRows,
      frozenCols,
    };
  }

  /**
   * Viewport used for rendering on `drawingCanvas` (full `.grid-root` coordinates).
   *
   * The drawing overlay canvas spans the entire grid root so it can share the
   * same coordinate space as other overlays (selection, outline, etc). To ensure
   * drawings never paint over the row/col headers, we pass `headerOffsetX/Y` so
   * the DrawingOverlay can clip to the cell grid body area.
   */
  getDrawingRenderViewport(sharedViewport?: GridViewportState): DrawingViewport {
    const layout = this.computeDrawingViewportLayout(sharedViewport);

    // Reset any legacy positioning so the drawing canvas always covers the full grid root.
    this.drawingCanvas.style.left = "0px";
    this.drawingCanvas.style.top = "0px";

    const scrollX = this.sharedGrid && sharedViewport ? sharedViewport.scrollX : this.scrollX;
    const scrollY = this.sharedGrid && sharedViewport ? sharedViewport.scrollY : this.scrollY;

    const viewport: DrawingViewport = {
      scrollX,
      scrollY,
      width: layout.rootWidth,
      height: layout.rootHeight,
      dpr: this.dpr,
      zoom: this.getZoom(),
      headerOffsetX: layout.headerOffsetX,
      headerOffsetY: layout.headerOffsetY,
      frozenRows: layout.frozenRows,
      frozenCols: layout.frozenCols,
      frozenWidthPx: layout.frozenBoundaryXRoot,
      frozenHeightPx: layout.frozenBoundaryYRoot,
    };

    const memo = this.drawingViewportMemo;
    if (!memo || memo.width !== viewport.width || memo.height !== viewport.height || memo.dpr !== viewport.dpr) {
      this.drawingOverlay.resize(viewport);
      this.drawingViewportMemo = {
        width: viewport.width,
        height: viewport.height,
        dpr: viewport.dpr,
      };
    }

    return viewport;
  }

  /**
   * Viewport used for hit testing + interactions on surfaces that include the row/col
   * headers (e.g. `selectionCanvas` in shared-grid mode).
   */
  getDrawingInteractionViewport(sharedViewport?: GridViewportState): DrawingViewport {
    const layout = this.computeDrawingViewportLayout(sharedViewport);
    const scrollX = this.sharedGrid && sharedViewport ? sharedViewport.scrollX : this.scrollX;
    const scrollY = this.sharedGrid && sharedViewport ? sharedViewport.scrollY : this.scrollY;
    return {
      scrollX,
      scrollY,
      width: layout.rootWidth,
      height: layout.rootHeight,
      dpr: this.dpr,
      zoom: this.getZoom(),
      headerOffsetX: layout.headerOffsetX,
      headerOffsetY: layout.headerOffsetY,
      frozenRows: layout.frozenRows,
      frozenCols: layout.frozenCols,
      frozenWidthPx: layout.frozenBoundaryXRoot,
      frozenHeightPx: layout.frozenBoundaryYRoot,
    };
  }

  private documentChangeAffectsDrawings(payload: any): boolean {
    if (!payload || typeof payload !== "object") return false;

    const source = typeof payload?.source === "string" ? payload.source : "";
    // Applying a new document snapshot can replace the drawing layer entirely.
    if (source === "applyState") return true;
    // Some integrations may publish drawings/images updates with a dedicated source tag.
    if (source === "drawings" || source === "images") return true;

    const sheetId = this.sheetId;
    const matchesSheet = (delta: any): boolean => String(delta?.sheetId ?? "") === sheetId;
    const touchesSheet = (deltas: any): boolean => Array.isArray(deltas) && deltas.some(matchesSheet);

    // Sheet view changes (frozen panes, row/col sizes, and/or drawing metadata stored on the view)
    // can affect how drawings are rendered even if the drawings list itself did not change.
    if (touchesSheet(payload?.sheetViewDeltas)) return true;

    // Sheet meta/order changes can change the active sheet or invalidate cached geometry.
    if (touchesSheet(payload?.sheetMetaDeltas) || payload?.sheetOrderDelta) return true;

    // Drawing deltas are usually per-sheet.
    if (
      touchesSheet(payload?.drawingsDeltas) ||
      touchesSheet(payload?.drawingDeltas) ||
      touchesSheet(payload?.sheetDrawingsDeltas) ||
      touchesSheet(payload?.sheetDrawingDeltas)
    ) {
      return true;
    }

    // Image updates may be workbook-wide; re-render so any referenced bitmaps refresh.
    if (Array.isArray(payload?.imagesDeltas) || Array.isArray(payload?.imageDeltas)) return true;

    if (payload?.drawingsChanged === true || payload?.imagesChanged === true) return true;

    return false;
  }

  private listDrawingObjectsForSheet(sheetId: string = this.sheetId): DrawingObject[] {
    const doc = this.document as any;
    const drawingsGetter = typeof doc.getSheetDrawings === "function" ? doc.getSheetDrawings : null;

    const cached = this.drawingObjectsCache;
    // `getSheetDrawings` is not part of DocumentController's stable public API yet, so some
    // tests (and older builds) monkeypatch it in after SpreadsheetApp construction. Include the
    // getter function identity in the cache key so we don't permanently cache an empty list
    // from the pre-monkeypatch state.
    let docObjects: DrawingObject[];
    if (cached && cached.sheetId === sheetId && cached.source === drawingsGetter) {
      docObjects = cached.objects;
    } else {
      let raw: unknown = null;
      if (drawingsGetter) {
        try {
          raw = drawingsGetter.call(doc, sheetId);
        } catch {
          raw = null;
        }
      }

      const isUiDrawingObject = (value: unknown): value is DrawingObject => {
        if (!value || typeof value !== "object") return false;
        const anyValue = value as any;
        return (
          typeof anyValue.id === "number" &&
          anyValue.kind &&
          typeof anyValue.kind.type === "string" &&
          anyValue.anchor &&
          typeof anyValue.anchor.type === "string"
        );
      };

      const objects: DrawingObject[] = (() => {
        if (raw == null) return [];
        // Already normalized UI objects.
        if (Array.isArray(raw)) {
          if (raw.length === 0) return [];
          if (raw.every(isUiDrawingObject)) return raw as DrawingObject[];
          // DocumentController drawings (or model objects) as a raw array (best-effort).
          return convertDocumentSheetDrawingsToUiDrawingObjects(raw, { sheetId });
        }

        if (isUiDrawingObject(raw)) return [raw];

        // Formula-model worksheet JSON blob ({ drawings: [...] }).
        if (typeof raw === "object") {
          const maybeWorksheet = raw as any;
          if (Array.isArray(maybeWorksheet.drawings)) {
            return convertModelWorksheetDrawingsToUiDrawingObjects(maybeWorksheet);
          }
          if (Array.isArray(maybeWorksheet.objects)) {
            const list = maybeWorksheet.objects as unknown[];
            if (list.every(isUiDrawingObject)) return list as DrawingObject[];
            return convertDocumentSheetDrawingsToUiDrawingObjects(list, { sheetId });
          }
        }

        return [];
      })();

      const ordered = (() => {
        if (objects.length <= 1) return objects;
        for (let i = 1; i < objects.length; i += 1) {
          // Treat missing/invalid zOrder as 0; this keeps the adapter resilient to older callers.
          const prev = typeof (objects[i - 1] as any)?.zOrder === "number" ? (objects[i - 1] as any).zOrder : 0;
          const curr = typeof (objects[i] as any)?.zOrder === "number" ? (objects[i] as any).zOrder : 0;
          if (prev > curr) {
            return [...objects].sort((a, b) => ((a as any).zOrder ?? 0) - ((b as any).zOrder ?? 0));
          }
        }
        return objects;
      })();

      this.drawingObjectsCache = { sheetId, objects: ordered, source: drawingsGetter };
      docObjects = ordered;
    }

    const finalObjects: DrawingObject[] =
      docObjects.length > 0 || !this.drawingsDemoEnabled
        ? docObjects
        : ([
            {
              id: 1,
              kind: { type: "shape", label: "Demo Drawing" },
              anchor: {
                type: "oneCell",
                from: { cell: { row: 0, col: 0 }, offset: { xEmu: pxToEmu(8), yEmu: pxToEmu(8) } },
                size: { cx: pxToEmu(240), cy: pxToEmu(120) },
              },
              zOrder: 0,
            },
          ] as DrawingObject[]);

    if (finalObjects !== docObjects) {
      this.drawingObjectsCache = { sheetId, objects: finalObjects, source: drawingsGetter };
    }
    return finalObjects;
  }

  private setDrawingObjectsForSheet(objects: DrawingObject[]): void {
    const doc = this.document as any;
    const drawingsGetter = typeof doc.getSheetDrawings === "function" ? doc.getSheetDrawings : null;
    this.drawingObjectsCache = { sheetId: this.sheetId, objects, source: drawingsGetter };
  }

  private listCanvasChartDrawingObjectsForSheet(sheetId: string, zBase: number): DrawingObject[] {
    const charts = this.chartStore.listCharts().filter((chart) => chart.sheetId === sheetId);
    if (charts.length === 0) return [];
    return charts.map((chart, idx) => chartRecordToDrawingObject(chart, zBase + idx));
  }

  private invalidateCanvasChartsForActiveSheet(): void {
    const charts = this.chartStore.listCharts();
    for (const chart of charts) {
      if (chart.sheetId !== this.sheetId) continue;
      this.chartCanvasStoreAdapter.invalidate(chart.id);
    }
  }

  private renderDrawings(sharedViewport?: GridViewportState): void {
    // In shared-grid mode the renderer can emit an initial viewport notification while
    // SpreadsheetApp is still constructing (e.g. unit tests that stub
    // `requestAnimationFrame` synchronously). Guard against accessing the overlay before it
    // is initialized.
    const overlay = (this as any).drawingOverlay as DrawingOverlay | undefined;
    if (!overlay) return;
    const viewport = this.getDrawingRenderViewport(sharedViewport);
    const baseObjects = this.listDrawingObjectsForSheet();
    this.drawingObjects = baseObjects;
    if (this.selectedDrawingId != null && !baseObjects.some((o) => o.id === this.selectedDrawingId)) {
      this.selectedDrawingId = null;
    }

    const objects: DrawingObject[] = (() => {
      if (!this.useCanvasCharts) return baseObjects;
      const maxZ = baseObjects.reduce((acc, obj) => Math.max(acc, obj.zOrder), -1);
      const charts = this.listCanvasChartDrawingObjectsForSheet(this.sheetId, maxZ + 1);
      return charts.length > 0 ? [...baseObjects, ...charts] : baseObjects;
    })();

    if (this.useCanvasCharts && this.selectedDrawingId == null && this.selectedChartId != null) {
      const selectedChartDrawingId = chartStoreIdToDrawingId(this.selectedChartId);
      if (!objects.some((o) => o.id === selectedChartDrawingId)) {
        this.selectedChartId = null;
      }
    }

    const selectedOverlayId =
      this.selectedDrawingId != null
        ? this.selectedDrawingId
        : this.useCanvasCharts && this.selectedChartId != null
          ? chartStoreIdToDrawingId(this.selectedChartId)
          : null;

    const keepChartIds = new Set<string>();
    for (const obj of objects) {
      if (obj.kind.type !== "chart") continue;
      const id = typeof obj.kind.chartId === "string" ? obj.kind.chartId.trim() : "";
      if (id) keepChartIds.add(id);
    }
    // Avoid retaining cached offscreen surfaces for charts that are no longer on the active sheet.
    this.drawingChartRenderer.pruneSurfaces(keepChartIds);

    overlay.setSelectedId(selectedOverlayId);
    void overlay.render(objects, viewport).catch((err) => {
      console.warn("Drawing overlay render failed", err);
    });
  }

  private isImageReferencedByAnyDrawing(imageId: string): boolean {
    const id = String(imageId ?? "").trim();
    if (!id) return false;

    const docAny: any = this.document as any;
    const getSheetDrawings = typeof docAny.getSheetDrawings === "function" ? (docAny.getSheetDrawings as (sheetId: string) => unknown) : null;
    if (!getSheetDrawings) return false;

    const sheetIds = (() => {
      try {
        const ids = this.document.getSheetIds?.();
        return Array.isArray(ids) ? ids : [this.sheetId];
      } catch {
        return [this.sheetId];
      }
    })();

    for (const sheetId of sheetIds) {
      let raw: unknown = null;
      try {
        raw = getSheetDrawings.call(docAny, sheetId);
      } catch {
        raw = null;
      }
      if (!Array.isArray(raw) || raw.length === 0) continue;

      const objects = convertDocumentSheetDrawingsToUiDrawingObjects(raw, { sheetId });
      for (const obj of objects) {
        if (obj.kind.type === "image" && obj.kind.imageId === id) return true;
      }
    }

    return false;
  }

  private renderAuditing(): void {
    if (this.auditingMode === "off") {
      // Avoid clearing the full auditing canvas on every scroll when auditing overlays are disabled.
      // We only need to clear when (1) overlays were previously rendered or (2) the canvas was resized/reinitialized.
      if (this.auditingWasRendered || this.auditingNeedsClear) {
        this.auditingRenderer.clear(this.auditingCtx);
        this.auditingWasRendered = false;
        this.auditingNeedsClear = false;
      }
      return;
    }

    this.auditingNeedsClear = false;
    this.auditingRenderer.clear(this.auditingCtx);

    const clipRect = this.sharedGrid
      ? (() => {
          const viewport = this.sharedGrid.renderer.scroll.getViewportState();
          const headerWidth = Math.max(0, this.sharedGrid.renderer.getColWidth(0));
          const headerHeight = Math.max(0, this.sharedGrid.renderer.getRowHeight(0));
          return {
            x: headerWidth,
            y: headerHeight,
            width: Math.max(0, viewport.width - headerWidth),
            height: Math.max(0, viewport.height - headerHeight),
          };
        })()
      : {
          x: this.rowHeaderWidth,
          y: this.colHeaderHeight,
          width: this.viewportWidth(),
          height: this.viewportHeight(),
        };

    this.auditingCtx.save();
    this.auditingCtx.beginPath();
    this.auditingCtx.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height);
    this.auditingCtx.clip();

    this.auditingRenderer.render(this.auditingCtx, this.auditingHighlights, {
      getCellRect: (row, col) => this.getCellRect({ row, col }),
    });

    this.auditingCtx.restore();
    this.auditingWasRendered = true;
  }

  private renderPresence(): void {
    if (this.sharedGrid) return;
    const ctx = this.presenceCtx;
    const renderer = this.presenceRenderer;
    if (!ctx || !renderer) return;

    renderer.clear(ctx);
    if (this.remotePresences.length === 0) return;

    this.ensureViewportMappingCurrent();
    const clipRect = {
      x: this.rowHeaderWidth,
      y: this.colHeaderHeight,
      width: this.viewportWidth(),
      height: this.viewportHeight(),
    };

    ctx.save();
    ctx.beginPath();
    ctx.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height);
    ctx.clip();

    renderer.render(ctx, this.remotePresences, {
      getCellRect: (row: number, col: number) => this.getCellRect({ row, col }),
    });

    ctx.restore();
  }

  private renderSelection(): void {
    if (this.sharedGrid) {
      // Selection rendering is handled by the shared CanvasGridRenderer selection layer.
      // We still need to keep the in-place editor aligned when the viewport scrolls or resizes.
      if (this.editor.isOpen()) {
        const gridCell = this.gridCellFromDocCell(this.selection.active);
        const rect = this.sharedGrid.getCellRect(gridCell.row, gridCell.col);
        if (rect) this.editor.reposition(rect);
      }
      return;
    }

    this.ensureViewportMappingCurrent();
    const clipRect = {
      x: this.rowHeaderWidth,
      y: this.colHeaderHeight,
      width: this.viewportWidth(),
      height: this.viewportHeight(),
    };

    const renderer = this.formulaBar?.isFormulaEditing() ? this.formulaSelectionRenderer : this.selectionRenderer;

    const visibleRows = this.frozenVisibleRows.length
      ? [...this.frozenVisibleRows, ...this.visibleRows]
      : this.visibleRows;
    const visibleCols = this.frozenVisibleCols.length
      ? [...this.frozenVisibleCols, ...this.visibleCols]
      : this.visibleCols;

    renderer.render(
      this.selectionCtx,
      this.selection,
      {
        getCellRect: (cell) => this.getCellRect(cell),
        visibleRows,
        visibleCols,
      },
      {
        clipRect,
      }
    );

    if (!this.formulaBar?.isFormulaEditing() && this.fillPreviewRange) {
      const startRect = this.getCellRect({ row: this.fillPreviewRange.startRow, col: this.fillPreviewRange.startCol });
      const endRect = this.getCellRect({ row: this.fillPreviewRange.endRow, col: this.fillPreviewRange.endCol });
      if (startRect && endRect) {
        const x = startRect.x;
        const y = startRect.y;
        const width = endRect.x + endRect.width - startRect.x;
        const height = endRect.y + endRect.height - startRect.y;

        this.selectionCtx.save();
        this.selectionCtx.beginPath();
        this.selectionCtx.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height);
        this.selectionCtx.clip();
        this.selectionCtx.strokeStyle = resolveCssVar("--warning", { fallback: "CanvasText" });
        this.selectionCtx.lineWidth = 2;
        this.selectionCtx.setLineDash([4, 3]);
        this.selectionCtx.strokeRect(x + 0.5, y + 0.5, width - 1, height - 1);
        this.selectionCtx.restore();
      }
    }

    if (this.selectedDrawingId != null) {
      const objects = this.listDrawingObjectsForSheet();
      const selected = objects.find((obj) => obj.id === this.selectedDrawingId) ?? null;
      if (selected) {
        const viewport = this.getDrawingInteractionViewport();
        const rootRect = drawingObjectToViewportRect(selected, viewport, this.drawingGeom);
        const transform = selected.transform;
        const hasNonIdentityTransform = !!(
          transform &&
          (transform.rotationDeg !== 0 || transform.flipH || transform.flipV)
        );

        const stroke = resolveCssVar("--formula-grid-selection-border", {
          fallback: resolveCssVar("--selection-border", { fallback: "CanvasText" })
        });
        const handleFill = resolveCssVar("--formula-grid-bg", {
          fallback: resolveCssVar("--bg-primary", { fallback: "Canvas" })
        });
        const handleSize = RESIZE_HANDLE_SIZE_PX;
        const half = handleSize / 2;

        this.selectionCtx.save();
        this.selectionCtx.beginPath();
        this.selectionCtx.rect(clipRect.x, clipRect.y, clipRect.width, clipRect.height);
        this.selectionCtx.clip();

        this.selectionCtx.strokeStyle = stroke;
        this.selectionCtx.lineWidth = 2;
        this.selectionCtx.setLineDash([]);
        if (hasNonIdentityTransform) {
          const cx = rootRect.x + rootRect.width / 2;
          const cy = rootRect.y + rootRect.height / 2;
          this.selectionCtx.save();
          this.selectionCtx.translate(cx, cy);
          this.selectionCtx.rotate((transform!.rotationDeg * Math.PI) / 180);
          this.selectionCtx.scale(transform!.flipH ? -1 : 1, transform!.flipV ? -1 : 1);
          this.selectionCtx.strokeRect(-rootRect.width / 2, -rootRect.height / 2, rootRect.width, rootRect.height);
          this.selectionCtx.restore();
        } else {
          this.selectionCtx.strokeRect(rootRect.x, rootRect.y, rootRect.width, rootRect.height);
        }

        this.selectionCtx.fillStyle = handleFill;
        this.selectionCtx.strokeStyle = stroke;
        this.selectionCtx.lineWidth = 1;
        for (const c of getResizeHandleCenters(rootRect, transform)) {
          this.selectionCtx.beginPath();
          this.selectionCtx.rect(c.x - half, c.y - half, handleSize, handleSize);
          this.selectionCtx.fill();
          this.selectionCtx.stroke();
        }

        this.selectionCtx.restore();
      }
    }

    // If scrolling/resizing happened during editing, keep the editor aligned.
    if (this.editor.isOpen()) {
      const rect = this.getCellRect(this.selection.active);
      if (rect) this.editor.reposition(rect);
    }
  }

  private updateStatus(): void {
    const activeCell = this.selection.active;
    const activeA1 = cellToA1(activeCell);
    this.status.activeCell.textContent = activeA1;
    const selectionRangeText = (() => {
      const ranges = this.selection.ranges;
      if (ranges.length !== 1) return `${ranges.length} ranges`;
      const range = ranges[0];
      if (!range) return `${ranges.length} ranges`;
      // Excel-like name box formatting for full-row/column selections.
      if (this.selection.type === "column") {
        return `${colToName(range.startCol)}:${colToName(range.endCol)}`;
      }
      if (this.selection.type === "row") {
        return `${range.startRow + 1}:${range.endRow + 1}`;
      }
      return rangeToA1(range);
    })();
    this.status.selectionRange.textContent = selectionRangeText;

    // `getCellDisplayValue` internally recomputes the computed value. We need the computed value
    // anyway for the formula bar, so compute once and reuse to avoid duplicate formula evaluation.
    const computed = this.getCellComputedValue(activeCell);
    this.status.activeValue.textContent = computed == null ? "" : this.formatCellValueForDisplay(activeCell, computed);
    this.updateSelectionStats();

    if (this.formulaBar) {
      const input = this.formulaBar.isEditing() ? "" : this.getCellInputText(activeCell);
      this.formulaBar.setActiveCell({ address: activeA1, input, value: computed, nameBox: selectionRangeText });
      if (!this.formulaBar.isEditing()) {
        this.formulaBarCompletion?.update();
      }
    }

    this.renderCommentsPanel();

    for (const listener of this.selectionListeners) {
      listener(this.selection);
    }

    this.maybeScheduleAuditingUpdate();
  }

  private updateSelectionStats(): void {
    const sumEl = this.status.selectionSum;
    const avgEl = this.status.selectionAverage;
    const countEl = this.status.selectionCount;
    if (!sumEl && !avgEl && !countEl) return;

    const summary = this.getSelectionSummary();
    // `getSelectionSummary` returns `null` when there are no numeric values, but the
    // status bar expects stable numeric formatting (and shouldn't crash Intl formatting).
    const sum = summary.sum ?? 0;
    const avg = summary.average ?? 0;
    const count = summary.countNonEmpty;
    const formatter =
      this.selectionStatsFormatter ?? (this.selectionStatsFormatter = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 }));

    if (sumEl) sumEl.textContent = `Sum: ${formatter.format(sum)}`;
    if (avgEl) avgEl.textContent = `Avg: ${formatter.format(avg)}`;
    if (countEl) countEl.textContent = `Count: ${formatter.format(count)}`;
  }

  private syncEngineNow(): void {
    (this.engine as unknown as { syncNow?: () => void }).syncNow?.();
  }

  private handleDrawingKeyDown(e: KeyboardEvent): boolean {
    if (this.selectedDrawingId == null) return false;
    // Never hijack keys while editing text (cell editor, formula bar, inline edit).
    if (this.isEditing()) return false;

    const primary = e.ctrlKey || e.metaKey;
    const key = e.key;

    if (key === "Escape") {
      e.preventDefault();
      e.stopPropagation();
      this.selectDrawing(null);
      this.focus();
      return true;
    }

    if (key === "ArrowLeft" || key === "ArrowRight" || key === "ArrowUp" || key === "ArrowDown") {
      e.preventDefault();
      e.stopPropagation();
      const step = e.shiftKey ? 10 : 1;
      const { dxPx, dyPx } = (() => {
        switch (key) {
          case "ArrowLeft":
            return { dxPx: -step, dyPx: 0 };
          case "ArrowRight":
            return { dxPx: step, dyPx: 0 };
          case "ArrowUp":
            return { dxPx: 0, dyPx: -step };
          case "ArrowDown":
            return { dxPx: 0, dyPx: step };
        }
      })();
      // `shiftAnchor` expects deltas in screen pixels and converts to sheet units internally
      // based on the provided `zoom` value.
      this.nudgeSelectedDrawing(dxPx, dyPx);
      this.focus();
      return true;
    }

    if (key === "Delete" || key === "Backspace") {
      e.preventDefault();
      this.deleteSelectedDrawing();
      this.focus();
      return true;
    }

    if (primary && !e.altKey && !e.shiftKey && (key === "d" || key === "D")) {
      e.preventDefault();
      this.duplicateSelectedDrawing();
      this.focus();
      return true;
    }

    // Excel-like z-order shortcuts.
    if (primary && !e.altKey && !e.shiftKey && e.code === "BracketRight") {
      e.preventDefault();
      this.bringSelectedDrawingForward();
      this.focus();
      return true;
    }

    if (primary && !e.altKey && !e.shiftKey && e.code === "BracketLeft") {
      e.preventDefault();
      this.sendSelectedDrawingBackward();
      this.focus();
      return true;
    }

    // Optional: Ctrl/Cmd+Shift+]/[ for to-front/to-back.
    if (primary && !e.altKey && e.shiftKey && e.code === "BracketRight") {
      e.preventDefault();
      this.bringSelectedDrawingToFront();
      this.focus();
      return true;
    }

    if (primary && !e.altKey && e.shiftKey && e.code === "BracketLeft") {
      e.preventDefault();
      this.sendSelectedDrawingToBack();
      this.focus();
      return true;
    }

    return false;
  }

  private nudgeSelectedDrawing(dxPx: number, dyPx: number): void {
    const selectedId = this.selectedDrawingId;
    if (selectedId == null) return;
    if (!Number.isFinite(dxPx) || !Number.isFinite(dyPx) || (dxPx === 0 && dyPx === 0)) return;

    const setSheetDrawings =
      typeof (this.document as any).setSheetDrawings === "function" ? ((this.document as any).setSheetDrawings as Function) : null;
    // If drawings are backed by the document, respect read-only mode and avoid making any
    // in-memory moves that would diverge from the persisted state.
    if (setSheetDrawings && this.isReadOnly()) return;

    const objects = this.listDrawingObjectsForSheet();
    if (objects.length === 0) return;

    const index = objects.findIndex((obj) => obj.id === selectedId);
    if (index === -1) return;
    const selected = objects[index]!;

    const viewport = this.getDrawingInteractionViewport();
    const zoom = typeof viewport.zoom === "number" && Number.isFinite(viewport.zoom) && viewport.zoom > 0 ? viewport.zoom : 1;

    const nextAnchor = shiftAnchor(selected.anchor, dxPx, dyPx, this.drawingGeom, zoom);
    if (nextAnchor === selected.anchor) return;

    const nextObjects = objects.map((obj) => (obj.id === selectedId ? { ...obj, anchor: nextAnchor } : obj));

    // Update in-memory caches immediately so render/hit-test paths see the new positions even if
    // the DocumentController publishes drawing changes asynchronously.
    const docAny: any = this.document as any;
    const drawingsGetter = typeof docAny.getSheetDrawings === "function" ? docAny.getSheetDrawings : null;
    this.drawingObjectsCache = { sheetId: this.sheetId, objects: nextObjects, source: drawingsGetter };
    this.drawingHitTestIndex = null;
    this.drawingHitTestIndexObjects = null;
    this.scheduleDrawingsRender("keyboard:nudge");

    // Persist the move (and create an undo step) when the document supports sheet drawings.
    if (!setSheetDrawings) return;

    this.document.beginBatch({ label: "Move Picture" });
    try {
      setSheetDrawings.call(this.document, this.sheetId, nextObjects, { source: "drawings" });
      this.document.endBatch();
    } catch (err) {
      try {
        this.document.cancelBatch();
      } catch {
        // ignore
      }
      throw err;
    }
  }

  private handleShowFormulasShortcut(e: KeyboardEvent): boolean {
    const primary = e.ctrlKey || e.metaKey;
    if (!primary) return false;
    if (e.code !== "Backquote") return false;

    // Only trigger when *not* actively editing text.
    if (this.editor.isOpen()) return false;
    if (this.formulaBar?.isEditing()) return false;

    const target = e.target as HTMLElement | null;
    if (target) {
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return false;
    }

    e.preventDefault();
    // Prefer routing through the CommandRegistry so the command palette, ribbon,
    // and keyboard shortcuts all share a single canonical implementation. This
    // also ensures command execution tracking (and any future command-level
    // hooks/telemetry) sees the shortcut.
    const registry = typeof window !== "undefined" ? (window as any).__formulaCommandRegistry : null;
    if (registry && typeof registry.executeCommand === "function") {
      void registry.executeCommand("view.toggleShowFormulas");
    } else {
      this.toggleShowFormulas();
    }
    return true;
  }

  private handleAuditingShortcut(e: KeyboardEvent): boolean {
    const primary = e.ctrlKey || e.metaKey;
    if (!primary) return false;
    if (e.code !== "BracketLeft" && e.code !== "BracketRight") return false;

    if (this.editor.isOpen()) return false;
    if (this.formulaBar?.isEditing()) return false;

    const target = e.target as HTMLElement | null;
    if (target) {
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return false;
    }

    e.preventDefault();
    this.toggleAuditingComponent(e.code === "BracketLeft" ? "precedents" : "dependents");
    return true;
  }

  private toggleAuditingComponent(component: "precedents" | "dependents"): void {
    const hasPrecedents = this.auditingMode === "precedents" || this.auditingMode === "both";
    const hasDependents = this.auditingMode === "dependents" || this.auditingMode === "both";

    const nextPrecedents = component === "precedents" ? !hasPrecedents : hasPrecedents;
    const nextDependents = component === "dependents" ? !hasDependents : hasDependents;

    const nextMode: "off" | AuditingMode =
      nextPrecedents && nextDependents
        ? "both"
        : nextPrecedents
          ? "precedents"
          : nextDependents
            ? "dependents"
            : "off";

    if (nextMode === this.auditingMode) return;
    this.auditingMode = nextMode;
    this.auditingErrors = { precedents: null, dependents: null };

    if (nextMode === "off") {
      this.auditingHighlights = { precedents: new Set(), dependents: new Set() };
      this.updateAuditingLegend();
      this.renderAuditing();
      return;
    }

    this.updateAuditingLegend();
    this.scheduleAuditingUpdate();
    this.renderAuditing();
  }

  private auditingCacheKey(sheetId: string, row: number, col: number, transitive: boolean): string {
    return `${sheetId}:${row},${col}:${transitive ? "t" : "d"}`;
  }

  private getTauriInvoke(): ((cmd: string, args?: Record<string, unknown>) => Promise<unknown>) | null {
    const queued = (globalThis as any).__formulaQueuedInvoke;
    if (typeof queued === "function") {
      return queued;
    }

    const invoke = (globalThis as any).__TAURI__?.core?.invoke;
    if (typeof invoke !== "function") return null;
    return invoke;
  }

  private maybeScheduleAuditingUpdate(): void {
    if (this.auditingMode === "off") return;

    if (this.dragState) {
      this.auditingNeedsUpdateAfterDrag = true;
      return;
    }

    const cell = this.selection.active;
    const key = this.auditingCacheKey(this.sheetId, cell.row, cell.col, this.auditingTransitive);
    if (key === this.auditingLastCellKey) return;

    this.scheduleAuditingUpdate();
  }

  private scheduleAuditingUpdate(): void {
    if (this.auditingMode === "off") return;
    if (this.auditingUpdateScheduled) return;
    this.auditingUpdateScheduled = true;

    const update = new Promise<void>((resolve) => {
      queueMicrotask(() => {
        this.auditingUpdateScheduled = false;
        this.updateAuditingForActiveCell()
          .catch(() => {})
          .finally(() => resolve());
      });
    });
    this.auditingIdlePromise = update;
  }

  private async updateAuditingForActiveCell(): Promise<void> {
    if (this.auditingMode === "off") return;

    const invoke = this.getTauriInvoke();
    const cell = this.selection.active;
    const key = this.auditingCacheKey(this.sheetId, cell.row, cell.col, this.auditingTransitive);
    const cellA1 = cellToA1(cell);

    this.auditingLastCellKey = key;

    const cached = this.auditingCache.get(key);
    if (cached) {
      this.applyAuditingEntry(cached, cellA1);
      return;
    }

    if (!invoke) {
      const entry: AuditingCacheEntry = {
        precedents: [],
        dependents: [],
        precedentsError: "Auditing requires the desktop engine.",
        dependentsError: "Auditing requires the desktop engine.",
      };
      this.auditingCache.set(key, entry);
      this.applyAuditingEntry(entry, cellA1);
      return;
    }

    const requestId = ++this.auditingRequestId;

    const args = {
      sheet_id: this.sheetId,
      row: cell.row,
      col: cell.col,
      transitive: this.auditingTransitive,
    };

    let precedents: string[] = [];
    let dependents: string[] = [];
    let precedentsError: string | null = null;
    let dependentsError: string | null = null;

    try {
      const payload = await invoke("get_precedents", args);
      precedents = Array.isArray(payload) ? payload.map(String) : [];
    } catch (err) {
      precedentsError = String(err);
      precedents = [];
    }

    try {
      const payload = await invoke("get_dependents", args);
      dependents = Array.isArray(payload) ? payload.map(String) : [];
    } catch (err) {
      dependentsError = String(err);
      dependents = [];
    }

    const entry: AuditingCacheEntry = { precedents, dependents, precedentsError, dependentsError };
    this.auditingCache.set(key, entry);

    if (requestId !== this.auditingRequestId) return;
    if (key !== this.auditingLastCellKey) return;

    this.applyAuditingEntry(entry, cellA1);
  }

  private applyAuditingEntry(entry: AuditingCacheEntry, cellA1: string): void {
    const engine = {
      precedents: () => entry.precedents,
      dependents: () => entry.dependents,
    };

    const mode = this.auditingMode === "off" ? "both" : this.auditingMode;
    this.auditingHighlights = computeAuditingOverlays(engine, cellA1, mode, { transitive: this.auditingTransitive });
    this.auditingErrors = { precedents: entry.precedentsError, dependents: entry.dependentsError };

    this.updateAuditingLegend();
    this.renderAuditing();
  }

  private handleUndoRedoShortcut(e: KeyboardEvent): boolean {
    const undo = isUndoKeyboardEvent(e);
    const redo = !undo && isRedoKeyboardEvent(e);
    if (!undo && !redo) return false;

    // Only trigger spreadsheet undo/redo when *not* actively editing text.
    if (this.editor.isOpen()) return false;
    if (this.formulaBar?.isEditing()) return false;

    const target = e.target as HTMLElement | null;
    if (target) {
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return false;
    }

    // In read-only collab roles (viewer/commenter), we must not mutate local state via undo/redo.
    // Those changes will not sync, and can leave the UI diverged from the authoritative remote doc.
    if (this.isReadOnly()) {
      e.preventDefault();
      return true;
    }

    e.preventDefault();
    this.applyUndoRedo(undo ? "undo" : "redo");
    return true;
  }

  private applyUndoRedo(kind: "undo" | "redo"): boolean {
    let did = false;
    if (this.collabUndoService) {
      if (kind === "undo") {
        if (this.collabUndoService.canUndo()) {
          this.collabUndoService.undo();
          did = true;
        }
      } else {
        if (this.collabUndoService.canRedo()) {
          this.collabUndoService.redo();
          did = true;
        }
      }
    } else {
      did = kind === "undo" ? this.document.undo() : this.document.redo();
    }
    if (!did) return false;

    this.syncEngineNow();

    // Undo/redo can add/remove/hide sheets. Ensure the app doesn't keep rendering
    // a sheet that no longer exists (DocumentController lazily materializes sheets
    // on access, which would otherwise resurrect deleted sheets).
    const sheetIds = this.document.getSheetIds();
    const visibleSheetIds = this.document.getVisibleSheetIds();

    const hasSheet = sheetIds.includes(this.sheetId);
    const isVisible = visibleSheetIds.includes(this.sheetId);
    if ((!hasSheet || !isVisible) && sheetIds.length > 0) {
      const fallback = visibleSheetIds[0] ?? sheetIds[0];
      if (fallback && fallback !== this.sheetId) {
        this.activateSheet(fallback);
      }
    }

    // Undo/redo can affect sheet view state (e.g. frozen panes). Keep renderer + scrollbars in sync.
    this.syncFrozenPanes();
    if (this.sharedGrid) {
      // Shared grid rendering is driven by CanvasGridRenderer, but we still need to refresh
      // overlays (charts, auditing, etc) after changes.
      this.refresh();
    }

    // If the user is undoing/redoing comment edits before the provider has synced
    // (and before the comments root observer is attached), refresh the comment
    // indexes/panel eagerly so indicators and the sidebar stay in sync.
    this.maybeRefreshCommentsUiForLocalEdit();
    return true;
  }

  private isRowHidden(row: number): boolean {
    if (!Number.isInteger(row) || row < 0 || row >= this.limits.maxRows) return false;
    const entry = this.getOutlineForSheet(this.sheetId).rows.entry(row + 1);
    if (this.gridMode === "legacy") return isHidden(entry.hidden);
    return entry.hidden.user;
  }

  private isColHidden(col: number): boolean {
    if (!Number.isInteger(col) || col < 0 || col >= this.limits.maxCols) return false;
    const entry = this.getOutlineForSheet(this.sheetId).cols.entry(col + 1);
    if (this.gridMode === "legacy") return isHidden(entry.hidden);
    return entry.hidden.user;
  }

  private rebuildAxisVisibilityCache(): void {
    if (this.sharedGrid) {
      // Shared-grid mode intentionally does not build the legacy row/col visibility caches.
      // Building them is O(maxRows/maxCols) and would be prohibitively expensive for large sheets.
      return;
    }
    this.rowIndexByVisual = [];
    this.colIndexByVisual = [];
    this.rowToVisual.clear();
    this.colToVisual.clear();

    for (let r = 0; r < this.limits.maxRows; r += 1) {
      if (this.isRowHidden(r)) continue;
      this.rowToVisual.set(r, this.rowIndexByVisual.length);
      this.rowIndexByVisual.push(r);
    }

    for (let c = 0; c < this.limits.maxCols; c += 1) {
      if (this.isColHidden(c)) continue;
      this.colToVisual.set(c, this.colIndexByVisual.length);
      this.colIndexByVisual.push(c);
    }

    // Outline changes can affect scrollable content size.
    const didClamp = this.clampScroll();
    if (didClamp) this.hideCommentTooltip();
    this.syncScrollbars();
    if (didClamp) this.notifyScrollListeners();
  }

  private viewportWidth(): number {
    if (this.sharedGrid) {
      const viewport = this.sharedGrid.renderer.scroll.getViewportState();
      const frozenWidth = Math.min(viewport.frozenWidth, viewport.width);
      return Math.max(0, viewport.width - frozenWidth);
    }
    return Math.max(0, this.width - this.rowHeaderWidth);
  }

  private viewportHeight(): number {
    if (this.sharedGrid) {
      const viewport = this.sharedGrid.renderer.scroll.getViewportState();
      const frozenHeight = Math.min(viewport.frozenHeight, viewport.height);
      return Math.max(0, viewport.height - frozenHeight);
    }
    return Math.max(0, this.height - this.colHeaderHeight);
  }

  private contentWidth(): number {
    return this.colIndexByVisual.length * this.cellWidth;
  }

  private contentHeight(): number {
    return this.rowIndexByVisual.length * this.cellHeight;
  }

  private maxScrollX(): number {
    if (this.sharedGrid) {
      return this.sharedGrid.renderer.scroll.getViewportState().maxScrollX;
    }
    return Math.max(0, this.contentWidth() - this.viewportWidth());
  }

  private maxScrollY(): number {
    if (this.sharedGrid) {
      return this.sharedGrid.renderer.scroll.getViewportState().maxScrollY;
    }
    return Math.max(0, this.contentHeight() - this.viewportHeight());
  }

  private clampScroll(): boolean {
    const prevX = this.scrollX;
    const prevY = this.scrollY;

    if (this.sharedGrid) {
      const scroll = this.sharedGrid.getScroll();
      this.scrollX = scroll.x;
      this.scrollY = scroll.y;
      return this.scrollX !== prevX || this.scrollY !== prevY;
    }

    const maxX = this.maxScrollX();
    const maxY = this.maxScrollY();
    this.scrollX = Math.min(Math.max(0, this.scrollX), maxX);
    this.scrollY = Math.min(Math.max(0, this.scrollY), maxY);
    return this.scrollX !== prevX || this.scrollY !== prevY;
  }

  private setScrollInternal(nextX: number, nextY: number): boolean {
    if (!Number.isFinite(nextX)) nextX = 0;
    if (!Number.isFinite(nextY)) nextY = 0;
    if (this.sharedGrid) {
      const before = this.sharedGrid.getScroll();
      this.sharedGrid.scrollTo(nextX, nextY);
      const after = this.sharedGrid.getScroll();
      this.scrollX = after.x;
      this.scrollY = after.y;
      return before.x !== after.x || before.y !== after.y;
    }

    const prevX = this.scrollX;
    const prevY = this.scrollY;
    this.scrollX = nextX;
    this.scrollY = nextY;
    this.clampScroll();

    const changed = this.scrollX !== prevX || this.scrollY !== prevY;
    if (changed) {
      this.hideCommentTooltip();
      this.syncScrollbars();
      this.notifyScrollListeners();
    }
    return changed;
  }

  private scrollBy(deltaX: number, deltaY: number): void {
    const changed = this.setScrollInternal(this.scrollX + deltaX, this.scrollY + deltaY);
    if (changed) this.refresh("scroll");
  }

  private updateViewportMapping(): void {
    const availableWidth = this.viewportWidth();
    const availableHeight = this.viewportHeight();

    const totalRows = this.rowIndexByVisual.length;
    const totalCols = this.colIndexByVisual.length;

    const view = this.document.getSheetView(this.sheetId) as { frozenRows?: number; frozenCols?: number } | null;
    const nextFrozenRows = Number.isFinite(view?.frozenRows) ? Math.max(0, Math.trunc(view?.frozenRows ?? 0)) : 0;
    const nextFrozenCols = Number.isFinite(view?.frozenCols) ? Math.max(0, Math.trunc(view?.frozenCols ?? 0)) : 0;

    this.frozenRows = Math.min(nextFrozenRows, this.limits.maxRows);
    this.frozenCols = Math.min(nextFrozenCols, this.limits.maxCols);

    const frozenRowCount = this.lowerBound(this.rowIndexByVisual, this.frozenRows);
    const frozenColCount = this.lowerBound(this.colIndexByVisual, this.frozenCols);
    this.frozenVisibleRows = this.rowIndexByVisual.slice(0, frozenRowCount);
    this.frozenVisibleCols = this.colIndexByVisual.slice(0, frozenColCount);
    this.frozenHeight = frozenRowCount * this.cellHeight;
    this.frozenWidth = frozenColCount * this.cellWidth;

    const overscan = 1;
    const firstRow = Math.max(
      frozenRowCount,
      Math.floor((this.scrollY + this.frozenHeight) / this.cellHeight) - overscan
    );
    const lastRow = Math.min(totalRows, Math.ceil((this.scrollY + availableHeight) / this.cellHeight) + overscan);

    const firstCol = Math.max(
      frozenColCount,
      Math.floor((this.scrollX + this.frozenWidth) / this.cellWidth) - overscan
    );
    const lastCol = Math.min(totalCols, Math.ceil((this.scrollX + availableWidth) / this.cellWidth) + overscan);

    this.visibleRowStart = firstRow;
    this.visibleColStart = firstCol;
    this.visibleRows = this.rowIndexByVisual.slice(firstRow, lastRow);
    this.visibleCols = this.colIndexByVisual.slice(firstCol, lastCol);
  }

  private ensureViewportMappingCurrent(): void {
    const viewportWidth = this.viewportWidth();
    const viewportHeight = this.viewportHeight();
    const rowCount = this.rowIndexByVisual.length;
    const colCount = this.colIndexByVisual.length;

    const state = this.viewportMappingState;
    if (
      state &&
      state.scrollX === this.scrollX &&
      state.scrollY === this.scrollY &&
      state.viewportWidth === viewportWidth &&
      state.viewportHeight === viewportHeight &&
      state.rowCount === rowCount &&
      state.colCount === colCount
    ) {
      return;
    }

    this.updateViewportMapping();
    this.viewportMappingState = { scrollX: this.scrollX, scrollY: this.scrollY, viewportWidth, viewportHeight, rowCount, colCount };
  }

  private computeScrollbarThumb(options: {
    scrollPos: number;
    viewportSize: number;
    contentSize: number;
    trackSize: number;
    minThumbSize?: number;
    out?: { size: number; offset: number };
  }): { size: number; offset: number } {
    const minThumbSize = options.minThumbSize ?? 24;
    const trackSize = Math.max(0, options.trackSize);
    const viewportSize = Math.max(0, options.viewportSize);
    const contentSize = Math.max(0, options.contentSize);
    const maxScroll = Math.max(0, contentSize - viewportSize);
    const scrollPos = Math.min(Math.max(0, options.scrollPos), maxScroll);

    const out = options.out;

    if (trackSize === 0) {
      if (out) {
        out.size = 0;
        out.offset = 0;
        return out;
      }
      return { size: 0, offset: 0 };
    }
    if (contentSize === 0 || maxScroll === 0) {
      if (out) {
        out.size = trackSize;
        out.offset = 0;
        return out;
      }
      return { size: trackSize, offset: 0 };
    }

    const rawThumbSize = (viewportSize / contentSize) * trackSize;
    const thumbSize = Math.min(trackSize, Math.max(minThumbSize, rawThumbSize));
    const thumbTravel = Math.max(0, trackSize - thumbSize);
    const offset = thumbTravel === 0 ? 0 : (scrollPos / maxScroll) * thumbTravel;

    if (out) {
      out.size = thumbSize;
      out.offset = offset;
      return out;
    }

    return { size: thumbSize, offset };
  }

  private syncScrollbars(): void {
    if (this.sharedGrid) {
      this.sharedGrid.syncScrollbars();
      return;
    }

    // Keep frozen metrics up-to-date so scrollbars reflect the scrollable region.
    this.updateViewportMapping();
    const maxX = this.maxScrollX();
    const maxY = this.maxScrollY();
    const showH = maxX > 0;
    const showV = maxY > 0;

    const padding = 2;
    const thickness = this.scrollbarThickness;

    const prevLayout = this.lastScrollbarLayout;
    const layoutChanged =
      prevLayout === null ||
      prevLayout.showV !== showV ||
      prevLayout.showH !== showH ||
      prevLayout.rowHeaderWidth !== this.rowHeaderWidth ||
      prevLayout.colHeaderHeight !== this.colHeaderHeight ||
      prevLayout.thickness !== thickness;

    if (layoutChanged) {
      // Track layout is a function of scrollbar visibility + header sizes; avoid rewriting these
      // styles on every scroll event.
      this.vScrollbarTrack.style.display = showV ? "block" : "none";
      this.hScrollbarTrack.style.display = showH ? "block" : "none";

      if (showV) {
        this.vScrollbarTrack.style.right = `${padding}px`;
        this.vScrollbarTrack.style.top = `${this.colHeaderHeight + padding}px`;
        this.vScrollbarTrack.style.bottom = `${(showH ? thickness : 0) + padding}px`;
        this.vScrollbarTrack.style.width = `${thickness}px`;
      } else {
        this.lastScrollbarThumb.vSize = null;
        this.lastScrollbarThumb.vOffset = null;
      }

      if (showH) {
        this.hScrollbarTrack.style.left = `${this.rowHeaderWidth + padding}px`;
        this.hScrollbarTrack.style.right = `${(showV ? thickness : 0) + padding}px`;
        this.hScrollbarTrack.style.bottom = `${padding}px`;
        this.hScrollbarTrack.style.height = `${thickness}px`;
      } else {
        this.lastScrollbarThumb.hSize = null;
        this.lastScrollbarThumb.hOffset = null;
      }

      if (prevLayout) {
        prevLayout.showV = showV;
        prevLayout.showH = showH;
        prevLayout.rowHeaderWidth = this.rowHeaderWidth;
        prevLayout.colHeaderHeight = this.colHeaderHeight;
        prevLayout.thickness = thickness;
      } else {
        this.lastScrollbarLayout = {
          showV,
          showH,
          rowHeaderWidth: this.rowHeaderWidth,
          colHeaderHeight: this.colHeaderHeight,
          thickness
        };
      }
    }

    if (showV) {
      const trackSize = Math.max(0, this.height - (this.colHeaderHeight + padding) - ((showH ? thickness : 0) + padding));
      const { size, offset } = this.computeScrollbarThumb({
        scrollPos: this.scrollY,
        viewportSize: Math.max(0, this.viewportHeight() - this.frozenHeight),
        contentSize: Math.max(0, this.contentHeight() - this.frozenHeight),
        trackSize,
        out: this.scrollbarThumbScratch.v
      });

      if (this.lastScrollbarThumb.vSize !== size) {
        this.vScrollbarThumb.style.height = `${size}px`;
        this.lastScrollbarThumb.vSize = size;
      }
      if (this.lastScrollbarThumb.vOffset !== offset) {
        this.vScrollbarThumb.style.transform = `translateY(${offset}px)`;
        this.lastScrollbarThumb.vOffset = offset;
      }
    } else {
      this.lastScrollbarThumb.vSize = null;
      this.lastScrollbarThumb.vOffset = null;
    }

    if (showH) {
      const trackSize = Math.max(0, this.width - (this.rowHeaderWidth + padding) - ((showV ? thickness : 0) + padding));
      const { size, offset } = this.computeScrollbarThumb({
        scrollPos: this.scrollX,
        viewportSize: Math.max(0, this.viewportWidth() - this.frozenWidth),
        contentSize: Math.max(0, this.contentWidth() - this.frozenWidth),
        trackSize,
        out: this.scrollbarThumbScratch.h
      });

      if (this.lastScrollbarThumb.hSize !== size) {
        this.hScrollbarThumb.style.width = `${size}px`;
        this.lastScrollbarThumb.hSize = size;
      }
      if (this.lastScrollbarThumb.hOffset !== offset) {
        this.hScrollbarThumb.style.transform = `translateX(${offset}px)`;
        this.lastScrollbarThumb.hOffset = offset;
      }
    } else {
      this.lastScrollbarThumb.hSize = null;
      this.lastScrollbarThumb.hOffset = null;
    }
  }

  private onWheel(e: WheelEvent): void {
    const target = e.target as HTMLElement | null;
    if (target?.closest('[data-testid="comments-panel"]')) return;
    if (e.ctrlKey) return;

    let deltaX = wheelDeltaToPixels(e.deltaX, e.deltaMode, { pageSize: this.viewportWidth() });
    let deltaY = wheelDeltaToPixels(e.deltaY, e.deltaMode, { pageSize: this.viewportHeight() });

    // Common UX: shift+wheel scrolls horizontally.
    if (e.shiftKey && deltaX === 0) {
      deltaX = deltaY;
      deltaY = 0;
    }

    if (deltaX === 0 && deltaY === 0) return;
    e.preventDefault();
    this.scrollBy(deltaX, deltaY);
  }

  private onGridDragOver(e: DragEvent): void {
    if (this.isEditing()) return;
    const dt = e.dataTransfer;
    if (!dt) return;
    const types = Array.from(dt.types ?? []);
    if (!types.includes("Files")) return;

    // Only allow drop if at least one file is an image. Prefer `items` metadata during
    // dragover; `files` may not be populated until drop in some browsers.
    const hasImage =
      Array.from(dt.items ?? []).some((item) => item.kind === "file" && item.type.startsWith("image/")) ||
      Array.from(dt.files ?? []).some((file) => file.type.startsWith("image/"));
    if (!hasImage) return;

    e.preventDefault();
    try {
      dt.dropEffect = "copy";
    } catch {
      // ignore
    }
  }

  private onGridDrop(e: DragEvent): void {
    if (this.isEditing()) return;
    const dt = e.dataTransfer;
    if (!dt) return;

    const imageFiles = Array.from(dt.files ?? []).filter((file) => file.type.startsWith("image/"));
    if (imageFiles.length === 0) return;

    e.preventDefault();

    const placeAt = this.pickCellAtClientPoint(e.clientX, e.clientY) ?? this.getActiveCell();
    void this.insertPicturesFromFiles(imageFiles, { placeAt });
  }

  private onScrollbarThumbPointerDown(e: PointerEvent, axis: "x" | "y"): void {
    e.preventDefault();
    e.stopPropagation();

    const thumb = axis === "y" ? this.vScrollbarThumb : this.hScrollbarThumb;
    const track = axis === "y" ? this.vScrollbarTrack : this.hScrollbarTrack;

    const trackRect = track.getBoundingClientRect();
    const thumbRect = thumb.getBoundingClientRect();

    const pointerPos = axis === "y" ? e.clientY : e.clientX;
    const thumbStart = axis === "y" ? thumbRect.top : thumbRect.left;
    const trackStart = axis === "y" ? trackRect.top : trackRect.left;
    const trackSize = axis === "y" ? trackRect.height : trackRect.width;
    const thumbSize = axis === "y" ? thumbRect.height : thumbRect.width;

    const grabOffset = pointerPos - thumbStart;
    const thumbTravel = Math.max(0, trackSize - thumbSize);
    const maxScroll = axis === "y" ? this.maxScrollY() : this.maxScrollX();

    this.scrollbarDrag = { axis, pointerId: e.pointerId, grabOffset, thumbTravel, trackStart, maxScroll };

    (thumb as HTMLElement).setPointerCapture(e.pointerId);
  }

  private onScrollbarTrackPointerDown(e: PointerEvent, axis: "x" | "y"): void {
    // Clicking the track should scroll, but should not start a selection drag.
    e.preventDefault();
    e.stopPropagation();

    const thumb = axis === "y" ? this.vScrollbarThumb : this.hScrollbarThumb;
    const track = axis === "y" ? this.vScrollbarTrack : this.hScrollbarTrack;
    const trackRect = track.getBoundingClientRect();
    const thumbRect = thumb.getBoundingClientRect();

    const trackSize = axis === "y" ? trackRect.height : trackRect.width;
    const thumbSize = axis === "y" ? thumbRect.height : thumbRect.width;
    const thumbTravel = Math.max(0, trackSize - thumbSize);
    if (thumbTravel === 0) return;

    const pointerPos = axis === "y" ? e.clientY - trackRect.top : e.clientX - trackRect.left;
    const targetOffset = pointerPos - thumbSize / 2;
    const clamped = Math.min(Math.max(0, targetOffset), thumbTravel);

    const maxScroll = axis === "y" ? this.maxScrollY() : this.maxScrollX();
    const nextScroll = (clamped / thumbTravel) * maxScroll;

    const changed =
      axis === "y"
        ? this.setScrollInternal(this.scrollX, nextScroll)
        : this.setScrollInternal(nextScroll, this.scrollY);
    if (changed) this.refresh("scroll");
  }

  private onScrollbarThumbPointerMove(e: PointerEvent): void {
    const drag = this.scrollbarDrag;
    if (!drag) return;
    if (e.pointerId !== drag.pointerId) return;

    const pointerPos = drag.axis === "y" ? e.clientY : e.clientX;
    const thumbOffset = pointerPos - drag.trackStart - drag.grabOffset;
    const clamped = Math.min(Math.max(0, thumbOffset), drag.thumbTravel);
    const nextScroll = drag.thumbTravel === 0 ? 0 : (clamped / drag.thumbTravel) * drag.maxScroll;

    const changed =
      drag.axis === "y"
        ? this.setScrollInternal(this.scrollX, nextScroll)
        : this.setScrollInternal(nextScroll, this.scrollY);

    if (changed) this.refresh("scroll");
  }

  private onScrollbarThumbPointerUp(e: PointerEvent): void {
    const drag = this.scrollbarDrag;
    if (!drag) return;
    if (e.pointerId !== drag.pointerId) return;
    this.scrollbarDrag = null;
  }

  private scrollCellIntoView(cell: CellCoord, paddingPx = 8): boolean {
    if (this.sharedGrid) {
      const before = this.sharedGrid.getScroll();
      const gridCell = this.gridCellFromDocCell(cell);
      this.sharedGrid.scrollToCell(gridCell.row, gridCell.col, { align: "auto", padding: paddingPx });
      const after = this.sharedGrid.getScroll();
      this.scrollX = after.x;
      this.scrollY = after.y;
      return before.x !== after.x || before.y !== after.y;
    }

    // Ensure frozen pane metrics are current even when called outside `renderGrid()`.
    this.updateViewportMapping();

    const visualRow = this.rowToVisual.get(cell.row);
    const visualCol = this.colToVisual.get(cell.col);
    if (visualRow === undefined || visualCol === undefined) return false;

    const viewportWidth = this.viewportWidth();
    const viewportHeight = this.viewportHeight();
    if (viewportWidth <= 0 || viewportHeight <= 0) return false;

    const left = visualCol * this.cellWidth;
    const top = visualRow * this.cellHeight;
    const right = left + this.cellWidth;
    const bottom = top + this.cellHeight;

    const pad = Math.max(0, paddingPx);
    let nextX = this.scrollX;
    let nextY = this.scrollY;

    if (cell.col >= this.frozenCols) {
      const minX = nextX + this.frozenWidth + pad;
      const maxX = nextX + viewportWidth - pad;
      if (left < minX) {
        nextX = left - this.frozenWidth - pad;
      } else if (right > maxX) {
        nextX = right - viewportWidth + pad;
      }
    }

    if (cell.row >= this.frozenRows) {
      const minY = nextY + this.frozenHeight + pad;
      const maxY = nextY + viewportHeight - pad;
      if (top < minY) {
        nextY = top - this.frozenHeight - pad;
      } else if (bottom > maxY) {
        nextY = bottom - viewportHeight + pad;
      }
    }

    return this.setScrollInternal(nextX, nextY);
  }

  private scrollCellToCenter(cell: CellCoord): boolean {
    if (this.sharedGrid) {
      const before = this.sharedGrid.getScroll();
      const gridCell = this.gridCellFromDocCell(cell);
      this.sharedGrid.scrollToCell(gridCell.row, gridCell.col, { align: "center" });
      const after = this.sharedGrid.getScroll();
      this.scrollX = after.x;
      this.scrollY = after.y;
      return before.x !== after.x || before.y !== after.y;
    }

    // Ensure frozen pane metrics are current even when called outside `renderGrid()`.
    this.updateViewportMapping();

    if (this.rowIndexByVisual.length === 0 || this.colIndexByVisual.length === 0) return false;

    // Hidden rows/cols collapse to zero size; treat them as sharing the origin of the next visible
    // row/col so scrolling remains stable.
    const visualRowRaw = this.rowToVisual.get(cell.row) ?? this.lowerBound(this.rowIndexByVisual, cell.row);
    const visualColRaw = this.colToVisual.get(cell.col) ?? this.lowerBound(this.colIndexByVisual, cell.col);
    const visualRow = Math.max(0, Math.min(this.rowIndexByVisual.length - 1, visualRowRaw));
    const visualCol = Math.max(0, Math.min(this.colIndexByVisual.length - 1, visualColRaw));

    const viewportWidth = this.viewportWidth();
    const viewportHeight = this.viewportHeight();
    if (viewportWidth <= 0 || viewportHeight <= 0) return false;

    const scrollableViewportWidth = Math.max(0, viewportWidth - this.frozenWidth);
    const scrollableViewportHeight = Math.max(0, viewportHeight - this.frozenHeight);

    let nextX = this.scrollX;
    let nextY = this.scrollY;

    if (scrollableViewportWidth > 0 && cell.col >= this.frozenCols) {
      nextX =
        visualCol * this.cellWidth + this.cellWidth / 2 - (this.frozenWidth + scrollableViewportWidth / 2);
    }

    if (scrollableViewportHeight > 0 && cell.row >= this.frozenRows) {
      nextY =
        visualRow * this.cellHeight + this.cellHeight / 2 - (this.frozenHeight + scrollableViewportHeight / 2);
    }

    return this.setScrollInternal(nextX, nextY);
  }

  private scrollRangeIntoView(range: Range, paddingPx = 8): boolean {
    if (this.sharedGrid) {
      const before = this.sharedGrid.getScroll();
      const start = this.gridCellFromDocCell({ row: range.startRow, col: range.startCol });
      const end = this.gridCellFromDocCell({ row: range.endRow, col: range.endCol });
      const pad = Math.max(0, paddingPx);
      // Best-effort: attempt to bring both the start and end cells into view.
      // If the range is larger than the viewport this will still keep the active cell visible.
      this.sharedGrid.scrollToCell(start.row, start.col, { align: "start", padding: pad });
      this.sharedGrid.scrollToCell(end.row, end.col, { align: "end", padding: pad });
      const after = this.sharedGrid.getScroll();
      this.scrollX = after.x;
      this.scrollY = after.y;
      return before.x !== after.x || before.y !== after.y;
    }

    // Ensure frozen pane metrics are current even when called outside `renderGrid()`.
    this.updateViewportMapping();

    const startRow = Math.max(0, Math.min(this.limits.maxRows - 1, range.startRow));
    const endRow = Math.max(0, Math.min(this.limits.maxRows - 1, range.endRow));
    const startCol = Math.max(0, Math.min(this.limits.maxCols - 1, range.startCol));
    const endCol = Math.max(0, Math.min(this.limits.maxCols - 1, range.endCol));

    const startVisualRow = this.visualIndexForRow(startRow);
    const endVisualRow = this.visualIndexForRow(endRow);
    const startVisualCol = this.visualIndexForCol(startCol);
    const endVisualCol = this.visualIndexForCol(endCol);

    const left = Math.min(startVisualCol, endVisualCol) * this.cellWidth;
    const top = Math.min(startVisualRow, endVisualRow) * this.cellHeight;
    const right = (Math.max(startVisualCol, endVisualCol) + 1) * this.cellWidth;
    const bottom = (Math.max(startVisualRow, endVisualRow) + 1) * this.cellHeight;

    const viewportWidth = this.viewportWidth();
    const viewportHeight = this.viewportHeight();
    if (viewportWidth <= 0 || viewportHeight <= 0) return false;

    const pad = Math.max(0, paddingPx);
    let nextX = this.scrollX;
    let nextY = this.scrollY;

    const affectsX = Math.max(startCol, endCol) >= this.frozenCols;
    const affectsY = Math.max(startRow, endRow) >= this.frozenRows;

    const scrollableViewportWidth = Math.max(0, viewportWidth - this.frozenWidth);
    const scrollableViewportHeight = Math.max(0, viewportHeight - this.frozenHeight);
    const scrollableLeft = Math.max(left, this.frozenWidth);
    const scrollableTop = Math.max(top, this.frozenHeight);

    // Only attempt to fully fit the range when it fits within the viewport.
    // Otherwise, fall back to keeping the active cell visible.
    if (affectsX && right - scrollableLeft <= scrollableViewportWidth - pad * 2) {
      const minX = nextX + this.frozenWidth + pad;
      const maxX = nextX + viewportWidth - pad;
      if (scrollableLeft < minX) {
        nextX = scrollableLeft - this.frozenWidth - pad;
      } else if (right > maxX) {
        nextX = right - viewportWidth + pad;
      }
    }

    if (affectsY && bottom - scrollableTop <= scrollableViewportHeight - pad * 2) {
      const minY = nextY + this.frozenHeight + pad;
      const maxY = nextY + viewportHeight - pad;
      if (scrollableTop < minY) {
        nextY = scrollableTop - this.frozenHeight - pad;
      } else if (bottom > maxY) {
        nextY = bottom - viewportHeight + pad;
      }
    }

    return this.setScrollInternal(nextX, nextY);
  }

  private renderOutlineControls(): void {
    // Outline controls rely on the legacy renderer's row/col visibility caches.
    // Shared-grid mode intentionally does not implement outline-group collapsing yet (only user-hidden
    // rows/cols are supported), so keep the outline toggle controls disabled.
    if (this.sharedGrid) {
      for (const button of this.outlineButtons.values()) button.remove();
      this.outlineButtons.clear();
      return;
    }

    const outline = this.getOutlineForSheet(this.sheetId);
    if (!outline.pr.showOutlineSymbols) {
      for (const button of this.outlineButtons.values()) button.remove();
      this.outlineButtons.clear();
      return;
    }

    const keep = new Set<string>();
    const size = 14;
    const padding = 4;
    const originX = this.rowHeaderWidth;
    const originY = this.colHeaderHeight;

    // Row group toggles live in the row header.
    for (let visualRow = 0; visualRow < this.visibleRows.length; visualRow++) {
      const rowIndex = this.visibleRows[visualRow]!;
      const summaryIndex = rowIndex + 1; // 1-based
      const entry = outline.rows.entry(summaryIndex);
      const details = groupDetailRange(outline.rows, summaryIndex, entry.level, outline.pr.summaryBelow);
      if (!details) continue;

      const key = `row:${summaryIndex}`;
      keep.add(key);
      let button = this.outlineButtons.get(key);
      if (!button) {
        button = document.createElement("button");
        button.className = "outline-toggle";
        button.type = "button";
        button.setAttribute("data-testid", `outline-toggle-row-${summaryIndex}`);
        button.addEventListener("click", (e) => {
          e.preventDefault();
          e.stopPropagation();
          this.getOutlineForSheet(this.sheetId).toggleRowGroup(summaryIndex);
          this.onOutlineUpdated();
        });
        button.addEventListener("pointerdown", (e) => {
          e.stopPropagation();
        });
        this.outlineButtons.set(key, button);
        this.outlineLayer.appendChild(button);
      }

      button.textContent = entry.collapsed ? "+" : "-";
      button.style.left = `${padding}px`;
      const visualIndex = this.visibleRowStart + visualRow;
      const rowTop = originY + visualIndex * this.cellHeight - this.scrollY;
      button.style.top = `${rowTop + (this.cellHeight - size) / 2}px`;
      button.style.width = `${size}px`;
      button.style.height = `${size}px`;
    }

    // Column group toggles live in the column header.
    for (let visualCol = 0; visualCol < this.visibleCols.length; visualCol++) {
      const colIndex = this.visibleCols[visualCol]!;
      const summaryIndex = colIndex + 1; // 1-based
      const entry = outline.cols.entry(summaryIndex);
      const details = groupDetailRange(outline.cols, summaryIndex, entry.level, outline.pr.summaryRight);
      if (!details) continue;

      const key = `col:${summaryIndex}`;
      keep.add(key);
      let button = this.outlineButtons.get(key);
      if (!button) {
        button = document.createElement("button");
        button.className = "outline-toggle";
        button.type = "button";
        button.setAttribute("data-testid", `outline-toggle-col-${summaryIndex}`);
        button.addEventListener("click", (e) => {
          e.preventDefault();
          e.stopPropagation();
          this.getOutlineForSheet(this.sheetId).toggleColGroup(summaryIndex);
          this.onOutlineUpdated();
        });
        button.addEventListener("pointerdown", (e) => {
          e.stopPropagation();
        });
        this.outlineButtons.set(key, button);
        this.outlineLayer.appendChild(button);
      }

      button.textContent = entry.collapsed ? "+" : "-";
      const visualIndex = this.visibleColStart + visualCol;
      const colLeft = originX + visualIndex * this.cellWidth - this.scrollX;
      button.style.left = `${colLeft + (this.cellWidth - size) / 2}px`;
      button.style.top = `${padding}px`;
      button.style.width = `${size}px`;
      button.style.height = `${size}px`;
    }

    for (const [key, button] of this.outlineButtons) {
      if (keep.has(key)) continue;
      button.remove();
      this.outlineButtons.delete(key);
    }
  }

  private onOutlineUpdated(): void {
    if (this.gridMode === "legacy") {
      this.rebuildAxisVisibilityCache();
    }
    this.ensureActiveCellVisible();
    if (this.sharedGrid) this.syncSharedGridAxisSizesFromDocument();
    this.scrollCellIntoView(this.selection.active);
    if (this.sharedGrid) this.syncSharedGridSelectionFromState({ scrollIntoView: false });
    this.refresh();
    this.focus();
  }

  private closestVisibleIndexInRange(values: number[], target: number, start: number, end: number): number | null {
    if (values.length === 0) return null;
    const rangeStart = Math.min(start, end);
    const rangeEnd = Math.max(start, end);

    const startIdx = this.lowerBound(values, rangeStart);
    const endExclusive = this.lowerBound(values, rangeEnd + 1);
    if (startIdx >= endExclusive) return null;

    const idx = this.lowerBound(values, target);
    const clampedIdx = Math.min(Math.max(idx, startIdx), endExclusive - 1);

    let bestIdx = clampedIdx;
    let bestValue = values[bestIdx] ?? null;
    if (bestValue == null) return null;
    let bestDist = Math.abs(bestValue - target);

    const belowIdx = idx - 1;
    if (belowIdx >= startIdx && belowIdx < endExclusive) {
      const below = values[belowIdx];
      if (below != null) {
        const dist = Math.abs(below - target);
        if (dist < bestDist) {
          bestDist = dist;
          bestIdx = belowIdx;
          bestValue = below;
        }
      }
    }

    const aboveIdx = idx;
    if (aboveIdx >= startIdx && aboveIdx < endExclusive) {
      const above = values[aboveIdx];
      if (above != null) {
        const dist = Math.abs(above - target);
        if (dist < bestDist) {
          bestDist = dist;
          bestIdx = aboveIdx;
          bestValue = above;
        }
      }
    }

    return bestValue;
  }

  private ensureActiveCellVisible(): void {
    const range = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0] ?? null;
    let { row, col } = this.selection.active;
    let canPreserveSelection = range != null;

    if (this.isRowHidden(row)) {
      const withinRange = (() => {
        if (!range) return null;
        if (!this.sharedGrid) {
          return this.closestVisibleIndexInRange(this.rowIndexByVisual, row, range.startRow, range.endRow);
        }

        // Shared-grid mode intentionally does not build row/col visibility caches (they would be
        // O(maxRows/maxCols) for Excel-scale sheets). Instead, scan within the active selection
        // range to find a visible row, preferring forward (Excel-like).
        const start = Math.max(0, Math.min(range.startRow, range.endRow));
        const end = Math.min(this.limits.maxRows - 1, Math.max(range.startRow, range.endRow));
        for (let r = Math.max(start, row); r <= end; r += 1) {
          if (!this.isRowHidden(r)) return r;
        }
        for (let r = Math.min(end, row - 1); r >= start; r -= 1) {
          if (!this.isRowHidden(r)) return r;
        }
        return null;
      })();
      if (withinRange != null) {
        row = withinRange;
      } else {
        row = this.findNextVisibleRow(row, 1) ?? this.findNextVisibleRow(row, -1) ?? row;
        canPreserveSelection = false;
      }
    }
    if (this.isColHidden(col)) {
      const withinRange = (() => {
        if (!range) return null;
        if (!this.sharedGrid) {
          return this.closestVisibleIndexInRange(this.colIndexByVisual, col, range.startCol, range.endCol);
        }

        const start = Math.max(0, Math.min(range.startCol, range.endCol));
        const end = Math.min(this.limits.maxCols - 1, Math.max(range.startCol, range.endCol));
        for (let c = Math.max(start, col); c <= end; c += 1) {
          if (!this.isColHidden(c)) return c;
        }
        for (let c = Math.min(end, col - 1); c >= start; c -= 1) {
          if (!this.isColHidden(c)) return c;
        }
        return null;
      })();
      if (withinRange != null) {
        col = withinRange;
      } else {
        col = this.findNextVisibleCol(col, 1) ?? this.findNextVisibleCol(col, -1) ?? col;
        canPreserveSelection = false;
      }
    }

    if (row !== this.selection.active.row || col !== this.selection.active.col) {
      if (canPreserveSelection) {
        this.selection = buildSelection(
          {
            ranges: this.selection.ranges,
            active: { row, col },
            anchor: this.selection.anchor,
            activeRangeIndex: this.selection.activeRangeIndex,
          },
          this.limits
        );
      } else {
        // If there are no visible cells inside the current selection range (e.g. a fully-hidden row),
        // fall back to collapsing to the nearest visible cell to keep interaction predictable.
        this.selection = setActiveCell(this.selection, { row, col }, this.limits);
      }
    }
  }

  private findNextVisibleRow(start: number, dir: 1 | -1): number | null {
    let row = start + dir;
    while (row >= 0 && row < this.limits.maxRows) {
      if (!this.isRowHidden(row)) return row;
      row += dir;
    }
    return null;
  }

  private findNextVisibleCol(start: number, dir: 1 | -1): number | null {
    let col = start + dir;
    while (col >= 0 && col < this.limits.maxCols) {
      if (!this.isColHidden(col)) return col;
      col += dir;
    }
    return null;
  }

  private getCellRect(cell: CellCoord) {
    if (this.sharedGrid) {
      const headerRows = this.sharedHeaderRows();
      const headerCols = this.sharedHeaderCols();
      const gridRow = cell.row + headerRows;
      const gridCol = cell.col + headerCols;
      return this.sharedGrid.getCellRect(gridRow, gridCol);
    }

    if (cell.row < 0 || cell.row >= this.limits.maxRows) return null;
    if (cell.col < 0 || cell.col >= this.limits.maxCols) return null;

    const rowDirect = this.rowToVisual.get(cell.row);
    const colDirect = this.colToVisual.get(cell.col);

    // Even when the outline hides rows/cols, downstream overlays still need a
    // stable coordinate space. Hidden rows/cols collapse to zero size and share
    // the same origin as the next visible row/col.
    const visualRow = rowDirect ?? this.lowerBound(this.rowIndexByVisual, cell.row);
    const visualCol = colDirect ?? this.lowerBound(this.colIndexByVisual, cell.col);

    const frozenRows = this.frozenRows;
    const frozenCols = this.frozenCols;

    return {
      x: this.rowHeaderWidth + visualCol * this.cellWidth - (cell.col < frozenCols ? 0 : this.scrollX),
      y: this.colHeaderHeight + visualRow * this.cellHeight - (cell.row < frozenRows ? 0 : this.scrollY),
      width: colDirect == null ? 0 : this.cellWidth,
      height: rowDirect == null ? 0 : this.cellHeight
    };
  }

  private cellFromPoint(pointX: number, pointY: number): CellCoord {
    // Pointer capture means we can receive coordinates outside the grid bounds
    // while the user is dragging a selection. Clamp to the current viewport so
    // we select the edge cell instead of snapping to the end of the sheet.
    const maxX = Math.max(this.rowHeaderWidth, this.width - 1);
    const maxY = Math.max(this.colHeaderHeight, this.height - 1);
    const clampedX = Math.min(Math.max(pointX, this.rowHeaderWidth), maxX);
    const clampedY = Math.min(Math.max(pointY, this.colHeaderHeight), maxY);

    const frozenEdgeX = this.rowHeaderWidth + this.frozenWidth;
    const frozenEdgeY = this.colHeaderHeight + this.frozenHeight;
    const localX = clampedX - this.rowHeaderWidth;
    const localY = clampedY - this.colHeaderHeight;

    const sheetX = clampedX < frozenEdgeX ? localX : this.scrollX + localX;
    const sheetY = clampedY < frozenEdgeY ? localY : this.scrollY + localY;

    const colVisual = Math.floor(sheetX / this.cellWidth);
    const rowVisual = Math.floor(sheetY / this.cellHeight);

    const safeColVisual = Math.max(0, Math.min(this.colIndexByVisual.length - 1, colVisual));
    const safeRowVisual = Math.max(0, Math.min(this.rowIndexByVisual.length - 1, rowVisual));

    const col = this.colIndexByVisual[safeColVisual] ?? 0;
    const row = this.rowIndexByVisual[safeRowVisual] ?? 0;
    return { row, col };
  }

  private computeFillDragTarget(
    source: Range,
    cell: CellCoord
  ): { targetRange: Range; endCell: CellCoord } {
    const clamp = (value: number, min: number, max: number) => Math.max(min, Math.min(max, value));
    const srcTop = source.startRow;
    const srcBottom = source.endRow;
    const srcLeft = source.startCol;
    const srcRight = source.endCol;

    const rowExtension = cell.row < srcTop ? cell.row - srcTop : cell.row > srcBottom ? cell.row - srcBottom : 0;
    const colExtension = cell.col < srcLeft ? cell.col - srcLeft : cell.col > srcRight ? cell.col - srcRight : 0;

    if (rowExtension === 0 && colExtension === 0) {
      const endCell: CellCoord = {
        row: clamp(cell.row, srcTop, srcBottom),
        col: clamp(cell.col, srcLeft, srcRight)
      };
      return { targetRange: source, endCell };
    }

    const axis =
      rowExtension !== 0 && colExtension !== 0
        ? Math.abs(rowExtension) >= Math.abs(colExtension)
          ? "vertical"
          : "horizontal"
        : rowExtension !== 0
          ? "vertical"
          : "horizontal";

    if (axis === "vertical") {
      const targetRange: Range =
        rowExtension > 0
          ? {
              startRow: source.startRow,
              endRow: Math.max(source.endRow, cell.row),
              startCol: source.startCol,
              endCol: source.endCol
            }
          : {
              startRow: Math.min(source.startRow, cell.row),
              endRow: source.endRow,
              startCol: source.startCol,
              endCol: source.endCol
            };

      const endCell: CellCoord = {
        row: clamp(cell.row, targetRange.startRow, targetRange.endRow),
        col: clamp(cell.col, targetRange.startCol, targetRange.endCol)
      };

      return { targetRange, endCell };
    }

    const targetRange: Range =
      colExtension > 0
        ? {
            startRow: source.startRow,
            endRow: source.endRow,
            startCol: source.startCol,
            endCol: Math.max(source.endCol, cell.col)
          }
        : {
            startRow: source.startRow,
            endRow: source.endRow,
            startCol: Math.min(source.startCol, cell.col),
            endCol: source.endCol
          };

    const endCell: CellCoord = {
      row: clamp(cell.row, targetRange.startRow, targetRange.endRow),
      col: clamp(cell.col, targetRange.startCol, targetRange.endCol)
    };

    return { targetRange, endCell };
  }

  private maybeStartDragAutoScroll(): void {
    if (this.disposed) return;
    if (!this.dragState || !this.dragPointerPos) return;

    const margin = 8;
    const { x, y } = this.dragPointerPos;
    const left = this.rowHeaderWidth;
    const top = this.colHeaderHeight;
    const right = this.width;
    const bottom = this.height;

    const outside = x < left - margin || x > right + margin || y < top - margin || y > bottom + margin;
    if (!outside) {
      if (this.dragAutoScrollRaf != null) {
        if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(this.dragAutoScrollRaf);
        else globalThis.clearTimeout(this.dragAutoScrollRaf);
        this.dragAutoScrollRaf = null;
      }
      return;
    }

    if (this.dragAutoScrollRaf != null) return;

    const schedule =
      typeof requestAnimationFrame === "function"
        ? requestAnimationFrame
        : (cb: FrameRequestCallback) =>
            globalThis.setTimeout(() => cb(typeof performance !== "undefined" ? performance.now() : Date.now()), 16);

    const tick = () => {
      this.dragAutoScrollRaf = null;
      if (!this.dragState || !this.dragPointerPos) return;

      const px = this.dragPointerPos.x;
      const py = this.dragPointerPos.y;

      let deltaX = 0;
      let deltaY = 0;

      if (px < left - margin) {
        deltaX = -Math.min(this.cellWidth, left - margin - px);
      } else if (px > right + margin) {
        deltaX = Math.min(this.cellWidth, px - (right + margin));
      }

      if (py < top - margin) {
        deltaY = -Math.min(this.cellHeight, top - margin - py);
      } else if (py > bottom + margin) {
        deltaY = Math.min(this.cellHeight, py - (bottom + margin));
      }

      if (deltaX === 0 && deltaY === 0) return;

      const didScroll = this.setScrollInternal(this.scrollX + deltaX, this.scrollY + deltaY);
      if (!didScroll) return;

      const cell = this.cellFromPoint(px, py);
      if (this.dragState.mode === "fill") {
        const state = this.dragState;
        const source = state.sourceRange;
        const { targetRange, endCell } = this.computeFillDragTarget(source, cell);
        state.targetRange = targetRange;
        state.endCell = endCell;
        this.fillPreviewRange =
          targetRange.startRow === source.startRow &&
          targetRange.endRow === source.endRow &&
          targetRange.startCol === source.startCol &&
          targetRange.endCol === source.endCol
            ? null
            : targetRange;
      } else {
        this.selection = extendSelectionToCell(this.selection, cell, this.limits);

        if (this.dragState.mode === "formula" && this.formulaBar) {
          const r = this.selection.ranges[0];
          if (r) {
            const rangeSheetId =
              this.formulaEditCell && this.formulaEditCell.sheetId !== this.sheetId ? this.sheetId : undefined;
            const rangeSheetName = rangeSheetId ? this.resolveSheetDisplayNameById(rangeSheetId) : undefined;
            this.formulaBar.updateRangeSelection(
              {
                start: { row: r.startRow, col: r.startCol },
                end: { row: r.endRow, col: r.endCol }
              },
              rangeSheetName
            );
          }
        }
      }

      // Repaint immediately (we're already in rAF / a timer tick).
      this.renderGrid();
      this.renderCharts(false);
      this.renderReferencePreview();
      this.renderSelection();
      this.updateStatus();

      this.dragAutoScrollRaf = schedule(tick) as unknown as number;
    };

    this.dragAutoScrollRaf = schedule(tick) as unknown as number;
  }

  private onPointerDown(e: PointerEvent): void {
    const editorWasOpen = this.editor.isOpen();

    const rect = this.root.getBoundingClientRect();
    this.rootLeft = rect.left;
    this.rootTop = rect.top;
    this.rootPosLastMeasuredAtMs =
      typeof performance !== "undefined" && typeof performance.now === "function" ? performance.now() : Date.now();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    const primaryButton = e.pointerType !== "mouse" || e.button === 0;

    const formulaEditing = this.formulaBar?.isFormulaEditing() === true;
    if (!formulaEditing) {
      // Drawing hit testing must happen before cell-selection logic so clicks on
      // overlaid objects (charts/images/shapes) behave like Excel.
      const drawingViewport = this.getDrawingInteractionViewport();
      const drawings = this.listDrawingObjectsForSheet();

      // Allow grabbing a resize handle for the current drawing selection even when the
      // pointer is slightly outside the object's bounds (handles extend beyond the
      // selection outline).
      if (primaryButton && this.selectedDrawingId != null) {
        const selected = drawings.find((obj) => obj.id === this.selectedDrawingId) ?? null;
        if (selected) {
          const headerOffsetX = Number.isFinite(drawingViewport.headerOffsetX)
            ? Math.max(0, drawingViewport.headerOffsetX!)
            : 0;
          const headerOffsetY = Number.isFinite(drawingViewport.headerOffsetY)
            ? Math.max(0, drawingViewport.headerOffsetY!)
            : 0;
          if (x >= headerOffsetX && y >= headerOffsetY) {
            const selectedBounds = drawingObjectToViewportRect(selected, drawingViewport, this.drawingGeom);
            const handle = hitTestResizeHandle(selectedBounds, x, y, selected.transform);
            if (handle) {
              if (editorWasOpen) {
                this.editor.commit("command");
              }
              e.preventDefault();
              this.renderSelection();
              this.focus();

              const scroll = effectiveScrollForAnchor(selected.anchor, drawingViewport);
              const startSheetX = x - headerOffsetX + scroll.scrollX;
              const startSheetY = y - headerOffsetY + scroll.scrollY;
              this.drawingGesture = {
                pointerId: e.pointerId,
                mode: "resize",
                objectId: selected.id,
                handle,
                startSheetX,
                startSheetY,
                startAnchor: selected.anchor,
                startWidthPx: selectedBounds.width,
                startHeightPx: selectedBounds.height,
                transform: selected.transform,
                aspectRatio:
                  selected.kind.type === "image" && selectedBounds.width > 0 && selectedBounds.height > 0
                    ? selectedBounds.width / selectedBounds.height
                    : null,
              };
              try {
                this.root.setPointerCapture(e.pointerId);
              } catch {
                // Best-effort; some environments (tests/jsdom) may not implement pointer capture.
              }
              return;
            }
          }
        }
      }

      const hitIndex = this.getDrawingHitTestIndex(drawings);
      const hit = hitTestDrawings(hitIndex, drawingViewport, x, y, this.drawingGeom);
      if (hit) {
        if (editorWasOpen) {
          this.editor.commit("command");
        }
        const prevSelected = this.selectedDrawingId;
        this.selectedDrawingId = hit.object.id;
        if (prevSelected !== hit.object.id) {
          this.dispatchDrawingSelectionChanged();
        }
        this.renderSelection();
        this.focus();

        // Begin drag/resize gesture for primary-button interactions.
        if (primaryButton) {
          e.preventDefault();
          const handle = hitTestResizeHandle(hit.bounds, x, y, hit.object.transform);
          const scroll = effectiveScrollForAnchor(hit.object.anchor, drawingViewport);
          const headerOffsetX = Number.isFinite(drawingViewport.headerOffsetX)
            ? Math.max(0, drawingViewport.headerOffsetX!)
            : 0;
          const headerOffsetY = Number.isFinite(drawingViewport.headerOffsetY)
            ? Math.max(0, drawingViewport.headerOffsetY!)
            : 0;
          const startSheetX = x - headerOffsetX + scroll.scrollX;
          const startSheetY = y - headerOffsetY + scroll.scrollY;
          this.drawingGesture = handle
            ? {
                pointerId: e.pointerId,
                mode: "resize",
                objectId: hit.object.id,
                handle,
                startSheetX,
                startSheetY,
                startAnchor: hit.object.anchor,
                startWidthPx: hit.bounds.width,
                startHeightPx: hit.bounds.height,
                transform: hit.object.transform,
                aspectRatio:
                  hit.object.kind.type === "image" && hit.bounds.width > 0 && hit.bounds.height > 0
                    ? hit.bounds.width / hit.bounds.height
                    : null,
              }
            : {
                pointerId: e.pointerId,
                mode: "drag",
                objectId: hit.object.id,
                startSheetX,
                startSheetY,
                startAnchor: hit.object.anchor,
              };
          try {
            this.root.setPointerCapture(e.pointerId);
          } catch {
            // Best-effort; some environments (tests/jsdom) may not implement pointer capture.
          }
        }

        return;
      }
    }

    if (editorWasOpen) return;

    // Clicking outside any drawing clears drawing selection.
    if (primaryButton && this.selectedDrawingId != null) {
      this.selectedDrawingId = null;
      this.dispatchDrawingSelectionChanged();
    }

    // Right/middle clicks should not start drag selection, but we still want right-click to
    // apply to the cell under the cursor when the click happens outside the current selection.
    // (Excel keeps selection when right-clicking inside it, but moves selection when right-clicking
    // outside.)
    if (e.pointerType === "mouse" && e.button !== 0) {
      if (x >= this.rowHeaderWidth && y >= this.colHeaderHeight) {
        const cell = this.cellFromPoint(x, y);
        const inSelection = this.selection.ranges.some(
          (range) =>
            cell.row >= range.startRow &&
            cell.row <= range.endRow &&
            cell.col >= range.startCol &&
            cell.col <= range.endCol
        );
        if (!inSelection) {
          this.selection = setActiveCell(this.selection, cell, this.limits);
          this.renderSelection();
          this.updateStatus();
        }
      }
      this.focus();
      return;
    }

    const primary = e.ctrlKey || e.metaKey;

    // Top-left corner selects the entire sheet.
    if (x < this.rowHeaderWidth && y < this.colHeaderHeight) {
      this.selection = selectAll(this.limits);
      this.renderSelection();
      this.updateStatus();
      this.focus();
      return;
    }

    // Column header selects entire column.
    if (y < this.colHeaderHeight && x >= this.rowHeaderWidth) {
      const frozenEdgeX = this.rowHeaderWidth + this.frozenWidth;
      const localX = x - this.rowHeaderWidth;
      const sheetX = x < frozenEdgeX ? localX : this.scrollX + localX;
      const visualCol = Math.floor(sheetX / this.cellWidth);
      const safeVisualCol = Math.max(0, Math.min(this.colIndexByVisual.length - 1, visualCol));
      const col = this.colIndexByVisual[safeVisualCol] ?? 0;
      this.selection = selectColumns(this.selection, col, col, { additive: primary }, this.limits);
      this.renderSelection();
      this.updateStatus();
      this.focus();
      return;
    }

    // Row header selects entire row.
    if (x < this.rowHeaderWidth && y >= this.colHeaderHeight) {
      const frozenEdgeY = this.colHeaderHeight + this.frozenHeight;
      const localY = y - this.colHeaderHeight;
      const sheetY = y < frozenEdgeY ? localY : this.scrollY + localY;
      const visualRow = Math.floor(sheetY / this.cellHeight);
      const safeVisualRow = Math.max(0, Math.min(this.rowIndexByVisual.length - 1, visualRow));
      const row = this.rowIndexByVisual[safeVisualRow] ?? 0;
      this.selection = selectRows(this.selection, row, row, { additive: primary }, this.limits);
      this.renderSelection();
      this.updateStatus();
      this.focus();
      return;
    }

    const cell = this.cellFromPoint(x, y);

    if (this.formulaBar?.isFormulaEditing()) {
      e.preventDefault();
      this.dragState = { pointerId: e.pointerId, mode: "formula" };
      this.dragPointerPos = { x, y };
      this.root.setPointerCapture(e.pointerId);
      this.selection = setActiveCell(this.selection, cell, this.limits);
      this.renderSelection();
      this.updateStatus();
      const rangeSheetId =
        this.formulaEditCell && this.formulaEditCell.sheetId !== this.sheetId ? this.sheetId : undefined;
      const rangeSheetName = rangeSheetId ? this.resolveSheetDisplayNameById(rangeSheetId) : undefined;
      this.formulaBar.beginRangeSelection(
        {
          start: { row: cell.row, col: cell.col },
          end: { row: cell.row, col: cell.col }
        },
        rangeSheetName
      );
      return;
    }

    const fillHandle = this.getFillHandleRect();
    if (
      !this.isReadOnly() &&
      fillHandle &&
      x >= fillHandle.x &&
      x <= fillHandle.x + fillHandle.width &&
      y >= fillHandle.y &&
      y <= fillHandle.y + fillHandle.height
    ) {
      e.preventDefault();
      const sourceRange = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0];
      if (sourceRange) {
        const fillMode: FillHandleMode = e.altKey ? "formulas" : e.ctrlKey || e.metaKey ? "copy" : "series";
        this.dragState = {
          pointerId: e.pointerId,
          mode: "fill",
          sourceRange,
          targetRange: sourceRange,
          endCell: { row: sourceRange.endRow, col: sourceRange.endCol },
          fillMode,
          activeRangeIndex: this.selection.activeRangeIndex
        };
        this.dragPointerPos = { x, y };
        this.fillPreviewRange = null;
        this.root.setPointerCapture(e.pointerId);
        this.focus();
        return;
      }
    }

    // Ctrl/Cmd+click on a URL-like cell value should open it externally instead
    // of being treated as an additive selection gesture.
    if (primary && e.pointerType === "mouse" && e.button === 0) {
      const state = this.document.getCell(this.sheetId, cell) as { value: unknown; formula: string | null };
      const renderedText = (() => {
        if (!state) return "";
        if (state.formula != null) {
          if (this.showFormulas) return state.formula;
          const computed = this.getCellComputedValue(cell);
          return computed == null ? "" : String(computed);
        }
        if (isRichTextValue(state.value)) return state.value.text;
        if (state.value != null) return String(state.value);
        return "";
      })();

      if (typeof renderedText === "string" && looksLikeExternalHyperlink(renderedText)) {
        // Match normal click behavior (make the clicked cell active) while still
        // allowing the OS browser open behavior behind Ctrl/Cmd.
        this.selection = setActiveCell(this.selection, cell, this.limits);
        this.renderSelection();
        this.updateStatus();
        this.focus();

        void openExternalHyperlink(renderedText.trim(), {
          shellOpen,
          confirmUntrustedProtocol: async (message) => nativeDialogs.confirm(message),
        }).catch(() => {
          // Best-effort: link opening should not crash grid interaction.
        });
        return;
      }
    }

    this.dragState = { pointerId: e.pointerId, mode: "normal" };
    this.dragPointerPos = { x, y };
    this.root.setPointerCapture(e.pointerId);
    if (e.shiftKey) {
      this.selection = extendSelectionToCell(this.selection, cell, this.limits);
    } else if (primary) {
      this.selection = addCellToSelection(this.selection, cell, this.limits);
    } else {
      this.selection = setActiveCell(this.selection, cell, this.limits);
    }

    this.renderSelection();
    this.updateStatus();
    this.focus();
  }

  private onPointerMove(e: PointerEvent): void {
    if (this.scrollbarDrag) {
      this.onScrollbarThumbPointerMove(e);
      return;
    }

    if (this.drawingGesture) {
      if (e.pointerId !== this.drawingGesture.pointerId) return;
      if (this.editor.isOpen()) return;

      const x = e.clientX - this.rootLeft;
      const y = e.clientY - this.rootTop;

      const viewport = this.getDrawingInteractionViewport();
      const zoom = Number.isFinite(viewport.zoom) && (viewport.zoom as number) > 0 ? (viewport.zoom as number) : 1;
      const gesture = this.drawingGesture;
      const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
      const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;
      const frozenBoundaryX = Number.isFinite(viewport.frozenWidthPx) ? Math.max(headerOffsetX, viewport.frozenWidthPx!) : headerOffsetX;
      const frozenBoundaryY = Number.isFinite(viewport.frozenHeightPx) ? Math.max(headerOffsetY, viewport.frozenHeightPx!) : headerOffsetY;
      const inHeader = x < headerOffsetX || y < headerOffsetY;
      const pointInFrozenCols = !inHeader && x < frozenBoundaryX;
      const pointInFrozenRows = !inHeader && y < frozenBoundaryY;
      // Absolute anchors always scroll; oneCell/twoCell anchors use the frozen pane under the pointer.
      const alwaysScroll = gesture.startAnchor.type === "absolute";
      const scrollX = alwaysScroll ? viewport.scrollX : pointInFrozenCols ? 0 : viewport.scrollX;
      const scrollY = alwaysScroll ? viewport.scrollY : pointInFrozenRows ? 0 : viewport.scrollY;
      const sheetX = x - headerOffsetX + scrollX;
      const sheetY = y - headerOffsetY + scrollY;

      let dxPx = sheetX - gesture.startSheetX;
      let dyPx = sheetY - gesture.startSheetY;

      const nextAnchor = (() => {
        if (gesture.mode !== "resize") {
          return shiftAnchor(gesture.startAnchor, dxPx, dyPx, this.drawingGeom, zoom);
        }

        const transform = gesture.transform;
        const hasNonIdentityTransform = !!(
          transform &&
          (transform.rotationDeg !== 0 || transform.flipH || transform.flipV)
        );

        if (e.shiftKey && gesture.aspectRatio != null) {
          if (hasNonIdentityTransform) {
            const local = inverseTransformVector(dxPx, dyPx, transform!);
            const lockedLocal = lockAspectRatioResize({
              handle: gesture.handle,
              dx: local.x,
              dy: local.y,
              startWidthPx: gesture.startWidthPx,
              startHeightPx: gesture.startHeightPx,
              aspectRatio: gesture.aspectRatio,
              minSizePx: 8,
            });
            const world = applyTransformVector(lockedLocal.dx, lockedLocal.dy, transform!);
            dxPx = world.x;
            dyPx = world.y;
          } else {
            const locked = lockAspectRatioResize({
              handle: gesture.handle,
              dx: dxPx,
              dy: dyPx,
              startWidthPx: gesture.startWidthPx,
              startHeightPx: gesture.startHeightPx,
              aspectRatio: gesture.aspectRatio,
              minSizePx: 8,
            });
            dxPx = locked.dx;
            dyPx = locked.dy;
          }
        }

        return resizeAnchor(gesture.startAnchor, gesture.handle, dxPx, dyPx, this.drawingGeom, transform, zoom);
      })();

      const objects = this.listDrawingObjectsForSheet();
      const nextObjects = objects.map((obj) => (obj.id === gesture.objectId ? { ...obj, anchor: nextAnchor } : obj));
      const doc = this.document as any;
      const drawingsGetter = typeof doc.getSheetDrawings === "function" ? doc.getSheetDrawings : null;
      this.drawingObjectsCache = { sheetId: this.sheetId, objects: nextObjects, source: drawingsGetter };
      this.renderDrawings();

      this.renderSelection();
      return;
    }

    const target = e.target as HTMLElement | null;
    const useOffsetCoords =
      target === this.root ||
      target === this.selectionCanvas ||
      target === this.gridCanvas ||
      target === this.referenceCanvas ||
      target === this.auditingCanvas ||
      target === this.presenceCanvas;

    if (!this.dragState) {
      if (target && !useOffsetCoords) {
        if (
          this.vScrollbarTrack.contains(target) ||
          this.hScrollbarTrack.contains(target) ||
          this.outlineLayer.contains(target)
        ) {
          this.hideCommentTooltip();
          this.root.style.cursor = "";
          return;
        }
      }

      // When the pointermove target is the selection canvas (or another full-viewport canvas
      // overlay), we can use `offsetX/Y` which are already in the correct coordinate space.
      // Only refresh cached root position when we need to fall back to client-relative coords.
      if (!useOffsetCoords) {
        this.maybeRefreshRootPosition();
      }
    }

    // During drag selection (pointer capture), use client-relative coordinates so that
    // moving outside the grid continues to produce out-of-bounds values. This is
    // required for behaviors like drag auto-scroll.
    const x = this.dragState ? e.clientX - this.rootLeft : useOffsetCoords ? e.offsetX : e.clientX - this.rootLeft;
    const y = this.dragState ? e.clientY - this.rootTop : useOffsetCoords ? e.offsetY : e.clientY - this.rootTop;
    if (this.dragPointerPos) {
      this.dragPointerPos.x = x;
      this.dragPointerPos.y = y;
    }

    // Hover cursor feedback is only meaningful when no buttons are pressed. When a drawing is being
    // dragged/resized via a separate overlay/controller, SpreadsheetApp still receives bubbled
    // pointermoves; keep the cursor fixed by skipping hover hit testing while buttons are down.
    if (!this.dragState && e.buttons) {
      this.hideCommentTooltip();
      return;
    }

    if (this.dragState) {
      if (e.pointerId !== this.dragState.pointerId) return;
      if (this.editor.isOpen()) return;
      this.hideCommentTooltip();
      const cell = this.cellFromPoint(x, y);

      if (this.dragState.mode === "fill") {
        const state = this.dragState;
        const source = state.sourceRange;
        const { targetRange, endCell } = this.computeFillDragTarget(source, cell);
        state.targetRange = targetRange;
        state.endCell = endCell;
        this.fillPreviewRange =
          targetRange.startRow === source.startRow &&
          targetRange.endRow === source.endRow &&
          targetRange.startCol === source.startCol &&
          targetRange.endCol === source.endCol
            ? null
            : targetRange;
        this.renderSelection();
        this.maybeStartDragAutoScroll();
        return;
      }

      this.selection = extendSelectionToCell(this.selection, cell, this.limits);
      this.renderSelection();
      this.updateStatus();

      if (this.dragState.mode === "formula" && this.formulaBar) {
        const r = this.selection.ranges[0];
        const rangeSheetId =
          this.formulaEditCell && this.formulaEditCell.sheetId !== this.sheetId ? this.sheetId : undefined;
        const rangeSheetName = rangeSheetId ? this.resolveSheetDisplayNameById(rangeSheetId) : undefined;
        this.formulaBar.updateRangeSelection(
          {
            start: { row: r.startRow, col: r.startCol },
            end: { row: r.endRow, col: r.endCol }
          },
          rangeSheetName
        );
      }

      this.maybeStartDragAutoScroll();
      return;
    }

    const fillHandle = this.getFillHandleRect();
    const overFillHandle =
      !this.isReadOnly() &&
      fillHandle &&
      x >= fillHandle.x &&
      x <= fillHandle.x + fillHandle.width &&
      y >= fillHandle.y &&
      y <= fillHandle.y + fillHandle.height;
    const chartCursor = this.chartCursorAtPoint(x, y);
    const drawingCursor = this.drawingCursorAtPoint(x, y);
    const nextCursor = chartCursor ?? drawingCursor ?? (overFillHandle ? "crosshair" : "");
    if (this.root.style.cursor !== nextCursor) {
      this.root.style.cursor = nextCursor;
    }

    if (chartCursor) {
      // Charts sit above cell content; suppress comment tooltips while hovering the chart area.
      this.hideCommentTooltip();
      return;
    }

    if (this.commentsPanelVisible) {
      // Don't show tooltips while the panel is open; it obscures the grid anyway.
      this.hideCommentTooltip();
      return;
    }

    if (x < 0 || y < 0 || x > this.width || y > this.height) {
      this.hideCommentTooltip();
      this.root.style.cursor = "";
      return;
    }

    if (x < this.rowHeaderWidth || y < this.colHeaderHeight) {
      this.hideCommentTooltip();
      this.root.style.cursor = "";
      return;
    }

    if (this.commentMetaByCoord.size === 0) {
      this.hideCommentTooltip();
      return;
    }

    const cell = this.cellFromPoint(x, y);
    const metaKey = cell.row * COMMENT_COORD_COL_STRIDE + cell.col;
    const preview = this.commentPreviewByCoord.get(metaKey);
    if (preview === undefined) {
      this.hideCommentTooltip();
      return;
    }

    if (
      this.lastHoveredCommentCellKey === metaKey &&
      this.lastHoveredCommentIndexVersion === this.commentIndexVersion &&
      this.commentTooltipVisible
    ) {
      // Keep tooltip pinned to the cursor without re-setting text content on every move.
      this.commentTooltip.style.setProperty("--comment-tooltip-x", `${x + 12}px`);
      this.commentTooltip.style.setProperty("--comment-tooltip-y", `${y + 12}px`);
      return;
    }

    this.lastHoveredCommentCellKey = metaKey;
    this.lastHoveredCommentIndexVersion = this.commentIndexVersion;
    this.commentTooltip.textContent = preview;
    this.commentTooltip.style.setProperty("--comment-tooltip-x", `${x + 12}px`);
    this.commentTooltip.style.setProperty("--comment-tooltip-y", `${y + 12}px`);
    this.commentTooltipVisible = true;
    this.commentTooltip.classList.add("comment-tooltip--visible");
  }

  private onPointerUp(e: PointerEvent): void {
    if (this.scrollbarDrag) {
      this.onScrollbarThumbPointerUp(e);
      return;
    }

    if (this.drawingGesture) {
      if (e.pointerId !== this.drawingGesture.pointerId) return;
      const gesture = this.drawingGesture;

      const x = e.clientX - this.rootLeft;
      const y = e.clientY - this.rootTop;

      const viewport = this.getDrawingInteractionViewport();
      const zoom = Number.isFinite(viewport.zoom) && (viewport.zoom as number) > 0 ? (viewport.zoom as number) : 1;
      const scroll = effectiveScrollForAnchor(gesture.startAnchor, viewport);
      const headerOffsetX = Number.isFinite(viewport.headerOffsetX) ? Math.max(0, viewport.headerOffsetX!) : 0;
      const headerOffsetY = Number.isFinite(viewport.headerOffsetY) ? Math.max(0, viewport.headerOffsetY!) : 0;
      const sheetX = x - headerOffsetX + scroll.scrollX;
      const sheetY = y - headerOffsetY + scroll.scrollY;

      let dxPx = sheetX - gesture.startSheetX;
      let dyPx = sheetY - gesture.startSheetY;

      const nextAnchor = (() => {
        if (gesture.mode !== "resize") {
          return shiftAnchor(gesture.startAnchor, dxPx, dyPx, this.drawingGeom, zoom);
        }

        const transform = gesture.transform;
        const hasNonIdentityTransform = !!(
          transform &&
          (transform.rotationDeg !== 0 || transform.flipH || transform.flipV)
        );

        if (e.shiftKey && gesture.aspectRatio != null) {
          if (hasNonIdentityTransform) {
            const local = inverseTransformVector(dxPx, dyPx, transform!);
            const lockedLocal = lockAspectRatioResize({
              handle: gesture.handle,
              dx: local.x,
              dy: local.y,
              startWidthPx: gesture.startWidthPx,
              startHeightPx: gesture.startHeightPx,
              aspectRatio: gesture.aspectRatio,
              minSizePx: 8,
            });
            const world = applyTransformVector(lockedLocal.dx, lockedLocal.dy, transform!);
            dxPx = world.x;
            dyPx = world.y;
          } else {
            const locked = lockAspectRatioResize({
              handle: gesture.handle,
              dx: dxPx,
              dy: dyPx,
              startWidthPx: gesture.startWidthPx,
              startHeightPx: gesture.startHeightPx,
              aspectRatio: gesture.aspectRatio,
              minSizePx: 8,
            });
            dxPx = locked.dx;
            dyPx = locked.dy;
          }
        }

        return resizeAnchor(gesture.startAnchor, gesture.handle, dxPx, dyPx, this.drawingGeom, transform, zoom);
      })();

      const objects = this.listDrawingObjectsForSheet();
      const nextObjects = objects.map((obj) => (obj.id === gesture.objectId ? { ...obj, anchor: nextAnchor } : obj));
      const doc = this.document as any;
      const drawingsGetter = typeof doc.getSheetDrawings === "function" ? doc.getSheetDrawings : null;
      this.drawingObjectsCache = { sheetId: this.sheetId, objects: nextObjects, source: drawingsGetter };
      this.renderDrawings();

      this.drawingGesture = null;
      try {
        this.root.releasePointerCapture(e.pointerId);
      } catch {
        // Best-effort; some environments (tests/jsdom) may not implement pointer capture.
      }
      // Ensure selection handles reflect the final position.
      this.renderSelection();
      return;
    }

    if (!this.dragState) return;
    if (e.pointerId !== this.dragState.pointerId) return;
    const state = this.dragState;
    this.dragState = null;
    this.dragPointerPos = null;
    if (this.dragAutoScrollRaf != null) {
      if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(this.dragAutoScrollRaf);
      else globalThis.clearTimeout(this.dragAutoScrollRaf);
    }
    this.dragAutoScrollRaf = null;

    if (state.mode === "fill") {
      this.fillPreviewRange = null;
      const shouldCommit = e.type === "pointerup";
      const { sourceRange, targetRange, endCell, fillMode, activeRangeIndex } = state;
      const changed =
        targetRange.startRow !== sourceRange.startRow ||
        targetRange.endRow !== sourceRange.endRow ||
        targetRange.startCol !== sourceRange.startCol ||
        targetRange.endCol !== sourceRange.endCol;

      if (changed && shouldCommit) {
        const applied = this.applyFill(sourceRange, targetRange, fillMode);
        if (!applied) {
          // Clear any preview overlay and keep selection stable.
          this.renderSelection();
          this.updateStatus();
          this.focus();
          return;
        }

        const existingRanges = this.selection.ranges;
        const nextActiveIndex = Math.max(0, Math.min(existingRanges.length - 1, activeRangeIndex));
        const updatedRanges = existingRanges.length === 0 ? [targetRange] : [...existingRanges];
        updatedRanges[nextActiveIndex] = targetRange;

        this.selection = buildSelection(
          {
            ranges: updatedRanges,
            active: endCell,
            anchor: endCell,
            activeRangeIndex: nextActiveIndex
          },
          this.limits
        );

        this.refresh();
        this.focus();
      } else {
        // Clear any preview overlay.
        this.renderSelection();
      }

      if (this.auditingNeedsUpdateAfterDrag) {
        this.auditingNeedsUpdateAfterDrag = false;
        this.scheduleAuditingUpdate();
      }
      return;
    }

    if (this.auditingNeedsUpdateAfterDrag) {
      this.auditingNeedsUpdateAfterDrag = false;
      this.scheduleAuditingUpdate();
    }

    if (state.mode === "formula" && this.formulaBar) {
      this.formulaBar.endRangeSelection();
      // Restore focus to the formula bar without clearing its insertion state mid-drag.
      this.formulaBar.focus();
    }
  }

  private applyFill(sourceRange: Range, targetRange: Range, mode: FillHandleMode): boolean {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return false;
    }
    if (this.isEditing()) return false;
    const toFillRange = (range: Range): FillEngineRange => ({
      startRow: range.startRow,
      endRow: range.endRow + 1,
      startCol: range.startCol,
      endCol: range.endCol + 1
    });

    const source = toFillRange(sourceRange);
    const union = toFillRange(targetRange);

    const deltaRange: FillEngineRange | null = (() => {
      const sameCols = source.startCol === union.startCol && source.endCol === union.endCol;
      const sameRows = source.startRow === union.startRow && source.endRow === union.endRow;

      if (sameCols) {
        if (union.endRow > source.endRow) {
          return { startRow: source.endRow, endRow: union.endRow, startCol: source.startCol, endCol: source.endCol };
        }
        if (union.startRow < source.startRow) {
          return { startRow: union.startRow, endRow: source.startRow, startCol: source.startCol, endCol: source.endCol };
        }
      }

      if (sameRows) {
        if (union.endCol > source.endCol) {
          return { startRow: source.startRow, endRow: source.endRow, startCol: source.endCol, endCol: union.endCol };
        }
        if (union.startCol < source.startCol) {
          return { startRow: source.startRow, endRow: source.endRow, startCol: union.startCol, endCol: source.startCol };
        }
      }

      return null;
    })();

    if (!deltaRange) return false;

    const sourceCells = (source.endRow - source.startRow) * (source.endCol - source.startCol);
    const targetCells = (deltaRange.endRow - deltaRange.startRow) * (deltaRange.endCol - deltaRange.startCol);
    if (sourceCells > MAX_FILL_CELLS || targetCells > MAX_FILL_CELLS) {
      try {
        showToast(
          `Fill range too large (>${MAX_FILL_CELLS.toLocaleString()} cells). Select fewer cells and try again.`,
          "warning"
        );
      } catch {
        // `showToast` requires a #toast-root; unit tests don't always include it.
      }
      return false;
    }

    const fillCoordScratch = { row: 0, col: 0 };
    const getCellComputedValue = (row: number, col: number) => {
      fillCoordScratch.row = row;
      fillCoordScratch.col = col;
      return this.getCellComputedValue(fillCoordScratch) as any;
    };

    const wasm = this.wasmEngine;
    if (wasm && mode !== "copy") {
      const task = applyFillCommitToDocumentControllerWithFormulaRewrite({
        document: this.document,
        sheetId: this.sheetId,
        sourceRange: source,
        targetRange: deltaRange,
        mode,
        getCellComputedValue,
        rewriteFormulasForCopyDelta: (requests) => wasm.rewriteFormulasForCopyDelta(requests),
        label: "Fill",
      })
        .catch(() => {
          applyFillCommitToDocumentController({
            document: this.document,
            sheetId: this.sheetId,
            sourceRange: source,
            targetRange: deltaRange,
            mode,
            getCellComputedValue,
          });
        })
        .finally(() => {
          this.refresh();
          this.focus();
        });
      this.idle.track(task);
      return true;
    }

    applyFillCommitToDocumentController({
      document: this.document,
      sheetId: this.sheetId,
      sourceRange: source,
      targetRange: deltaRange,
      mode,
      getCellComputedValue,
    });
    return true;
  }

  private applyFillShortcut(direction: "down" | "right" | "up" | "left", mode: Exclude<FillHandleMode, "copy">): void {
    const clampInt = (value: number, min: number, max: number): number => {
      const n = Math.trunc(value);
      return Math.min(max, Math.max(min, n));
    };

    const maxRowInclusive = Math.max(0, this.limits.maxRows - 1);
    const maxColInclusive = Math.max(0, this.limits.maxCols - 1);
    const maxRowExclusive = Math.max(0, this.limits.maxRows);
    const maxColExclusive = Math.max(0, this.limits.maxCols);

    const operations: Array<{ sourceRange: FillEngineRange; targetRange: FillEngineRange }> = [];

    for (const range of this.selection.ranges) {
      const startRow = clampInt(Math.min(range.startRow, range.endRow), 0, maxRowInclusive);
      const endRow = clampInt(Math.max(range.startRow, range.endRow), 0, maxRowInclusive);
      const startCol = clampInt(Math.min(range.startCol, range.endCol), 0, maxColInclusive);
      const endCol = clampInt(Math.max(range.startCol, range.endCol), 0, maxColInclusive);

      const height = Math.max(0, endRow - startRow + 1);
      const width = Math.max(0, endCol - startCol + 1);

      const seedSpan = mode === "series" ? 2 : 1;

      if (direction === "down" || direction === "up") {
        const seedRows = Math.min(seedSpan, height);
        if (height <= seedRows) continue;

        const sourceStartRow = direction === "down" ? startRow : endRow - seedRows + 1;
        const sourceEndRow = direction === "down" ? sourceStartRow + seedRows : endRow + 1;
        const targetStartRow = direction === "down" ? sourceStartRow + seedRows : startRow;
        const targetEndRow = direction === "down" ? endRow + 1 : sourceStartRow;

        const sourceRange: FillEngineRange = {
          startRow: clampInt(sourceStartRow, 0, maxRowInclusive),
          endRow: clampInt(sourceEndRow, 0, maxRowExclusive),
          startCol,
          endCol: Math.min(endCol + 1, maxColExclusive)
        };
        const targetRange: FillEngineRange = {
          startRow: clampInt(targetStartRow, 0, maxRowExclusive),
          endRow: clampInt(targetEndRow, 0, maxRowExclusive),
          startCol,
          endCol: Math.min(endCol + 1, maxColExclusive)
        };

        if (targetRange.endRow <= targetRange.startRow) continue;
        if (sourceRange.endRow <= sourceRange.startRow) continue;
        if (sourceRange.endCol <= sourceRange.startCol) continue;
        operations.push({ sourceRange, targetRange });
        continue;
      }
      const seedCols = Math.min(seedSpan, width);
      if (width <= seedCols) continue;

      const sourceStartCol = direction === "right" ? startCol : endCol - seedCols + 1;
      const sourceEndCol = direction === "right" ? sourceStartCol + seedCols : endCol + 1;
      const targetStartCol = direction === "right" ? sourceStartCol + seedCols : startCol;
      const targetEndCol = direction === "right" ? endCol + 1 : sourceStartCol;

      const sourceRange: FillEngineRange = {
        startRow,
        endRow: Math.min(endRow + 1, maxRowExclusive),
        startCol: clampInt(sourceStartCol, 0, maxColInclusive),
        endCol: clampInt(sourceEndCol, 0, maxColExclusive)
      };
      const targetRange: FillEngineRange = {
        startRow,
        endRow: Math.min(endRow + 1, maxRowExclusive),
        startCol: clampInt(targetStartCol, 0, maxColExclusive),
        endCol: clampInt(targetEndCol, 0, maxColExclusive)
      };

      if (targetRange.endCol <= targetRange.startCol) continue;
      if (sourceRange.endCol <= sourceRange.startCol) continue;
      if (sourceRange.endRow <= sourceRange.startRow) continue;
      operations.push({ sourceRange, targetRange });
    }

    if (operations.length === 0) return;

    let totalTargetCells = 0;
    for (const op of operations) {
      const rows = Math.max(0, op.targetRange.endRow - op.targetRange.startRow);
      const cols = Math.max(0, op.targetRange.endCol - op.targetRange.startCol);
      totalTargetCells += rows * cols;
      if (totalTargetCells > MAX_FILL_CELLS) break;
    }

    if (totalTargetCells > MAX_FILL_CELLS) {
      try {
        showToast(
          `Selection too large to fill (>${MAX_FILL_CELLS.toLocaleString()} cells). Select fewer cells and try again.`,
          "warning"
        );
      } catch {
        // `showToast` requires a #toast-root; unit tests don't always include it.
      }
      return;
    }

    const label = (() => {
      if (mode === "series") {
        switch (direction) {
          case "down":
            return "Series Down";
          case "right":
            return "Series Right";
          case "up":
            return "Series Up";
          case "left":
            return "Series Left";
        }
      }

      switch (direction) {
        case "down":
          return t("command.edit.fillDown");
        case "right":
          return t("command.edit.fillRight");
        case "up":
          return t("command.edit.fillUp");
        case "left":
          return t("command.edit.fillLeft");
      }
    })();

    const wasm = this.wasmEngine;

    // Explicit batch so multi-range selections become a single undo step.
    const fillCoordScratch = { row: 0, col: 0 };
    const getCellComputedValue = (row: number, col: number) => {
      fillCoordScratch.row = row;
      fillCoordScratch.col = col;
      return this.getCellComputedValue(fillCoordScratch) as any;
    };

    // When possible, prefer engine-backed formula shifting for the fill shortcut. For multi-range
    // selections we compute all edits (async) first, then apply them in a single DocumentController
    // batch so the operation remains one undo step.
    if (wasm && operations.length === 1) {
      const op = operations[0]!;
      const task = applyFillCommitToDocumentControllerWithFormulaRewrite({
        document: this.document,
        sheetId: this.sheetId,
        sourceRange: op.sourceRange,
        targetRange: op.targetRange,
        mode,
        getCellComputedValue,
        rewriteFormulasForCopyDelta: (requests) => wasm.rewriteFormulasForCopyDelta(requests),
        label,
      })
        .catch(() => {
          // Fall back to legacy fill behavior if the worker is unavailable.
          this.document.beginBatch({ label });
          try {
            applyFillCommitToDocumentController({
              document: this.document,
              sheetId: this.sheetId,
              sourceRange: op.sourceRange,
              targetRange: op.targetRange,
              mode,
              getCellComputedValue,
            });
          } finally {
            this.document.endBatch();
          }
        })
        .finally(() => {
          this.refresh();
          this.focus();
        });
      this.idle.track(task);
      return;
    }

    if (wasm && operations.length > 1) {
      const task = (async () => {
        const coordScratch = { row: 0, col: 0 };
        const edits: Array<{ row: number; col: number; value: unknown }> = [];

        for (const op of operations) {
          const computed = await computeFillEditsForDocumentControllerWithFormulaRewrite({
            document: this.document,
            sheetId: this.sheetId,
            sourceRange: op.sourceRange,
            targetRange: op.targetRange,
            mode,
            getCellComputedValue,
            rewriteFormulasForCopyDelta: (requests) => wasm.rewriteFormulasForCopyDelta(requests),
          });
          for (const edit of computed) {
            edits.push(edit);
          }
        }

        if (edits.length === 0) return;

        this.document.beginBatch({ label });
        try {
          for (const edit of edits) {
            coordScratch.row = edit.row;
            coordScratch.col = edit.col;
            this.document.setCellInput(this.sheetId, coordScratch, edit.value);
          }
        } finally {
          this.document.endBatch();
        }
      })()
        .catch(() => {
          // Fall back to legacy fill behavior if the worker is unavailable.
          this.document.beginBatch({ label });
          try {
            for (const op of operations) {
              applyFillCommitToDocumentController({
                document: this.document,
                sheetId: this.sheetId,
                sourceRange: op.sourceRange,
                targetRange: op.targetRange,
                mode,
                getCellComputedValue,
              });
            }
          } finally {
            this.document.endBatch();
          }
        })
        .finally(() => {
          this.refresh();
          this.focus();
        });
      this.idle.track(task);
      return;
    }

    this.document.beginBatch({ label });
    try {
      for (const op of operations) {
        applyFillCommitToDocumentController({
          document: this.document,
          sheetId: this.sheetId,
          sourceRange: op.sourceRange,
          targetRange: op.targetRange,
          mode,
          getCellComputedValue
        });
      }
    } finally {
      this.document.endBatch();
    }

    this.refresh();
    this.focus();
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (this.inlineEditController.isOpen()) {
      return;
    }
    if (this.editor.isOpen()) {
      // The editor handles Enter/Tab/Escape itself. We keep focus on the textarea.
      return;
    }

    const target = e.target as HTMLElement | null;
    if (target && (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)) {
      return;
    }

    // Drawings take precedence over grid keyboard shortcuts when an object is selected.
    // This includes overriding `Ctrl/Cmd+D` (Fill Down) and `Ctrl/Cmd+[` / `Ctrl/Cmd+]` (Auditing)
    // so drawing manipulation remains usable in shared-grid mode.
    if (this.handleDrawingKeyDown(e)) {
      return;
    }

    // Other desktop UI surfaces (menus, global shortcuts, etc) may handle keyboard events
    // at the window level. If an earlier handler already called `preventDefault()`, treat
    // the event as consumed and don't apply spreadsheet keyboard behavior on top.
    if (e.defaultPrevented) {
      return;
    }

    if (e.key === "Escape") {
      if (this.dragState?.mode === "fill") {
        e.preventDefault();
        const state = this.dragState;
        this.dragState = null;
        this.dragPointerPos = null;
        this.fillPreviewRange = null;
        if (this.dragAutoScrollRaf != null) {
          if (typeof cancelAnimationFrame === "function") cancelAnimationFrame(this.dragAutoScrollRaf);
          else globalThis.clearTimeout(this.dragAutoScrollRaf);
        }
        this.dragAutoScrollRaf = null;

        try {
          this.root.releasePointerCapture(state.pointerId);
        } catch {
          // ignore
        }

        this.renderSelection();

        if (this.auditingNeedsUpdateAfterDrag) {
          this.auditingNeedsUpdateAfterDrag = false;
          this.scheduleAuditingUpdate();
        }
        return;
      }

      if (this.sharedGrid?.cancelFillHandleDrag()) {
        e.preventDefault();
        return;
      }

      // Excel-like: Escape clears chart selection / cancels an in-progress chart drag.
      //
      // Do not interfere with formula bar editing: Escape should cancel the edit there (handled below).
      if (!this.formulaBar?.isEditing() && !this.formulaEditCell) {
        if (this.chartDragState) {
          e.preventDefault();
          const state = this.chartDragState;
          this.chartDragState = null;
          this.chartDragAbort?.abort();
          this.chartDragAbort = null;
          // Revert the anchor to the initial pointerdown snapshot.
          this.chartStore.updateChartAnchor(state.chartId, state.startAnchor as any);
          return;
        }
        if (this.selectedChartId != null) {
          e.preventDefault();
          this.setSelectedChartId(null);
          return;
        }
      }
    }

    // While the formula bar is actively editing, Enter/Escape should commit/cancel the formula
    // edit even if focus temporarily moved back to the grid (Excel-style range selection mode).
    //
    // This prevents "Enter moves selection" / "Escape does nothing" behavior while the user
    // is still building a formula in the formula bar.
    if (this.formulaBar?.isEditing() || this.formulaEditCell) {
      const primary = e.ctrlKey || e.metaKey;
      if (e.key === "Escape") {
        this.endKeyboardRangeSelection();
        e.preventDefault();
        this.formulaBar?.cancelEdit();
        return;
      }
      // Match FormulaBarView: Enter commits, Alt+Enter inserts newline.
      if (e.key === "Enter" && !e.altKey) {
        this.endKeyboardRangeSelection();
        e.preventDefault();
        this.formulaBar?.commitEdit("enter", e.shiftKey);
        return;
      }
      // Match FormulaBarView: Tab/Shift+Tab commits (and the app navigates selection). Prevent
      // browser focus traversal while editing, even if the grid temporarily has focus.
      if (e.key === "Tab") {
        this.endKeyboardRangeSelection();
        e.preventDefault();
        this.formulaBar?.commitEdit("tab", e.shiftKey);
        return;
      }

      // Excel-like: while editing a formula, F4 toggles absolute/relative references.
      // Route it to the formula bar even if focus is temporarily on the grid (range selection mode).
      if (
        e.key === "F4" &&
        !e.altKey &&
        !e.ctrlKey &&
        !e.metaKey &&
        this.formulaBar &&
        this.formulaBar.textarea.value.trim().startsWith("=")
      ) {
        this.endKeyboardRangeSelection();
        e.preventDefault();
        const textarea = this.formulaBar.textarea;
        const prevText = textarea.value;
        const cursorStart = textarea.selectionStart ?? prevText.length;
        const cursorEnd = textarea.selectionEnd ?? prevText.length;
        const toggled = toggleA1AbsoluteAtCursor(prevText, cursorStart, cursorEnd);
        if (toggled) {
          textarea.value = toggled.text;
          textarea.setSelectionRange(toggled.cursorStart, toggled.cursorEnd);
          textarea.dispatchEvent(new Event("input", { bubbles: true }));
        }
        this.formulaBar.focus();
        return;
      }

      // In range-selection mode, focus may temporarily move to the grid. Ensure deletion keys still
      // edit the formula bar text (and do not clear sheet contents).
      if ((e.key === "Backspace" || e.key === "Delete") && this.formulaBar) {
        this.endKeyboardRangeSelection();
        e.preventDefault();
        const textarea = this.formulaBar.textarea;
        const current = textarea.value;
        const selStart = textarea.selectionStart ?? current.length;
        const selEnd = textarea.selectionEnd ?? current.length;
        const start = Math.max(0, Math.min(selStart, selEnd, current.length));
        const end = Math.max(0, Math.min(Math.max(selStart, selEnd), current.length));

        const nextValue = (() => {
          if (start !== end) {
            return current.slice(0, start) + current.slice(end);
          }
          if (e.key === "Backspace") {
            if (start === 0) return current;
            return current.slice(0, start - 1) + current.slice(start);
          }
          // Delete
          if (start >= current.length) return current;
          return current.slice(0, start) + current.slice(start + 1);
        })();

        if (nextValue !== current) {
          textarea.value = nextValue;
        }

        const cursor = (() => {
          if (start !== end) return start;
          if (e.key === "Backspace") return Math.max(0, start - 1);
          return start;
        })();
        textarea.setSelectionRange(cursor, cursor);
        textarea.dispatchEvent(new Event("input", { bubbles: true }));
        this.formulaBar.focus();
        return;
      }

      // Route common "text editing" shortcuts (undo/redo/select-all/clipboard) into the formula bar
      // even when focus is currently on the grid (range selection mode).
      if (this.formulaBar && primary && !e.altKey) {
        const textarea = this.formulaBar.textarea;
        const key = e.key.toLowerCase();

        if (isUndoKeyboardEvent(e) || isRedoKeyboardEvent(e)) {
          this.endKeyboardRangeSelection();
          e.preventDefault();
          this.formulaBar.focus();
          try {
            document.execCommand(isUndoKeyboardEvent(e) ? "undo" : "redo", false);
          } catch {
            // Best-effort.
          }
          return;
        }

        if (!e.shiftKey && key === "a") {
          this.endKeyboardRangeSelection();
          e.preventDefault();
          this.formulaBar.focus();
          textarea.setSelectionRange(0, textarea.value.length);
          return;
        }

        if (!e.shiftKey && (key === "c" || key === "x" || key === "v")) {
          this.endKeyboardRangeSelection();
          e.preventDefault();
          this.formulaBar.focus();
          const command = key === "c" ? "copy" : key === "x" ? "cut" : "paste";
          try {
            document.execCommand(command, false);
          } catch {
            // Best-effort: clipboard execCommand can be blocked by platform permissions.
          }
          return;
        }
      }
    }

    if (this.handleUndoRedoShortcut(e)) return;
    if (this.handleShowFormulasShortcut(e)) return;
    if (this.handleAuditingShortcut(e)) return;
    if (this.handleClipboardShortcut(e)) return;
    if (this.handleFormattingShortcut(e)) return;
    if (this.handleInsertDateTimeShortcut(e)) return;
    if (this.handleAutoSumShortcut(e)) return;
    if (this.handleInsertImageShortcut(e)) return;

    // Editing
    // Excel-style: Shift+F2 adds/edits a comment (we wire this to "Add Comment").
    if (e.key === "F2" && e.shiftKey) {
      // Avoid opening comment UI while the formula bar is actively editing (range selection mode).
      if (this.formulaBar?.isEditing() || this.formulaEditCell) {
        // Prevent the global keybinding layer from running the Comments command while the user is
        // selecting ranges for a formula in the formula bar.
        e.preventDefault();
        return;
      }
      e.preventDefault();
      this.openCommentsPanel();
      this.focusNewCommentInput();
      return;
    }

    if (e.key === "F2") {
      // In-cell editing should never start while the formula bar is actively editing (range selection mode).
      if (this.formulaBar?.isEditing() || this.formulaEditCell) return;
      e.preventDefault();
      const cell = this.selection.active;
      if (this.isReadOnly()) {
        showCollabEditRejectedToast([
          { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
        ]);
        return;
      }
      const bounds = this.getCellRect(cell);
      if (!bounds) return;
      const initialValue = this.getCellInputText(cell);
      this.editor.open(cell, bounds, initialValue, { cursor: "end" });
      this.updateEditState();
      return;
    }

    const primary = e.ctrlKey || e.metaKey;
    // Excel-style fill shortcuts:
    // - Ctrl/Cmd+D: Fill Down
    // - Ctrl/Cmd+R: Fill Right
    // Only trigger when *not* editing text (editor, formula bar, inline edit).
    if (primary && !e.altKey && !e.shiftKey && (e.key === "d" || e.key === "D")) {
      if (this.formulaBar?.isEditing() || this.formulaEditCell) return;
      e.preventDefault();
      this.fillDown();
      return;
    }

    if (primary && !e.altKey && !e.shiftKey && (e.key === "r" || e.key === "R")) {
      if (this.formulaBar?.isEditing() || this.formulaEditCell) return;
      e.preventDefault();
      this.fillRight();
      return;
    }

    if (primary && (e.key === "k" || e.key === "K")) {
      // Inline edit (Cmd/Ctrl+K) should not trigger while the formula bar is actively editing.
      if (this.formulaBar?.isEditing() || this.formulaEditCell) return;
      e.preventDefault();
      this.openInlineAiEdit();
      return;
    }
    if ((e.key === "Backspace" || e.key === "Delete") && this.selectedDrawingId != null) {
      // Picture deletion should not fire while the formula bar is editing (including
      // range-selection mode).
      if (this.formulaBar?.isEditing() || this.formulaEditCell) return;
      e.preventDefault();
      this.deleteSelectedDrawing();
      return;
    }
    if (e.key === "Delete") {
      // Delete should never clear sheet contents while the formula bar is editing (including
      // range-selection mode, where focus may temporarily move to the grid).
      if (this.formulaBar?.isEditing() || this.formulaEditCell) return;
      e.preventDefault();
      this.clearSelectionContents();
      return;
    }

    // Ctrl/Cmd+Shift+M toggles the comments panel.
    if (primary && e.shiftKey && (e.key === "m" || e.key === "M")) {
      e.preventDefault();
      this.toggleCommentsPanel();
      return;
    }

    // Selection shortcuts
    if (primary && !e.shiftKey && (e.key === "a" || e.key === "A")) {
      e.preventDefault();
      this.selection = selectAll(this.limits);
      if (this.sharedGrid) this.syncSharedGridSelectionFromState();
      this.renderSelection();
      this.updateStatus();
      return;
    }

    // Excel-style select current region: Ctrl/Cmd+Shift+* (aka Ctrl/Cmd+Shift+8).
    // Use `code===Digit8` to catch layouts where `key` is not "*".
    //
    // Some keyboards also have a dedicated Numpad "*" key; Excel typically accepts Ctrl+*
    // there even without Shift, so we include `code===NumpadMultiply` as a best-effort.
    if (
      primary &&
      !e.altKey &&
      ((e.shiftKey && (e.code === "Digit8" || e.key === "*" || e.key === "8")) || e.code === "NumpadMultiply")
    ) {
      e.preventDefault();
      this.selectCurrentRegion();
      return;
    }

    if (primary && e.code === "Space") {
      // Ctrl+Space selects entire column.
      e.preventDefault();
      this.selection = selectColumns(this.selection, this.selection.active.col, this.selection.active.col, {}, this.limits);
      if (this.sharedGrid) this.syncSharedGridSelectionFromState();
      this.renderSelection();
      this.updateStatus();
      return;
    }

    // Sheet navigation (Excel-style): Ctrl+PgUp / Ctrl+PgDn.
    //
    // NOTE: The desktop shell owns sheet metadata (ordering + visibility) via a sheet store.
    // Do not implement sheet navigation here based on DocumentController sheet ids; instead,
    // allow the global keybinding layer to dispatch `workbook.previousSheet`/`workbook.nextSheet`
    // using the sheet store's visible order.
    if (!primary && e.shiftKey && e.code === "Space") {
      // Shift+Space selects entire row.
      e.preventDefault();
      this.selection = selectRows(this.selection, this.selection.active.row, this.selection.active.row, {}, this.limits);
      if (this.sharedGrid) this.syncSharedGridSelectionFromState();
      this.renderSelection();
      this.updateStatus();
      return;
    }

    // Outline grouping shortcuts (Excel-style): Alt+Shift+Right/Left.
    if (e.altKey && e.shiftKey && (e.key === "ArrowRight" || e.key === "ArrowLeft")) {
      if (this.sharedGrid) {
        // Shared-grid mode doesn't currently implement outline groups.
        // Treat the shortcut as a no-op (legacy outline logic rebuilds visibility caches which are
        // too expensive for large sheets).
        e.preventDefault();
        return;
      }
      e.preventDefault();
      const range = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0];
      if (!range) return;

      const startRow = range.startRow + 1;
      const endRow = range.endRow + 1;
      const startCol = range.startCol + 1;
      const endCol = range.endCol + 1;

      const outline = this.getOutlineForSheet(this.sheetId);
      if (e.key === "ArrowRight") {
        if (this.selection.type === "column") {
          outline.groupCols(startCol, endCol);
          outline.recomputeOutlineHiddenCols();
        } else {
          outline.groupRows(startRow, endRow);
          outline.recomputeOutlineHiddenRows();
        }
      } else {
        if (this.selection.type === "column") {
          outline.ungroupCols(startCol, endCol);
        } else {
          outline.ungroupRows(startRow, endRow);
        }
      }

      this.onOutlineUpdated();
      return;
    }

    // Page navigation (Excel-style): PageUp/PageDown move by approximately one viewport.
    // Alt+PageUp/PageDown scroll horizontally.
    if (!primary && (e.key === "PageDown" || e.key === "PageUp")) {
      e.preventDefault();
      this.ensureActiveCellVisible();
      const dir = e.key === "PageDown" ? 1 : -1;
      if (this.sharedGrid) {
        const viewport = this.sharedGrid.renderer.scroll.getViewportState();
        const scrollableHeight = Math.max(0, viewport.height - viewport.frozenHeight);
        const scrollableWidth = Math.max(0, viewport.width - viewport.frozenWidth);

        const rowAxis = this.sharedGrid.renderer.scroll.rows;
        const colAxis = this.sharedGrid.renderer.scroll.cols;
        const pageRows = Math.max(1, Math.floor(scrollableHeight / rowAxis.defaultSize));
        const pageCols = Math.max(1, Math.floor(scrollableWidth / colAxis.defaultSize));

        if (e.altKey) {
          const col = Math.max(
            0,
            Math.min(this.limits.maxCols - 1, this.selection.active.col + dir * pageCols)
          );
          this.selection = e.shiftKey
            ? extendSelectionToCell(this.selection, { row: this.selection.active.row, col }, this.limits)
            : setActiveCell(this.selection, { row: this.selection.active.row, col }, this.limits);
        } else {
          const row = Math.max(
            0,
            Math.min(this.limits.maxRows - 1, this.selection.active.row + dir * pageRows)
          );
          this.selection = e.shiftKey
            ? extendSelectionToCell(this.selection, { row, col: this.selection.active.col }, this.limits)
            : setActiveCell(this.selection, { row, col: this.selection.active.col }, this.limits);
        }
      } else {
        const visualRow = this.rowToVisual.get(this.selection.active.row);
        const visualCol = this.colToVisual.get(this.selection.active.col);
        if (visualRow === undefined || visualCol === undefined) return;

        if (e.altKey) {
          const pageCols = Math.max(1, Math.floor(this.viewportWidth() / this.cellWidth));
          const nextColVisual = Math.min(
            Math.max(0, visualCol + dir * pageCols),
            Math.max(0, this.colIndexByVisual.length - 1)
          );
          const col = this.colIndexByVisual[nextColVisual] ?? 0;
          this.selection = e.shiftKey
            ? extendSelectionToCell(this.selection, { row: this.selection.active.row, col }, this.limits)
            : setActiveCell(this.selection, { row: this.selection.active.row, col }, this.limits);
        } else {
          const pageRows = Math.max(1, Math.floor(this.viewportHeight() / this.cellHeight));
          const nextRowVisual = Math.min(
            Math.max(0, visualRow + dir * pageRows),
            Math.max(0, this.rowIndexByVisual.length - 1)
          );
          const row = this.rowIndexByVisual[nextRowVisual] ?? 0;
          this.selection = e.shiftKey
            ? extendSelectionToCell(this.selection, { row, col: this.selection.active.col }, this.limits)
            : setActiveCell(this.selection, { row, col: this.selection.active.col }, this.limits);
        }
      }

      this.ensureActiveCellVisible();
      const didScroll = this.scrollCellIntoView(this.selection.active);
      if (this.sharedGrid) this.syncSharedGridSelectionFromState({ scrollIntoView: false });
      else if (didScroll) this.ensureViewportMappingCurrent();
      this.renderSelection();
      this.updateStatus();
      if (didScroll) this.refresh("scroll");
      return;
    }

    // Home/End without Ctrl/Cmd.
    // Home: first visible column in the current row.
    // End: last visible column in the current row.
    if (!primary && !e.altKey && (e.key === "Home" || e.key === "End")) {
      e.preventDefault();
      this.ensureActiveCellVisible();
      const row = this.selection.active.row;
      const col = (() => {
        if (!this.sharedGrid) {
          return e.key === "Home"
            ? (this.colIndexByVisual[0] ?? 0)
            : (this.colIndexByVisual[this.colIndexByVisual.length - 1] ?? 0);
        }
        return e.key === "Home" ? 0 : this.limits.maxCols - 1;
      })();
      this.selection = e.shiftKey
        ? extendSelectionToCell(this.selection, { row, col }, this.limits)
        : setActiveCell(this.selection, { row, col }, this.limits);
      this.ensureActiveCellVisible();
      const didScroll = this.scrollCellIntoView(this.selection.active);
      if (this.sharedGrid) this.syncSharedGridSelectionFromState({ scrollIntoView: false });
      else if (didScroll) this.ensureViewportMappingCurrent();
      this.renderSelection();
      this.updateStatus();
      if (didScroll) this.refresh("scroll");
      return;
    }

    // Excel-like "start typing to edit" behavior: any printable key begins edit
    // mode and replaces the cell contents.
    if (!primary && !e.altKey && e.key.length === 1) {
      // When the formula bar is editing, do not start editing the active cell. Instead, treat
      // printable key presses as formula bar input even if focus temporarily moved back to the grid
      // (Excel-style range selection mode).
      //
      // NOTE: We avoid calling `insertIntoFormulaBar()` here because it contains special-case logic
      // for inserting leading `=` templates (e.g. `=SUM()`), which would break normal typing of
      // `=` inside formulas.
      if (this.formulaBar?.isEditing() || this.formulaEditCell) {
        const bar = this.formulaBar;
        if (bar) {
          this.endKeyboardRangeSelection();
          e.preventDefault();
          const textarea = bar.textarea;
          const current = textarea.value;
          const selStart = textarea.selectionStart ?? current.length;
          const selEnd = textarea.selectionEnd ?? current.length;
          const start = Math.max(0, Math.min(selStart, selEnd, current.length));
          const end = Math.max(0, Math.min(Math.max(selStart, selEnd), current.length));
          const next = current.slice(0, start) + e.key + current.slice(end);
          textarea.value = next;
          const cursor = Math.max(0, Math.min(start + e.key.length, next.length));
          textarea.setSelectionRange(cursor, cursor);
          textarea.dispatchEvent(new Event("input", { bubbles: true }));
          bar.focus();
        }
        return;
      }
      e.preventDefault();
      if (this.isReadOnly()) return;
      const cell = this.selection.active;
      const bounds = this.getCellRect(cell);
      if (!bounds) return;
      this.editor.open(cell, bounds, e.key, { cursor: "end" });
      this.updateEditState();
      return;
    }

    const next = navigateSelectionByKey(
      this.selection,
      e.key,
      { shift: e.shiftKey, primary },
      this.usedRangeProvider(),
      this.limits
    );
    if (!next) return;

    e.preventDefault();
    this.selection = next;

    // Keyboard "point mode": while editing a formula in the formula bar, arrow-key navigation
    // should update the formula draft by inserting/replacing a reference token that matches the
    // current selection range.
    if (
      this.formulaBar?.isFormulaEditing() &&
      !e.isComposing &&
      (e.key === "ArrowUp" || e.key === "ArrowDown" || e.key === "ArrowLeft" || e.key === "ArrowRight")
    ) {
      const r = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0];
      if (r) {
        const rangeSheetId = this.formulaEditCell && this.formulaEditCell.sheetId !== this.sheetId ? this.sheetId : undefined;
        const rangeSheetName = rangeSheetId ? this.resolveSheetDisplayNameById(rangeSheetId) : undefined;
        const a1Range = { start: { row: r.startRow, col: r.startCol }, end: { row: r.endRow, col: r.endCol } };
        if (this.keyboardRangeSelectionActive) {
          this.formulaBar.updateRangeSelection(a1Range, rangeSheetName);
        } else {
          this.formulaBar.beginRangeSelection(a1Range, rangeSheetName);
          this.keyboardRangeSelectionActive = true;
        }
      }
    }
    const didScroll = this.scrollCellIntoView(this.selection.active);
    if (this.sharedGrid) this.syncSharedGridSelectionFromState({ scrollIntoView: false });
    else if (didScroll) this.ensureViewportMappingCurrent();
    this.renderSelection();
    this.updateStatus();
    if (didScroll) this.refresh("scroll");
  }

  private endKeyboardRangeSelection(): void {
    if (!this.keyboardRangeSelectionActive) return;
    this.keyboardRangeSelectionActive = false;
    this.formulaBar?.endRangeSelection();
  }

  private handleInsertDateTimeShortcut(e: KeyboardEvent): boolean {
    const primary = e.ctrlKey || e.metaKey;
    if (!primary || e.altKey) return false;

    // Excel-style:
    // - Ctrl+; inserts the current date
    // - Ctrl+Shift+; inserts the current time
    //
    // `e.code` is layout-independent and stays "Semicolon" whether Shift is pressed (":" on US keyboards).
    if (e.code !== "Semicolon") return false;

    // Do not trigger while editing.
    if (this.editor.isOpen()) return false;
    if (this.inlineEditController.isOpen()) return false;
    if (this.formulaBar?.isEditing() || this.formulaEditCell) return false;
    e.preventDefault();
    if (e.shiftKey) this.insertTime();
    else this.insertDate();
    return true;
  }

  private insertCurrentDateTimeIntoSelection(kind: "date" | "time"): void {
    this.insertCurrentDateTimeIntoSelectionExcelSerial(kind);
  }

  private insertCurrentDateTimeIntoSelectionExcelSerial(kind: "date" | "time"): void {
    const label = kind === "date" ? t("command.edit.insertDate") : t("command.edit.insertTime");
    const numberFormat = kind === "date" ? "yyyy-mm-dd" : "hh:mm:ss";

    const now = new Date();
    const serial = (() => {
      if (kind === "date") {
        // Excel dates are stored as timezone-agnostic serial numbers. Interpret the current
        // local calendar date as a UTC day for deterministic storage (mirrors parseScalar).
        const utcDate = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
        return dateToExcelSerial(utcDate);
      }

      const seconds = now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds();
      return seconds / 86_400;
    })();

    const selectionRanges =
      this.selection.ranges.length > 0
        ? this.selection.ranges
        : [
            {
              startRow: this.selection.active.row,
              endRow: this.selection.active.row,
              startCol: this.selection.active.col,
              endCol: this.selection.active.col,
            },
          ];

    let totalCells = 0;
    for (const range of selectionRanges) {
      const r = normalizeSelectionRange(range);
      const rows = Math.max(0, r.endRow - r.startRow + 1);
      const cols = Math.max(0, r.endCol - r.startCol + 1);
      totalCells += rows * cols;
      if (totalCells > MAX_DATE_TIME_INSERT_CELLS) break;
    }

    const ranges =
      totalCells > MAX_DATE_TIME_INSERT_CELLS
        ? [
            {
              startRow: this.selection.active.row,
              endRow: this.selection.active.row,
              startCol: this.selection.active.col,
              endCol: this.selection.active.col,
            },
          ]
        : selectionRanges;

    this.document.beginBatch({ label });
    try {
      for (const range of ranges) {
        const r = normalizeSelectionRange(range);
        const rowCount = Math.max(0, r.endRow - r.startRow + 1);
        const colCount = Math.max(0, r.endCol - r.startCol + 1);
        if (rowCount === 0 || colCount === 0) continue;

        // Use shared row data to avoid allocating per-row arrays for uniform fills.
        const rowValues = Array(colCount).fill(serial);
        const values = Array(rowCount).fill(rowValues);

        this.document.setRangeValues(
          this.sheetId,
          { start: { row: r.startRow, col: r.startCol }, end: { row: r.endRow, col: r.endCol } },
          values,
        );
        this.document.setRangeFormat(
          this.sheetId,
          { start: { row: r.startRow, col: r.startCol }, end: { row: r.endRow, col: r.endCol } },
          { numberFormat },
        );
      }
    } finally {
      this.document.endBatch();
    }
  }

  private handleAutoSumShortcut(e: KeyboardEvent): boolean {
    if (!e.altKey) return false;
    if (e.code !== "Equal") return false;
    // Avoid hijacking Ctrl/Cmd-modified shortcuts.
    if (e.ctrlKey || e.metaKey) return false;

    // Only trigger when not actively editing.
    if (this.formulaBar?.isEditing() || this.formulaEditCell) return false;

    e.preventDefault();
    this.autoSum();
    return true;
  }

  private handleInsertImageShortcut(e: KeyboardEvent): boolean {
    const primary = e.ctrlKey || e.metaKey;
    if (!primary || !e.shiftKey || e.altKey) return false;
    if (e.key.toLowerCase() !== "i") return false;
    if (this.formulaBar?.isEditing() || this.formulaEditCell) return false;
    e.preventDefault();
    this.insertImageFromLocalFile();
    return true;
  }

  private autoSumSelection(fn: "SUM" | "AVERAGE" | "COUNT" | "MAX" | "MIN"): void {
    const sheetId = this.sheetId;
    const coordScratch = { row: 0, col: 0 };

    const normalizeRange = (range: Range): Range => ({
      startRow: Math.min(range.startRow, range.endRow),
      endRow: Math.max(range.startRow, range.endRow),
      startCol: Math.min(range.startCol, range.endCol),
      endCol: Math.max(range.startCol, range.endCol),
    });

    const getCellState = (row: number, col: number): { value: unknown; formula: string | null } | null => {
      coordScratch.row = row;
      coordScratch.col = col;
      const state = this.document.getCell(sheetId, coordScratch) as { value: unknown; formula: string | null };
      return state ?? null;
    };

    const isEmptyCell = (row: number, col: number): boolean => {
      const state = getCellState(row, col);
      if (!state) return true;
      return state.value == null && state.formula == null;
    };

    const chooseTargetFromSelection = (): { target: CellCoord; formulaRange: Range } | null => {
      if (this.selection.ranges.length !== 1) return null;
      const range = normalizeRange(this.selection.ranges[0]!);
      const isSingleRow = range.startRow === range.endRow;
      const isSingleCol = range.startCol === range.endCol;

      // Excel-style: if the user selects a vertical or horizontal range of values,
      // AutoSum inserts the SUM formula just below / to the right (or in the last
      // selected cell if it's empty).
      if (isSingleCol && range.endRow > range.startRow) {
        const bottom = { row: range.endRow, col: range.startCol };
        if (isEmptyCell(bottom.row, bottom.col)) {
          const formulaRange: Range = {
            startRow: range.startRow,
            endRow: range.endRow - 1,
            startCol: range.startCol,
            endCol: range.endCol,
          };
          if (formulaRange.endRow < formulaRange.startRow) return null;
          return { target: bottom, formulaRange };
        }

        const nextRow = range.endRow + 1;
        if (nextRow >= this.limits.maxRows) return null;
        return { target: { row: nextRow, col: range.startCol }, formulaRange: range };
      }

      if (isSingleRow && range.endCol > range.startCol) {
        const right = { row: range.startRow, col: range.endCol };
        if (isEmptyCell(right.row, right.col)) {
          const formulaRange: Range = {
            startRow: range.startRow,
            endRow: range.endRow,
            startCol: range.startCol,
            endCol: range.endCol - 1,
          };
          if (formulaRange.endCol < formulaRange.startCol) return null;
          return { target: right, formulaRange };
        }

        const nextCol = range.endCol + 1;
        if (nextCol >= this.limits.maxCols) return null;
        return { target: { row: range.startRow, col: nextCol }, formulaRange: range };
      }

      return null;
    };

    const selected = chooseTargetFromSelection();
    if (selected) {
      // Move the active cell to the insertion point (Excel behavior) and keep the
      // operation as a single undoable edit.
      this.selection = setActiveCell(this.selection, selected.target, this.limits);
      if (this.sharedGrid) this.syncSharedGridSelectionFromState({ scrollIntoView: true });
      else this.scrollCellIntoView(selected.target);

      const formula = `=${fn}(${rangeToA1(selected.formulaRange)})`;
      this.applyEdit(sheetId, selected.target, formula, { label: "AutoSum" });
      return;
    }

    const active = this.selection.active;

    const isNumericishCell = (row: number, col: number): boolean => {
      const state = getCellState(row, col);
      if (!state) return false;
      if (state.formula != null) {
        coordScratch.row = row;
        coordScratch.col = col;
        const computed = this.getCellComputedValue(coordScratch);
        return typeof computed === "number" && Number.isFinite(computed);
      }
      return coerceNumber(state.value) != null;
    };

    const formulaRange = (() => {
      // Prefer a contiguous numeric block above the active cell in the same column.
      if (active.row > 0 && isNumericishCell(active.row - 1, active.col)) {
        let startRow = active.row - 1;
        while (startRow > 0 && isNumericishCell(startRow - 1, active.col)) startRow -= 1;
        return { startRow, endRow: active.row - 1, startCol: active.col, endCol: active.col };
      }

      // Else, try a contiguous block to the left in the same row.
      if (active.col > 0 && isNumericishCell(active.row, active.col - 1)) {
        let startCol = active.col - 1;
        while (startCol > 0 && isNumericishCell(active.row, startCol - 1)) startCol -= 1;
        return { startRow: active.row, endRow: active.row, startCol, endCol: active.col - 1 };
      }

      return null;
    })();

    if (!formulaRange) return;

    const formula = `=${fn}(${rangeToA1(formulaRange)})`;
    this.applyEdit(this.sheetId, active, formula, { label: t("command.edit.autoSum") });
  }

  private shouldHandleSpreadsheetClipboardCommand(): boolean {
    if (this.formulaBar?.isEditing() || this.formulaEditCell) return false;
    const target = document.activeElement as HTMLElement | null;
    if (!target) return true;
    const tag = target.tagName;
    if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return false;
    return true;
  }

  /**
   * Clipboard actions are normally triggered by keyboard shortcuts on the grid.
   * These wrappers exist so menus/commands can execute the same behavior.
   */
  copyToClipboard(): Promise<void> {
    if (!this.shouldHandleSpreadsheetClipboardCommand()) return Promise.resolve();
    const promise = this.copySelectionToClipboard();
    this.idle.track(promise);
    return promise.finally(() => {
      this.focus();
    });
  }

  cutToClipboard(): Promise<void> {
    if (!this.shouldHandleSpreadsheetClipboardCommand()) return Promise.resolve();
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return Promise.resolve();
    }
    const promise = this.cutSelectionToClipboard();
    this.idle.track(promise);
    return promise.finally(() => {
      this.focus();
    });
  }

  pasteFromClipboard(): Promise<void> {
    if (!this.shouldHandleSpreadsheetClipboardCommand()) return Promise.resolve();
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return Promise.resolve();
    }
    const promise = this.pasteClipboardToSelection();
    this.idle.track(promise);
    return promise.finally(() => {
      this.focus();
    });
  }

  private handleFormattingShortcut(e: KeyboardEvent): boolean {
    if (e.altKey) return false;

    // Formatting shortcuts should never fire while editing text.
    if (this.formulaBar?.isEditing() || this.formulaEditCell) return false;

    const target = e.target as HTMLElement | null;
    if (target) {
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return false;
    }

    const keyLower = e.key.toLowerCase();
    const primary = e.ctrlKey || e.metaKey;

    const action:
      | { kind: "bold" }
      | { kind: "underline" }
      | { kind: "italic" }
      | { kind: "strikethrough" }
      | { kind: "numberFormat"; preset: "currency" | "percent" | "date" }
      | null = (() => {
      // Text formatting.
      if (primary && !e.shiftKey && keyLower === "b") return { kind: "bold" };
      if (primary && !e.shiftKey && keyLower === "u") return { kind: "underline" };
      // Cmd+I is reserved for toggling the AI sidebar. Only bind italic to Ctrl+I.
      if (!e.shiftKey && keyLower === "i" && e.ctrlKey && !e.metaKey) return { kind: "italic" };
      // Excel: Ctrl+5 toggles strikethrough.
      if (!e.shiftKey && e.ctrlKey && !e.metaKey && (keyLower === "5" || e.code === "Digit5")) return { kind: "strikethrough" };

      // Number formats.
      if (primary && e.shiftKey && (e.key === "$" || e.code === "Digit4"))
        return { kind: "numberFormat", preset: "currency" };
      if (primary && e.shiftKey && (e.key === "%" || e.code === "Digit5"))
        return { kind: "numberFormat", preset: "percent" };
      if (primary && e.shiftKey && (e.key === "#" || e.code === "Digit3")) return { kind: "numberFormat", preset: "date" };

      return null;
    })();

    if (!action) return false;

    e.preventDefault();

    const selectionRanges = this.selection.ranges.length
      ? this.selection.ranges
      : [
          {
            startRow: this.selection.active.row,
            endRow: this.selection.active.row,
            startCol: this.selection.active.col,
            endCol: this.selection.active.col,
          },
        ];

    // When the UI selection is a full row/column/sheet, expand it to the canonical Excel bounds
    // so DocumentController can use fast layered-format paths (sheet/row/col style ids) without
    // enumerating every cell.
    //
    // Back-compat: older "legacy" grid selections used a 10k x 200 coordinate space. Those
    // persisted selections should still behave as full-row/col selections when applying
    // formatting shortcuts.
    const ranges = selectionRanges.map((range) => {
      const r = normalizeSelectionRange(range);
      const legacyMaxRows = 10_000;
      const legacyMaxCols = 200;
      const isFullColBand =
        r.startRow === 0 && (r.endRow === this.limits.maxRows - 1 || r.endRow === legacyMaxRows - 1);
      const isFullRowBand =
        r.startCol === 0 && (r.endCol === this.limits.maxCols - 1 || r.endCol === legacyMaxCols - 1);
      return {
        startRow: Math.max(0, r.startRow),
        startCol: Math.max(0, r.startCol),
        endRow: isFullColBand ? DEFAULT_GRID_LIMITS.maxRows - 1 : Math.max(0, r.endRow),
        endCol: isFullRowBand ? DEFAULT_GRID_LIMITS.maxCols - 1 : Math.max(0, r.endCol),
      };
    });

    const decision = evaluateFormattingSelectionSize(ranges, DEFAULT_GRID_LIMITS, {
      maxCells: MAX_KEYBOARD_FORMATTING_CELLS,
    });

    if (!decision.allowed) {
      try {
        showToast(
          "Selection is too large to format. Try selecting fewer cells or an entire row/column.",
          "warning"
        );
      } catch {
        // `showToast` requires a #toast-root; unit tests don't always include it.
      }
      return true;
    }

    if (this.isReadOnly() && !decision.allRangesBand) {
      showCollabEditRejectedToast([{ sheetId: this.sheetId, rejectionKind: "format", rejectionReason: "permission" }]);
      return true;
    }

    const batchLabel = (() => {
      switch (action.kind) {
        case "bold":
          return "Bold";
        case "italic":
          return "Italic";
        case "underline":
          return "Underline";
        case "strikethrough":
          return "Strikethrough";
        case "numberFormat":
          return "Number format";
      }
    })();

    this.document.beginBatch({ label: batchLabel });
    let applied = true;
    try {
      for (const range of ranges) {
        const docRange = {
          start: { row: range.startRow, col: range.startCol },
          end: { row: range.endRow, col: range.endCol }
        };

        switch (action.kind) {
          case "bold":
            if (toggleBold(this.document, this.sheetId, docRange) === false) applied = false;
            break;
          case "italic":
            if (toggleItalic(this.document, this.sheetId, docRange) === false) applied = false;
            break;
          case "underline":
            if (toggleUnderline(this.document, this.sheetId, docRange) === false) applied = false;
            break;
          case "strikethrough":
            if (toggleStrikethrough(this.document, this.sheetId, docRange) === false) applied = false;
            break;
          case "numberFormat":
            if (applyNumberFormatPreset(this.document, this.sheetId, docRange, action.preset) === false) applied = false;
            break;
        }
      }
    } finally {
      this.document.endBatch();
    }
    if (!applied) {
      try {
        showToast("Formatting could not be applied to the full selection. Try selecting fewer cells/rows.", "warning");
      } catch {
        // `showToast` requires a #toast-root; unit tests don't always include it.
      }
    }

    this.refresh();
    this.updateStatus();
    this.focus();
    return true;
  }
  private handleClipboardShortcut(e: KeyboardEvent): boolean {
    const primary = e.ctrlKey || e.metaKey;
    if (!primary || e.altKey || e.shiftKey) return false;

    if (this.formulaBar?.isEditing() || this.formulaEditCell) return false;

    const target = e.target as HTMLElement | null;
    if (target) {
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || target.isContentEditable) return false;
    }

    const key = e.key.toLowerCase();

    if (key === "c") {
      e.preventDefault();
      this.idle.track(this.copySelectionToClipboard());
      return true;
    }

    if (key === "x") {
      e.preventDefault();
      if (this.isReadOnly()) {
        const cell = this.selection.active;
        showCollabEditRejectedToast([
          { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
        ]);
        return true;
      }
      this.idle.track(this.cutSelectionToClipboard());
      return true;
    }

    if (key === "v") {
      e.preventDefault();
      if (this.isReadOnly()) {
        const cell = this.selection.active;
        showCollabEditRejectedToast([
          { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
        ]);
        return true;
      }
      this.idle.track(this.pasteClipboardToSelection());
      return true;
    }

    return false;
  }

  private async getClipboardProvider(): Promise<Awaited<ReturnType<typeof createClipboardProvider>>> {
    if (!this.clipboardProviderPromise) {
      this.clipboardProviderPromise = createClipboardProvider();
    }
    return this.clipboardProviderPromise;
  }

  private snapshotClipboardCells(range: Range): Array<Array<{ value: unknown; formula: string | null; styleId: number }>> {
    // Snapshot the *effective* style for each cell so internal copy/paste preserves
    // inherited (layered) formatting (sheet/row/col defaults) even when
    // `cell.styleId` is 0.
    //
    // This intentionally interns the resolved style into `document.styleTable`
    // so the internal paste path can keep pasting styleIds (fast) rather than
    // materializing per-cell format objects.
    const docAny = this.document as any;
    const styleIdByLayerKey = new Map<string, number>();

    // Best-effort access to the underlying sheet model to cheaply derive a stable
    // (sheet,row,col,cell,range-run) style id tuple key without computing full merges for
    // every copied cell.
    const sheetModel = (() => {
      try {
        // Ensure sheet exists (DocumentController creates sheets lazily).
        docAny?.model?.getCell?.(this.sheetId, 0, 0);
        return docAny?.model?.sheets?.get?.(this.sheetId) ?? null;
      } catch {
        return null;
      }
    })();

    const normalizeStyleId = (value: unknown): number => {
      const n = Number(value);
      return Number.isInteger(n) && n >= 0 ? n : 0;
    };

    const coordScratch = { row: 0, col: 0 };

    const styleIdForRowInRuns = (runs: unknown, row: number): number => {
      if (!Array.isArray(runs) || runs.length === 0) return 0;
      let lo = 0;
      let hi = runs.length - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        const run = runs[mid] as any;
        const startRow = Number(run?.startRow);
        const endRowExclusive = Number(run?.endRowExclusive);
        const styleId = Number(run?.styleId);
        if (!Number.isInteger(startRow) || !Number.isInteger(endRowExclusive) || !Number.isInteger(styleId)) return 0;
        if (row < startRow) hi = mid - 1;
        else if (row >= endRowExclusive) lo = mid + 1;
        else return styleId;
      }
      return 0;
    };

    const getStyleIdTupleKey = (row: number, col: number, cellStyleId: number): string | null => {
      // Prefer a public tuple helper if present (some controller implementations expose these).
      if (typeof docAny.getCellFormatStyleIds === "function") {
        try {
          coordScratch.row = row;
          coordScratch.col = col;
          const tuple = docAny.getCellFormatStyleIds(this.sheetId, coordScratch);
          if (Array.isArray(tuple) && tuple.length >= 4) {
            const normalized =
              tuple.length >= 5
                ? tuple.slice(0, 5).map(normalizeStyleId)
                : [tuple[0], tuple[1], tuple[2], tuple[3], 0].map(normalizeStyleId);
            return normalized.join(",");
          }
        } catch {
          // Ignore and fall back to the sheet model.
        }
      }

      // DocumentController layered-formatting storage (sheet/row/col/range-run/cell).
      if (sheetModel && typeof sheetModel === "object") {
        // These property names match the desktop DocumentController's internal model.
        // Keep legacy fallbacks (`sheetStyleId` / `sheetDefaultStyleId`, `rowStyles`, `colStyles`)
        // for older snapshots/adapters.
        const sheetDefaultStyleId = normalizeStyleId(
          (sheetModel as any).defaultStyleId ?? (sheetModel as any).sheetStyleId ?? (sheetModel as any).sheetDefaultStyleId
        );
        const rowStyleId = normalizeStyleId(
          (sheetModel as any).rowStyleIds?.get?.(row) ?? (sheetModel as any).rowStyles?.get?.(row)
        );
        const colStyleId = normalizeStyleId(
          (sheetModel as any).colStyleIds?.get?.(col) ?? (sheetModel as any).colStyles?.get?.(col)
        );
        const runsByCol = (sheetModel as any).formatRunsByCol;
        const runs =
          runsByCol && typeof runsByCol?.get === "function"
            ? runsByCol.get(col)
            : runsByCol && typeof runsByCol === "object"
              ? (runsByCol as any)[String(col)]
              : null;
        const runStyleId = styleIdForRowInRuns(runs, row);
        return [sheetDefaultStyleId, rowStyleId, colStyleId, normalizeStyleId(cellStyleId), normalizeStyleId(runStyleId)].join(",");
      }

      return null;
    };

    const cells: Array<Array<{ value: unknown; formula: string | null; styleId: number }>> = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      const outRow: Array<{ value: unknown; formula: string | null; styleId: number }> = [];
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        coordScratch.row = row;
        coordScratch.col = col;
        const cell = this.document.getCell(this.sheetId, coordScratch) as {
          value: unknown;
          formula: string | null;
          styleId: number;
        };
        const baseStyleId = normalizeStyleId(cell.styleId);

        const layerKey = getStyleIdTupleKey(row, col, baseStyleId);
        let styleId = layerKey ? styleIdByLayerKey.get(layerKey) : undefined;

        if (styleId === undefined) {
          // If everything is default, skip resolving/interning.
          if (layerKey === "0,0,0,0" || layerKey === "0,0,0,0,0") {
            styleId = 0;
          } else {
            const effectiveStyle = (() => {
              if (typeof docAny.getCellFormat === "function") return docAny.getCellFormat(this.sheetId, coordScratch);
              if (typeof docAny.getEffectiveCellStyle === "function") return docAny.getEffectiveCellStyle(this.sheetId, coordScratch);
              if (typeof docAny.getCellStyle === "function") return docAny.getCellStyle(this.sheetId, coordScratch);
              return this.document.styleTable.get(baseStyleId);
            })();

            styleId = this.document.styleTable.intern(effectiveStyle);
          }

          if (layerKey) styleIdByLayerKey.set(layerKey, styleId);
        }

        outRow.push({ value: cell.value ?? null, formula: cell.formula ?? null, styleId });
      }
      cells.push(outRow);
    }
    return cells;
  }

  private getClipboardCopyRange(): Range {
    const activeRange = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0];
    const activeCellFallback: Range = {
      startRow: this.selection.active.row,
      endRow: this.selection.active.row,
      startCol: this.selection.active.col,
      endCol: this.selection.active.col
    };

    if (!activeRange) return activeCellFallback;

    if (this.selection.type === "row" || this.selection.type === "column" || this.selection.type === "all") {
      const used = this.computeUsedRange();
      const clipped = used ? intersectRanges(activeRange, used) : null;
      if (clipped) return clipped;

      // If a user selects an entire row/column with no used-range overlap, still copy a
      // bounded slice (the selected row/column within the UI's grid limits). For an
      // entirely empty sheet (`all`), fall back to the active cell to avoid generating
      // a massive empty payload.
      if (this.selection.type === "all") return activeCellFallback;

      // Copying an entire (empty) Excel-scale column would allocate a huge 2D payload.
      // If there's no used-range overlap, treat it as a single-cell copy instead.
      if (this.selection.type === "column") return activeCellFallback;

      return activeRange;
    }

    return activeRange;
  }

  private async transcodeImageEntryToPng(entry: ImageEntry): Promise<Uint8Array | null> {
    if (entry.mimeType === "image/png") return entry.bytes;

    if (typeof document === "undefined") return null;

    const blob = new Blob([entry.bytes], { type: entry.mimeType || "application/octet-stream" });

    type Decoded = { source: CanvasImageSource; width: number; height: number };

    const decode = async (): Promise<Decoded | null> => {
      if (typeof createImageBitmap === "function") {
        try {
          const bitmap = await createImageBitmap(blob);
          return { source: bitmap, width: bitmap.width, height: bitmap.height };
        } catch {
          // Fall through to <img> decoding.
        }
      }

      if (typeof Image === "undefined" || typeof URL === "undefined") return null;

      const url = URL.createObjectURL(blob);
      try {
        const img = new Image();
        img.decoding = "async";
        const loaded = new Promise<void>((resolve, reject) => {
          img.onload = () => resolve();
          img.onerror = () => reject(new Error("image decode failed"));
        });
        img.src = url;
        await loaded;
        const width = (img as any).naturalWidth ?? img.width;
        const height = (img as any).naturalHeight ?? img.height;
        if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return null;
        return { source: img, width, height };
      } catch {
        return null;
      } finally {
        try {
          URL.revokeObjectURL(url);
        } catch {
          // ignore
        }
      }
    };

    const decoded = await decode();
    if (!decoded) return null;

    const canvas = document.createElement("canvas");
    canvas.width = decoded.width;
    canvas.height = decoded.height;
    const ctx = canvas.getContext("2d");
    if (!ctx) return null;
    ctx.drawImage(decoded.source, 0, 0);

    if (typeof canvas.toBlob === "function") {
      const pngBlob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, "image/png"));
      if (!pngBlob) return null;
      const buf = await pngBlob.arrayBuffer();
      return new Uint8Array(buf);
    }

    if (typeof canvas.toDataURL === "function" && typeof atob === "function") {
      try {
        const url = canvas.toDataURL("image/png");
        const comma = url.indexOf(",");
        if (comma === -1) return null;
        const base64 = url.slice(comma + 1);
        const binary = atob(base64);
        const out = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i += 1) {
          out[i] = binary.charCodeAt(i);
        }
        return out;
      } catch {
        return null;
      }
    }

    return null;
  }

  private async copySelectedDrawingToClipboard(): Promise<void> {
    const selectedId = this.selectedDrawingId;
    if (selectedId == null) return;

    const objects = this.listDrawingObjectsForSheet(this.sheetId);
    const selected = objects.find((obj) => obj.id === selectedId);
    if (!selected || selected.kind.type !== "image") {
      throw new Error("Selected drawing is not an image");
    }

    const dlp = this.dlpContext;
    if (dlp) {
      const fallback = this.selection.active;
      const anchor: any = selected.anchor;
      const fromCell =
        anchor && (anchor.type === "oneCell" || anchor.type === "twoCell") ? anchor.from?.cell : null;
      const row = Number(fromCell?.row);
      const col = Number(fromCell?.col);
      const cell =
        Number.isInteger(row) && row >= 0 && Number.isInteger(col) && col >= 0 ? { row, col } : fallback;
      enforceClipboardCopy({
        documentId: dlp.documentId,
        sheetId: this.sheetId,
        range: { start: { row: cell.row, col: cell.col }, end: { row: cell.row, col: cell.col } },
        classificationStore: dlp.classificationStore,
        policy: dlp.policy
      });
    }

    const imageId = selected.kind.imageId;
    let entry = this.drawingImages.get(imageId);
    const getAsync = (this.drawingImages as any)?.getAsync;
    if (!entry && typeof getAsync === "function") {
      try {
        entry = await getAsync.call(this.drawingImages, imageId);
      } catch {
        entry = undefined;
      }
    }
    if (!entry) {
      throw new Error("Selected drawing image data not found");
    }

    const pngBytes = await this.transcodeImageEntryToPng(entry);
    if (!pngBytes) {
      try {
        showToast("Copy picture not supported for this image type", "warning");
      } catch {
        // `showToast` requires a #toast-root; unit tests don't always include it.
      }
      throw new Error("Copy picture not supported for this image type");
    }

    const provider = await this.getClipboardProvider();
    await provider.write({ text: "", imagePng: pngBytes });
    // The system clipboard now contains an image; clear any stale "internal range copy" context.
    this.clipboardCopyContext = null;
  }

  private async cutSelectedDrawingToClipboard(): Promise<void> {
    const selectedId = this.selectedDrawingId;
    if (selectedId == null) return;

    await this.copySelectedDrawingToClipboard();

    const label = (() => {
      const translated = t("clipboard.cut");
      return translated === "clipboard.cut" ? "Cut" : translated;
    })();

    const sheetId = this.sheetId;
    const selected = this.listDrawingObjectsForSheet(sheetId).find((obj) => obj.id === selectedId) ?? null;
    const imageId = selected?.kind.type === "image" ? selected.kind.imageId : null;

    const docAny: any = this.document as any;
    const deleteDrawing =
      typeof docAny.deleteDrawing === "function"
        ? (docAny.deleteDrawing as (sheetId: string, drawingId: string | number, options?: unknown) => void)
        : null;
    const getSheetDrawings =
      typeof docAny.getSheetDrawings === "function" ? (docAny.getSheetDrawings as (sheetId: string) => unknown) : null;
    const deleteImage =
      typeof docAny.deleteImage === "function" ? (docAny.deleteImage as (imageId: string, options?: unknown) => void) : null;

    // `DrawingObject.id` is a UI-only numeric id (stable hash for some sources). When the underlying
    // DocumentController drawing ids are non-numeric strings, deleting by the UI id would be a no-op.
    // Map from UI id -> raw drawing ids so cut behaves like the Delete shortcut.
    const rawIdsToDelete = new Set<string | number>();
    if (getSheetDrawings) {
      let raw: unknown = null;
      try {
        raw = getSheetDrawings.call(docAny, sheetId);
      } catch {
        raw = null;
      }
      if (Array.isArray(raw)) {
        for (const entry of raw) {
          if (!entry || typeof entry !== "object") continue;
          let uiId: number | null = null;
          try {
            uiId = convertDocumentSheetDrawingsToUiDrawingObjects([entry], { sheetId })[0]?.id ?? null;
          } catch {
            uiId = null;
          }
          if (uiId !== selectedId) continue;
          const rawId = (entry as any).id;
          if (typeof rawId === "string") {
            const trimmed = rawId.trim();
            if (trimmed) rawIdsToDelete.add(trimmed);
          } else if (typeof rawId === "number" && Number.isFinite(rawId)) {
            rawIdsToDelete.add(rawId);
          }
        }
      }
    }
    if (rawIdsToDelete.size === 0) rawIdsToDelete.add(selectedId);

    this.document.beginBatch({ label });
    try {
      if (deleteDrawing) {
        for (const rawId of rawIdsToDelete) {
          try {
            deleteDrawing.call(docAny, sheetId, rawId, { label });
          } catch {
            // ignore
          }
        }
      } else if (typeof docAny.setSheetDrawings === "function" && typeof docAny.getSheetDrawings === "function") {
        const existing = docAny.getSheetDrawings(sheetId);
        const ids = new Set(Array.from(rawIdsToDelete, (id) => String(id)));
        const next = Array.isArray(existing)
          ? existing.filter((d: any) => !ids.has(String(d?.id ?? "")))
          : [];
        docAny.setSheetDrawings(sheetId, next, { label });
      }

      if (imageId && deleteImage && !this.isImageReferencedByAnyDrawing(imageId)) {
        try {
          deleteImage.call(docAny, imageId, { label });
        } catch {
          // ignore
        }
        this.drawingOverlay.invalidateImage(imageId);
      }
    } finally {
      this.document.endBatch();
    }

    this.selectedDrawingId = null;
    this.dispatchDrawingSelectionChanged();
    this.refresh();
    this.focus();
  }

  private async copySelectionToClipboard(): Promise<void> {
    try {
      if (this.selectedDrawingId != null) {
        await this.copySelectedDrawingToClipboard();
        return;
      }

      const range = this.getClipboardCopyRange();
      const rowCount = Math.max(0, range.endRow - range.startRow + 1);
      const colCount = Math.max(0, range.endCol - range.startCol + 1);
      const cellCount = rowCount * colCount;
      if (cellCount > MAX_CLIPBOARD_CELLS) {
        try {
          showToast(
            `Selection too large to copy (>${MAX_CLIPBOARD_CELLS.toLocaleString()} cells). Select fewer cells and try again.`,
            "warning"
          );
        } catch {
          // `showToast` requires a #toast-root; unit tests don't always include it.
        }
        return;
      }
      const cellRange = {
        start: { row: range.startRow, col: range.startCol },
        end: { row: range.endRow, col: range.endCol }
      };

      const dlp = this.dlpContext;
      if (dlp) {
        enforceClipboardCopy({
          documentId: dlp.documentId,
          sheetId: this.sheetId,
          range: cellRange,
          classificationStore: dlp.classificationStore,
          policy: dlp.policy
        });
      }

      // Build an Excel-compatible payload:
       // - `text/plain` should contain *display values* (including computed formula results).
       // - `text/html` can include both display values (cell text) and formulas (data-formula attr)
       //   so spreadsheet-to-spreadsheet pastes preserve formulas.
       const grid = getCellGridFromRange(this.document, this.sheetId, cellRange) as any[][];
       const coordScratch = { row: 0, col: 0 };
       const baseRow = cellRange.start.row;
       const baseCol = cellRange.start.col;
       for (let r = 0; r < grid.length; r += 1) {
         const row = grid[r] ?? [];
         for (let c = 0; c < row.length; c += 1) {
           const cell = row[c];
           if (!cell || cell.formula == null) continue;
          // When copying formulas, the clipboard payload should include the displayed value
          // (including the computed result of the formula). If the computed value is `null`,
           // treat it as an empty cell so plain-text consumers (and Paste Values) don't fall
           // back to copying the formula text.
           coordScratch.row = baseRow + r;
           coordScratch.col = baseCol + c;
           const computed = this.getCellComputedValue(coordScratch) as any;
           cell.value = computed ?? "";
         }
       }
       const payload = serializeCellGridToClipboardPayload(grid as any);
      const cells = this.snapshotClipboardCells(range);
      const provider = await this.getClipboardProvider();
      await provider.write(payload);
      this.clipboardCopyContext = { range, payload, cells };
    } catch (err) {
      const isDlpViolation = err instanceof DlpViolationError || (err as any)?.name === "DlpViolationError";
      if (isDlpViolation) {
        try {
          const message =
            typeof (err as any)?.message === "string" && (err as any).message.trim()
              ? String((err as any).message)
              : "Copy blocked by data loss prevention policy.";
          // Blocking copy/cut is expected under strict DLP policies; present this as a warning
          // (rather than an "error") so it reads as a policy restriction instead of a crash.
          showToast(message, "warning");
        } catch {
          // `showToast` requires a #toast-root; unit tests don't always include it.
        }
        return;
      }
      // Ignore clipboard failures (permissions, platform restrictions).
    }
  }

  private async pasteClipboardImageAsDrawing(content: unknown): Promise<boolean> {
    const maxBytes = Number(CLIPBOARD_LIMITS?.maxImageBytes) > 0 ? Number(CLIPBOARD_LIMITS.maxImageBytes) : 5 * 1024 * 1024;
    const mb = Math.round(maxBytes / 1024 / 1024);
    const anyContent = content as any;

    // Clipboard provider drops oversized images to avoid huge allocations; it also sets a
    // non-enumerable marker so we can show user feedback here.
    if (anyContent?.skippedOversizedImagePng === true) {
      try {
        showToast(`Image too large (>${mb}MB). Choose a smaller file.`, "warning");
      } catch {
        // `showToast` requires a #toast-root; unit tests don't always include it.
      }
      return true;
    }

    const direct = anyContent?.imagePng;
    const base64 = anyContent?.pngBase64;

    const bytes: Uint8Array | null = (() => {
      if (direct instanceof Uint8Array && direct.byteLength > 0) return direct;
      if (typeof base64 === "string" && base64.trim() !== "") {
        return decodeClipboardImageBase64ToBytes(base64, { maxBytes });
      }
      return null;
    })();

    if (!bytes) {
      const hadImageHint = direct != null || (typeof base64 === "string" && base64.trim() !== "");
      if (hadImageHint) {
        try {
          showToast("Unable to paste image from clipboard. Try copying the image again.", "warning");
        } catch {
          // `showToast` requires a #toast-root; unit tests don't always include it.
        }
        return true;
      }
      return false;
    }

    if (bytes.byteLength > maxBytes) {
      try {
        showToast(`Image too large (>${mb}MB). Choose a smaller file.`, "warning");
      } catch {
        // `showToast` requires a #toast-root; unit tests don't always include it.
      }
      return true;
    }

    // Guard against PNG decompression bombs: small compressed bytes can still decode into huge bitmaps.
    const dims = readPngDimensions(bytes);
    if (dims) {
      const MAX_DIMENSION = 10_000;
      const MAX_PIXELS = 50_000_000;
      if (dims.width > MAX_DIMENSION || dims.height > MAX_DIMENSION || dims.width * dims.height > MAX_PIXELS) {
        try {
          showToast("Image too large to paste.", "warning");
        } catch {
          // `showToast` requires a #toast-root; unit tests don't always include it.
        }
        return true;
      }
    }

    const docAny = this.document as any;
    if (typeof docAny.insertDrawing !== "function") {
      return false;
    }

    const uuid = (): string => {
      const randomUuid = (globalThis as any).crypto?.randomUUID as (() => string) | undefined;
      if (typeof randomUuid === "function") {
        try {
          return randomUuid.call((globalThis as any).crypto);
        } catch {
          // Fall through to pseudo-random below.
        }
      }
      return `${Date.now().toString(16)}_${Math.random().toString(16).slice(2)}`;
    };

    const DEFAULT_PICTURE_WIDTH_COLS = 5;
    const DEFAULT_PICTURE_HEIGHT_ROWS = 11;

    const existingDrawings = (() => {
      try {
        const raw = (this.document as any).getSheetDrawings?.(this.sheetId);
        return Array.isArray(raw) ? raw : [];
      } catch {
        return [];
      }
    })();
    let nextZOrder = existingDrawings.length;
    for (const raw of existingDrawings) {
      const maybe = raw as any;
      const z = Number(maybe?.zOrder ?? maybe?.z_order);
      if (Number.isFinite(z) && z >= nextZOrder) nextZOrder = z + 1;
    }

    const base = this.selection.active;
    const startRow = base.row;
    const startCol = base.col;

    const imageId = `image_${uuid()}.png`;
    const drawingId = createDrawingObjectId();
    const drawing = {
      // Store as a string to keep drawing ids JSON-friendly and stable across JS↔Rust↔Yjs hops.
      // The UI adapters normalize ids back to numbers for rendering/interaction.
      id: String(drawingId),
      kind: { type: "image", imageId },
      anchor: {
        type: "twoCell",
        from: { cell: { row: startRow, col: startCol }, offset: { xEmu: 0, yEmu: 0 } },
        to: {
          cell: { row: startRow + DEFAULT_PICTURE_HEIGHT_ROWS, col: startCol + DEFAULT_PICTURE_WIDTH_COLS },
          offset: { xEmu: 0, yEmu: 0 },
        },
      },
      zOrder: nextZOrder,
    };

    this.document.beginBatch({ label: "Paste Picture" });
    try {
      const imageEntry: ImageEntry = { id: imageId, bytes, mimeType: "image/png" };
      // Persist picture bytes out-of-band (IndexedDB) so they survive reloads without
      // bloating DocumentController snapshot payloads.
      this.drawingImages.set(imageEntry);
      // Preload the bitmap so the first overlay render can reuse the decode promise.
      void this.drawingOverlay.preloadImage(imageEntry).catch(() => {
        // ignore
      });
      try {
        this.imageBytesBinder?.onLocalImageInserted(imageEntry);
      } catch {
        // Best-effort: never fail paste due to collab image propagation.
      }
      docAny.insertDrawing(this.sheetId, drawing);
      this.document.endBatch();
    } catch (err) {
      this.document.cancelBatch();
      throw err;
    }

    this.drawingObjectsCache = null;
    const prevSelected = this.selectedDrawingId;
    this.selectedDrawingId = drawingId;
    this.drawingOverlay.setSelectedId(drawingId);
    this.drawingInteractionController?.setSelectedId(drawingId);
    if (prevSelected !== drawingId) {
      this.dispatchDrawingSelectionChanged();
    }
    this.renderDrawings(this.sharedGrid ? this.sharedGrid.renderer.scroll.getViewportState() : undefined);
    this.focus();
    return true;
  }

  async pasteClipboardToSelection(
    options: { mode?: "all" | "values" | "formulas" | "formats"; transpose?: boolean } = {}
  ): Promise<void> {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    try {
      const provider = await this.getClipboardProvider();
      const content = await provider.read();
      const start = { ...this.selection.active };
      const ctx = this.clipboardCopyContext;
      const mode = options.mode ?? "all";
      const transpose = options.transpose === true;
      let deltaRow = 0;
      let deltaCol = 0;

      const { isInternalPaste, nextContext } = reconcileClipboardCopyContextForPaste(ctx, content);
      this.clipboardCopyContext = nextContext;

      const externalGrid = mode === "all" && isInternalPaste ? null : parseClipboardContentToCellGrid(content);
      const internalCells = isInternalPaste ? ctx?.cells : null;
      const rowCount = internalCells ? internalCells.length : externalGrid?.length ?? 0;
      // Avoid `Math.max(0, ...rows.map(...))` spread: a tall paste can contain tens of thousands of
      // rows, which would exceed JS engines' argument limits.
      let colCount = 0;
      if (internalCells) {
        for (const row of internalCells) {
          const len = Array.isArray(row) ? row.length : 0;
          if (len > colCount) colCount = len;
        }
      }
      if (externalGrid) {
        for (const row of externalGrid) {
          const len = Array.isArray(row) ? row.length : 0;
          if (len > colCount) colCount = len;
        }
      }

      // Some clipboard backends provide `text/plain=""` alongside image data. Treat a purely-empty
      // parsed grid as "no usable tabular text" so image-only clipboard payloads paste as floating
      // pictures (Excel-like) rather than clearing a single cell.
      if (mode === "all" && !internalCells && externalGrid) {
        const anyContent = content as any;
        const hasImage =
          anyContent?.skippedOversizedImagePng === true ||
          (anyContent?.imagePng instanceof Uint8Array && anyContent.imagePng.byteLength > 0) ||
          (typeof anyContent?.pngBase64 === "string" && anyContent.pngBase64.trim() !== "");
        if (hasImage) {
          let hasMeaningfulCell = false;
          for (const row of externalGrid) {
            if (!Array.isArray(row) || row.length === 0) continue;
            for (const cell of row) {
              if (!cell || typeof cell !== "object") continue;
              const anyCell = cell as any;
              const formula = anyCell.formula;
              if (typeof formula === "string" && formula.trim() !== "") {
                hasMeaningfulCell = true;
                break;
              }
              const value = anyCell.value;
              if (value != null) {
                if (typeof value !== "string" || value.trim() !== "") {
                  hasMeaningfulCell = true;
                  break;
                }
              }
              const format = anyCell.format;
              if (format && typeof format === "object" && Object.keys(format).length > 0) {
                hasMeaningfulCell = true;
                break;
              }
            }
            if (hasMeaningfulCell) break;
          }
          if (!hasMeaningfulCell) {
            const handled = await this.pasteClipboardImageAsDrawing(content);
            if (handled) return;
          }
        }
      }
      if (rowCount === 0 || colCount === 0) {
        // Excel-style behavior: if the clipboard only contains an image (and no
        // tabular text/HTML), paste it as a floating picture anchored at the
        // active cell.
        if (mode === "all" && !internalCells) {
          const handled = await this.pasteClipboardImageAsDrawing(content);
          if (handled) return;
        }
        return;
      }

      const pastedCellCount = rowCount * colCount;
      if (pastedCellCount > MAX_CLIPBOARD_CELLS) {
        try {
          showToast(
            `Paste too large (>${MAX_CLIPBOARD_CELLS.toLocaleString()} cells). Paste fewer cells and try again.`,
            "warning"
          );
        } catch {
          // `showToast` requires a #toast-root; unit tests don't always include it.
        }
        return;
      }

      if (
        isInternalPaste
      ) {
        deltaRow = start.row - (ctx as any).range.startRow;
        deltaCol = start.col - (ctx as any).range.startCol;
      }

      const rewrittenInternalFormulas = new Map<number, string>();
      if (
        isInternalPaste &&
        internalCells &&
        (deltaRow !== 0 || deltaCol !== 0) &&
        (mode === "all" || mode === "formulas")
      ) {
        const requests: Array<{ formula: string; deltaRow: number; deltaCol: number }> = [];
        const keys: number[] = [];
        for (let r = 0; r < internalCells.length; r++) {
          const row = internalCells[r] ?? [];
          for (let c = 0; c < row.length; c++) {
            const rawFormula = row[c]?.formula;
            if (typeof rawFormula !== "string") continue;
            requests.push({ formula: rawFormula, deltaRow, deltaCol });
            keys.push(r * colCount + c);
          }
        }

        if (requests.length > 0 && this.wasmEngine) {
          try {
            const rewritten = await this.wasmEngine.rewriteFormulasForCopyDelta(requests);
            if (Array.isArray(rewritten) && rewritten.length === requests.length) {
              for (let i = 0; i < rewritten.length; i++) {
                const key = keys[i];
                if (typeof key !== "number") continue;
                const next = rewritten[i];
                if (typeof next === "string") {
                  rewrittenInternalFormulas.set(key, next);
                }
              }
            }
          } catch {
            // Ignore and fall back to best-effort shifting.
          }
        }
      }

      const values = (() => {
        if (transpose) {
          const srcRowCount = rowCount;
          const srcColCount = colCount;

          const makeTransposedGrid = (cellBuilder: (srcRow: number, srcCol: number) => any): any[][] => {
            const out: any[][] = [];
            for (let dstRow = 0; dstRow < srcColCount; dstRow += 1) {
              const outRow: any[] = [];
              for (let dstCol = 0; dstCol < srcRowCount; dstCol += 1) {
                // Transpose mapping: (srcRow, srcCol) -> (dstRow=srcCol, dstCol=srcRow)
                outRow.push(cellBuilder(dstCol, dstRow));
              }
              out.push(outRow);
            }
            return out;
          };

          const shiftInternalFormulaForTranspose = (rawFormula: string, srcRow: number, srcCol: number): string => {
            // Each source cell moves from:
            //  (srcStartRow + srcRow, srcStartCol + srcCol)
            // to (dstStartRow + srcCol, dstStartCol + srcRow)
            // so the required relative reference shift depends on the cell's original offset.
            const deltaRowForCell = deltaRow + srcCol - srcRow;
            const deltaColForCell = deltaCol + srcRow - srcCol;
            if (deltaRowForCell === 0 && deltaColForCell === 0) return rawFormula;
            return shiftA1References(rawFormula, deltaRowForCell, deltaColForCell);
          };

          if (mode === "all") {
            if (isInternalPaste) {
              return makeTransposedGrid((srcRow, srcCol) => {
                const cell = internalCells?.[srcRow]?.[srcCol];
                const rawFormula = cell?.formula ?? null;
                const formula =
                  rawFormula != null ? shiftInternalFormulaForTranspose(rawFormula, srcRow, srcCol) : null;
                if (formula != null) {
                  return { formula, styleId: cell?.styleId ?? 0 };
                }
                return { value: cell?.value ?? null, styleId: cell?.styleId ?? 0 };
              });
            }

            return makeTransposedGrid((srcRow, srcCol) => {
              const cell: any = externalGrid?.[srcRow]?.[srcCol] ?? null;
              const format = clipboardFormatToDocStyle(cell?.format ?? null);
              if (cell?.formula != null) {
                return { formula: cell.formula, format };
              }
              return { value: cell?.value ?? null, format };
            });
          }

          if (mode === "formats") {
            if (isInternalPaste) {
              return makeTransposedGrid((srcRow, srcCol) => {
                const cell = internalCells?.[srcRow]?.[srcCol];
                return { styleId: cell?.styleId ?? 0 };
              });
            }

            return makeTransposedGrid((srcRow, srcCol) => {
              const cell: any = externalGrid?.[srcRow]?.[srcCol] ?? null;
              return { format: clipboardFormatToDocStyle(cell?.format ?? null) };
            });
          }

          if (mode === "formulas") {
            if (isInternalPaste) {
              return makeTransposedGrid((srcRow, srcCol) => {
                const cell = internalCells?.[srcRow]?.[srcCol];
                const rawFormula = cell?.formula ?? null;
                const formula =
                  rawFormula != null ? shiftInternalFormulaForTranspose(rawFormula, srcRow, srcCol) : null;
                if (formula != null) return { formula };
                return { value: cell?.value ?? null };
              });
            }

            return makeTransposedGrid((srcRow, srcCol) => {
              const cell: any = externalGrid?.[srcRow]?.[srcCol] ?? null;
              if (cell?.formula != null) return { formula: cell.formula };
              return { value: cell?.value ?? null };
            });
          }

          // mode === "values"
          const source = externalGrid ?? (isInternalPaste ? (internalCells as any) : null);
          if (!source) return [];
          return makeTransposedGrid((srcRow, srcCol) => {
            const cell: any = source?.[srcRow]?.[srcCol] ?? null;
            return { value: cell?.value ?? null };
          });
        }

        if (mode === "all") {
          if (isInternalPaste) {
            return internalCells!.map((row, r) =>
              row.map((cell, c) => {
                const rawFormula = cell.formula;
                const formula =
                  typeof rawFormula === "string" && (deltaRow !== 0 || deltaCol !== 0)
                    ? (rewrittenInternalFormulas.get(r * colCount + c) ?? shiftA1References(rawFormula, deltaRow, deltaCol))
                    : rawFormula;
                if (formula != null) {
                  return { formula, styleId: cell.styleId };
                }
                return { value: cell.value ?? null, styleId: cell.styleId };
              })
            );
          }

          return externalGrid!.map((row) =>
            row.map((cell: any) => {
              const format = clipboardFormatToDocStyle(cell.format ?? null);
              const rawFormula = cell.formula;
              const formula =
                rawFormula != null && (deltaRow !== 0 || deltaCol !== 0)
                  ? shiftA1References(rawFormula, deltaRow, deltaCol)
                  : rawFormula;

              if (formula != null) {
                return { formula, format };
              }

              return { value: cell.value ?? null, format };
            })
          );
        }

        if (mode === "formats") {
          if (isInternalPaste) {
            return internalCells!.map((row) => row.map((cell) => ({ styleId: cell.styleId })));
          }
          return externalGrid!.map((row) =>
            row.map((cell: any) => ({ format: clipboardFormatToDocStyle(cell.format ?? null) }))
          );
        }

        if (mode === "formulas") {
          if (isInternalPaste) {
            return internalCells!.map((row, r) =>
              row.map((cell, c) => {
                const rawFormula = cell.formula;
                const formula =
                  typeof rawFormula === "string" && (deltaRow !== 0 || deltaCol !== 0)
                    ? (rewrittenInternalFormulas.get(r * colCount + c) ?? shiftA1References(rawFormula, deltaRow, deltaCol))
                    : rawFormula;
                if (formula != null) return { formula };
                return { value: cell.value ?? null };
              })
            );
          }
          return externalGrid!.map((row) =>
            row.map((cell: any) => {
              if (cell.formula != null) return { formula: cell.formula };
              return { value: cell.value ?? null };
            })
          );
        }

        // mode === "values"
        const source = externalGrid ?? (isInternalPaste ? (internalCells as any) : null);
        if (!source) return [];
        return source.map((row: any[]) => row.map((cell: any) => ({ value: cell?.value ?? null })));
      })();

      this.document.setRangeValues(this.sheetId, start, values, { label: t("clipboard.paste") });

      const pastedRowCount = values.length;
      const pastedColCount = Math.max(0, ...values.map((row: any) => (Array.isArray(row) ? row.length : 0)));
      if (pastedRowCount === 0 || pastedColCount === 0) return;

      const range: Range = {
        startRow: start.row,
        endRow: start.row + pastedRowCount - 1,
        startCol: start.col,
        endCol: start.col + pastedColCount - 1
      };
      this.selection = buildSelection({ ranges: [range], active: start, anchor: start, activeRangeIndex: 0 }, this.limits);

      this.syncEngineNow();
      this.refresh();
      this.focus();
    } catch (err) {
      const isDlpViolation = err instanceof DlpViolationError || (err as any)?.name === "DlpViolationError";
      if (isDlpViolation) {
        try {
          const message = typeof (err as any)?.message === "string" ? String((err as any).message) : "Paste blocked by policy.";
          showToast(message, "warning");
        } catch {
          // `showToast` requires a #toast-root; unit tests don't always include it.
        }
        return;
      }
      const isClipboardLimit = (err as any)?.name === "ClipboardParseLimitError";
      if (isClipboardLimit) {
        try {
          showToast(
            `Paste too large (>${MAX_CLIPBOARD_CELLS.toLocaleString()} cells). Paste fewer cells and try again.`,
            "warning"
          );
        } catch {
          // `showToast` requires a #toast-root; unit tests don't always include it.
        }
        return;
      }
      // Ignore clipboard failures (permissions, platform restrictions).
    }
  }

  private async cutSelectionToClipboard(): Promise<void> {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    try {
      if (this.selectedDrawingId != null) {
        await this.cutSelectedDrawingToClipboard();
        return;
      }

      const range = this.getClipboardCopyRange();
      const rowCount = Math.max(0, range.endRow - range.startRow + 1);
      const colCount = Math.max(0, range.endCol - range.startCol + 1);
      const cellCount = rowCount * colCount;
      if (cellCount > MAX_CLIPBOARD_CELLS) {
        try {
          showToast(
            `Selection too large to cut (>${MAX_CLIPBOARD_CELLS.toLocaleString()} cells). Select fewer cells and try again.`,
            "warning"
          );
        } catch {
          // `showToast` requires a #toast-root; unit tests don't always include it.
        }
        return;
      }
      const cellRange = {
        start: { row: range.startRow, col: range.startCol },
        end: { row: range.endRow, col: range.endCol }
      };

      const dlp = this.dlpContext;
      if (dlp) {
        enforceClipboardCopy({
          documentId: dlp.documentId,
          sheetId: this.sheetId,
          range: cellRange,
          classificationStore: dlp.classificationStore,
          policy: dlp.policy
        });
      }

       const grid = getCellGridFromRange(this.document, this.sheetId, cellRange) as any[][];
       const coordScratch = { row: 0, col: 0 };
       const baseRow = cellRange.start.row;
       const baseCol = cellRange.start.col;
       for (let r = 0; r < grid.length; r += 1) {
         const row = grid[r] ?? [];
         for (let c = 0; c < row.length; c += 1) {
           const cell = row[c];
           if (!cell || cell.formula == null) continue;
           coordScratch.row = baseRow + r;
           coordScratch.col = baseCol + c;
           const computed = this.getCellComputedValue(coordScratch) as any;
           cell.value = computed ?? "";
         }
       }
      const payload = serializeCellGridToClipboardPayload(grid as any);
      const cells = this.snapshotClipboardCells(range);
      const provider = await this.getClipboardProvider();
      await provider.write(payload);
      this.clipboardCopyContext = { range, payload, cells };

      const label = (() => {
        const translated = t("clipboard.cut");
        return translated === "clipboard.cut" ? "Cut" : translated;
      })();

      this.document.beginBatch({ label });
      this.document.clearRange(
        this.sheetId,
        {
          start: { row: range.startRow, col: range.startCol },
          end: { row: range.endRow, col: range.endCol }
        },
        { label }
      );
      this.document.endBatch();

      this.syncEngineNow();
      this.refresh();
      this.focus();
    } catch (err) {
      const isDlpViolation = err instanceof DlpViolationError || (err as any)?.name === "DlpViolationError";
      if (isDlpViolation) {
        try {
          const message =
            typeof (err as any)?.message === "string" && (err as any).message.trim()
              ? String((err as any).message)
              : "Cut blocked by data loss prevention policy.";
          // See `copySelectionToClipboard` for rationale.
          showToast(message, "warning");
        } catch {
          // Best-effort: if the toast UI isn't mounted, don't crash clipboard actions.
        }
      }

      // Ignore clipboard failures (permissions, platform restrictions).
    }
  }

  private getCellDisplayValue(cell: CellCoord): string {
    const value = this.getCellComputedValue(cell);
    if (value == null) return "";
    return this.formatCellValueForDisplay(cell, value);
  }

  private formatCellValueForDisplay(cell: CellCoord, value: SpreadsheetValue): string {
    if (value == null) return "";
    if (typeof value === "number" && Number.isFinite(value)) {
      const docStyle: any = this.document.getCellFormat(this.sheetId, cell);
      const rawNumberFormat = docStyle?.numberFormat ?? docStyle?.number_format;
      const numberFormat = typeof rawNumberFormat === "string" && rawNumberFormat.trim() !== "" ? rawNumberFormat : null;
      if (numberFormat) return formatValueWithNumberFormat(value, numberFormat);
    }
    return String(value);
  }

  private getCellInputText(cell: CellCoord): string {
    const state = this.document.getCell(this.sheetId, cell) as { value: unknown; formula: string | null };
    if (state?.formula != null) {
      return state.formula;
    }
    if (isRichTextValue(state?.value)) return state.value.text;
    if (state?.value != null) return String(state.value);
    return "";
  }

  getCellComputedValueForSheet(sheetId: string, cell: { row: number; col: number }): string | number | boolean | null {
    return this.getCellComputedValueForSheetInternal(sheetId, cell);
  }

  private getCellComputedValue(cell: CellCoord): SpreadsheetValue {
    return this.getCellComputedValueForSheetInternal(this.sheetId, cell);
  }

  private getCellComputedValueForSheetInternal(sheetId: string, cell: CellCoord): SpreadsheetValue {
    // The WASM engine currently cannot resolve sheet-qualified references (e.g. `Sheet2!A1`),
    // so when multiple sheets exist we fall back to the in-process evaluator for *all* formulas
    // to keep dependent values consistent.
    // Hot path: avoid allocating a fresh `string[]` on every render-time lookup.
    const sheetCount = (this.document as any)?.model?.sheets?.size;
    const useEngineCache = (typeof sheetCount === "number" ? sheetCount : this.document.getSheetIds().length) <= 1;
    if (useEngineCache) {
      // Hot path: shared-grid rendering calls this for *every* formula cell in view.
      // Avoid allocating A1/key strings on cache hits by using a numeric `{row,col}` cache.
      const sheetCache = this.getComputedValuesByCoordForSheet(sheetId);
      if (sheetCache && cell.col >= 0 && cell.col < COMPUTED_COORD_COL_STRIDE && cell.row >= 0) {
        const key = cell.row * COMPUTED_COORD_COL_STRIDE + cell.col;
        const cached = sheetCache.get(key);
        // `computedValuesByCoord` never stores `undefined`; a missing entry always returns undefined.
        if (cached !== undefined) return cached;
      }
    }

    const memo = new Map<string, Map<number, SpreadsheetValue>>();
    const stack = new Map<string, Set<number>>();
    return this.computeCellValue(sheetId, cell, memo, stack, { useEngineCache });
  }

  private resolveSheetIdByName(name: string): string | null {
    const trimmed = (() => {
      const raw = name.trim();
      // Sheet-qualified references can be quoted using Excel syntax: `'My Sheet'!A1`.
      // When we split on "!" we receive the quoted token (`'My Sheet'`) and need to
      // unquote it before resolving the display name -> stable id mapping.
      const quoted = /^'((?:[^']|'')+)'$/.exec(raw);
      if (quoted) return quoted[1]!.replace(/''/g, "'").trim();
      return raw;
    })();
    if (!trimmed) return null;

    const resolved = this.sheetNameResolver?.getSheetIdByName(trimmed);
    if (resolved) return resolved;

    // Allow stable sheet ids to pass through when they are known to the resolver.
    // This avoids treating real (but currently empty/unmaterialized) sheets as
    // "unknown" just because the DocumentController hasn't created the sheet yet.
    if (this.sheetNameResolver?.getSheetNameById(trimmed)) return trimmed;

    // Fallback to sheet ids to keep legacy formulas (and test fixtures) working.
    // Avoid allocating a fresh array via `getSheetIds()` in hot paths (formula evaluation).
    const sheets: Map<string, unknown> | undefined = (this.document as any)?.model?.sheets;
    if (sheets && typeof sheets.keys === "function") {
      if (sheets.has(trimmed)) return trimmed;
      const lower = trimmed.toLowerCase();
      for (const id of sheets.keys()) {
        if (typeof id === "string" && id.toLowerCase() === lower) return id;
      }
      return null;
    }

    const knownSheets = this.document.getSheetIds();
    const lower = trimmed.toLowerCase();
    return knownSheets.find((id) => id.toLowerCase() === lower) ?? null;
  }

  private evaluateFormulaBarArgumentPreview(expr: string): SpreadsheetValue | string {
    const raw = typeof expr === "string" ? expr : String(expr ?? "");
    const trimmedExpr = raw.trim();
    if (!trimmedExpr) return "(preview unavailable)";

    const editTarget = this.formulaEditCell ?? { sheetId: this.sheetId, cell: { ...this.selection.active } };
    const sheetId = editTarget.sheetId;
    const cellAddress = cellToA1(editTarget.cell);

    // Hard cap on the number of cell reads we allow for preview. This keeps the formula bar
    // responsive even when the argument expression references a large range.
    const MAX_CELL_READS = 5_000;
    const knownSheets =
      typeof this.document.getSheetIds === "function"
        ? (this.document.getSheetIds() as string[]).filter((s) => typeof s === "string" && s.length > 0)
        : [];
    const sheetExistsCache = new Map<string, boolean>();
    const sheetExists = (id: string): boolean => {
      const key = String(id ?? "").trim();
      if (!key) return false;
      const lower = key.toLowerCase();
      // The current sheet is always "known" for preview purposes, even if the DocumentController
      // hasn't materialized it yet (e.g. very early startup or in minimal unit tests).
      if (lower === sheetId.toLowerCase()) return true;
      const cached = sheetExistsCache.get(lower);
      if (cached !== undefined) return cached;
      let exists = false;
      try {
        // Prefer `getSheetMeta` because it checks both materialized sheets and
        // sheet metadata entries without creating sheets.
        if (typeof (this.document as any).getSheetMeta === "function") {
          exists = Boolean((this.document as any).getSheetMeta(key));
        } else {
          exists = knownSheets.some((s) => s.toLowerCase() === lower);
        }
      } catch {
        exists = false;
      }
      sheetExistsCache.set(lower, exists);
      return exists;
    };

    // Resolve named ranges (and allow undefined names to fall back to `#NAME?`).
    const resolveNameToReference = (name: string): string | null => {
      const key = String(name ?? "").trim().toUpperCase();
      if (!key) return null;

      for (const entry of this.searchWorkbook.names.values()) {
        const e: any = entry as any;
        const n = typeof e?.name === "string" ? e.name.trim().toUpperCase() : "";
        if (!n || n !== key) continue;
        const sheetName = typeof e?.sheetName === "string" ? (e.sheetName as string) : "";
        const range = e?.range;
        if (!range) continue;
        const a1 = rangeToA1(range);
        if (!a1) continue;
        const token = sheetName ? formatSheetNameForA1(sheetName) : "";
        const prefix = token ? `${token}!` : "";
        return `${prefix}${a1}`;
      }

      return null;
    };

    const resolveStructuredRefToReference = (refText: string): string | null => {
      const trimmed = String(refText ?? "").trim();
      // Excel structured refs are never sheet-qualified in formula text.
      if (!trimmed.includes("[") || trimmed.includes("!")) return null;

      const unescapeStructuredRefItem = (value: string): string => value.replaceAll("]]", "]");

      const findColumnIndex = (columns: unknown, columnName: string): number | null => {
        if (!Array.isArray(columns)) return null;
        const target = columnName.trim().toUpperCase();
        if (!target) return null;
        for (let i = 0; i < columns.length; i += 1) {
          const col = String(columns[i] ?? "").trim();
          if (!col) continue;
          if (col.toUpperCase() === target) return i;
        }
        return null;
      };

      // Support "This Row" structured references (`Table1[[#This Row],[Amount]]`, `Table1[@Amount]`)
      // when the edited cell is within the referenced table. This avoids rewriting them to an entire
      // column range (which would be misleading) and keeps previews useful for calculated columns.
      const resolveThisRowStructuredRef = (tableName: string, columnName: string): string | null => {
        const name = String(tableName ?? "").trim();
        if (!name) return null;
        const table: any = this.searchWorkbook.getTable(name);
        if (!table) return null;

        const startRow = typeof table.startRow === "number" ? Math.trunc(table.startRow) : null;
        const startCol = typeof table.startCol === "number" ? Math.trunc(table.startCol) : null;
        const endRow = typeof table.endRow === "number" ? Math.trunc(table.endRow) : null;
        const endCol = typeof table.endCol === "number" ? Math.trunc(table.endCol) : null;
        if (startRow == null || startCol == null || endRow == null || endCol == null) return null;
        if (startRow < 0 || startCol < 0 || endRow < 0 || endCol < 0) return null;

        const baseStartRow = Math.min(startRow, endRow);
        const baseEndRow = Math.max(startRow, endRow);
        const baseStartCol = Math.min(startCol, endCol);
        const baseEndCol = Math.max(startCol, endCol);

        const colIdx = findColumnIndex(table.columns, columnName);
        if (colIdx == null) return null;
        const col = baseStartCol + colIdx;
        if (col < baseStartCol || col > baseEndCol) return null;

        // Require the formula edit target to be on the same sheet as the table.
        const tableSheet =
          typeof table.sheetName === "string" && table.sheetName.trim()
            ? table.sheetName.trim()
            : typeof table.sheet === "string" && table.sheet.trim()
              ? table.sheet.trim()
              : sheetId;
        const resolvedSheetId = tableSheet ? this.resolveSheetIdByName(tableSheet) ?? tableSheet : "";
        if (resolvedSheetId && resolvedSheetId.toLowerCase() !== sheetId.toLowerCase()) return null;
        if (resolvedSheetId && !sheetExists(resolvedSheetId)) return null;

        // "This Row" refers to a single data/totals row (not the header).
        const row = editTarget.cell.row;
        const dataStartRow = baseStartRow + 1;
        if (row < dataStartRow || row > baseEndRow) return null;

        const addr = cellToA1({ row, col });
        const sheetToken = tableSheet ? formatSheetNameForA1(tableSheet) : "";
        const prefix = sheetToken ? `${sheetToken}!` : "";
        return `${prefix}${addr}`;
      };

      // Fast-path: resolve `Table1[[#This Row],[Column]]` and `Table1[@Column]` without going through
      // `extractFormulaReferences` (which would otherwise approximate #This Row as a whole-column ref).
      const escapedItem = "((?:[^\\]]|\\]\\])+)"; // match non-] or escaped `]]`
      const qualifiedRe = new RegExp(
        `^([A-Za-z_][A-Za-z0-9_.]*)\\[\\[\\s*${escapedItem}\\s*\\]\\s*,\\s*\\[\\s*${escapedItem}\\s*\\]\\]$`,
        "i",
      );
      const qualifiedMatch = qualifiedRe.exec(trimmed);
      if (qualifiedMatch) {
        const selector = unescapeStructuredRefItem(qualifiedMatch[2]!.trim());
        const normalizedSelector = selector.trim().replace(/\s+/g, " ").toLowerCase();
        if (normalizedSelector === "#this row") {
          const columnName = unescapeStructuredRefItem(qualifiedMatch[3]!.trim());
          const resolved = resolveThisRowStructuredRef(qualifiedMatch[1]!, columnName);
          if (resolved) return resolved;
          // If we can't resolve this-row semantics, treat it as unsupported rather than falling back
          // to a whole-column approximation.
          return null;
        }
      }

      const simpleRe = new RegExp(`^([A-Za-z_][A-Za-z0-9_.]*)\\[\\s*${escapedItem}\\s*\\]$`);
      const simpleMatch = simpleRe.exec(trimmed);
      if (simpleMatch) {
        const item = unescapeStructuredRefItem(simpleMatch[2]!.trim());
        if (item.startsWith("@")) {
          const columnName = item.slice(1).trim();
          if (columnName) {
            const resolved = resolveThisRowStructuredRef(simpleMatch[1]!, columnName);
            if (resolved) return resolved;
          }
          return null;
        }
      }

      const { references } = extractFormulaReferences(trimmed, undefined, undefined, { tables: this.searchWorkbook.tables as any });
      const first = references[0];
      if (!first) return null;
      if (first.start !== 0 || first.end !== trimmed.length) return null;

      const r = first.range;
      const sheet = typeof r.sheet === "string" && r.sheet.trim() ? r.sheet.trim() : sheetId;
      const resolvedSheetId = sheet ? this.resolveSheetIdByName(sheet) ?? sheet : "";
      // Avoid creating phantom sheets during preview evaluation.
      if (resolvedSheetId && !sheetExists(resolvedSheetId)) return null;

      let a1 = "";
      try {
        a1 = rangeToA1({ startRow: r.startRow, endRow: r.endRow, startCol: r.startCol, endCol: r.endCol });
      } catch {
        return null;
      }

      const sheetToken = sheet ? formatSheetNameForA1(sheet) : "";
      const prefix = sheetToken ? `${sheetToken}!` : "";
      return `${prefix}${a1}`;
    };

    const rewriteImplicitThisRowReferences = (text: string): string | null => {
      const input = String(text ?? "");
      if (!input.includes("[@")) return null;

      // Resolve the table name from the edit target (implicit `[@...]` refs are only meaningful
      // inside a table). If we can't identify a containing table, skip rewriting.
      const tableName = (() => {
        for (const entry of this.searchWorkbook.tables.values()) {
          const table: any = entry as any;
          const name = typeof table?.name === "string" ? table.name.trim() : "";
          if (!name) continue;

          const startRow = typeof table.startRow === "number" ? Math.trunc(table.startRow) : null;
          const startCol = typeof table.startCol === "number" ? Math.trunc(table.startCol) : null;
          const endRow = typeof table.endRow === "number" ? Math.trunc(table.endRow) : null;
          const endCol = typeof table.endCol === "number" ? Math.trunc(table.endCol) : null;
          if (startRow == null || startCol == null || endRow == null || endCol == null) continue;
          if (startRow < 0 || startCol < 0 || endRow < 0 || endCol < 0) continue;

          const baseStartRow = Math.min(startRow, endRow);
          const baseEndRow = Math.max(startRow, endRow);
          const baseStartCol = Math.min(startCol, endCol);
          const baseEndCol = Math.max(startCol, endCol);

          const tableSheet =
            typeof table.sheetName === "string" && table.sheetName.trim()
              ? table.sheetName.trim()
              : typeof table.sheet === "string" && table.sheet.trim()
                ? table.sheet.trim()
                : sheetId;
          const resolvedSheetId = tableSheet ? this.resolveSheetIdByName(tableSheet) ?? tableSheet : "";
          if (resolvedSheetId && resolvedSheetId.toLowerCase() !== sheetId.toLowerCase()) continue;

          const row = editTarget.cell.row;
          const col = editTarget.cell.col;
          const dataStartRow = baseStartRow + 1;
          if (row < dataStartRow || row > baseEndRow) continue;
          if (col < baseStartCol || col > baseEndCol) continue;

          return name;
        }
        return null;
      })();

      if (!tableName) return null;

      const isWhitespaceChar = (ch: string): boolean => ch === " " || ch === "\t" || ch === "\n" || ch === "\r";
      const isIdentifierPart = (ch: string): boolean =>
        (ch >= "A" && ch <= "Z") ||
        (ch >= "a" && ch <= "z") ||
        (ch >= "0" && ch <= "9") ||
        ch === "_" ||
        ch === ".";
      const isEscapedBracket = (value: string, index: number, depth: number): boolean => {
        if (value[index] !== "]" || value[index + 1] !== "]") return false;
        // When only a single bracket group is open, `]]` cannot represent nested closes.
        if (depth === 1) return true;
        if (value[index + 2] === "]") return true;
        let k = index + 2;
        while (k < value.length && isWhitespaceChar(value[k] ?? "")) k += 1;
        const after = value[k] ?? "";
        const isDelimiterAfterClose = after === "" || after === "," || after === ";" || after === "]" || after === ")";
        return !isDelimiterAfterClose;
      };

      let out = "";
      let inString = false;
      let changed = false;
      let i = 0;

      while (i < input.length) {
        const ch = input[i] ?? "";
        if (inString) {
          out += ch;
          if (ch === '"') {
            // Escaped quote inside a string literal: "" -> "
            if (input[i + 1] === '"') {
              out += '"';
              i += 2;
              continue;
            }
            inString = false;
          }
          i += 1;
          continue;
        }

        if (ch === '"') {
          out += ch;
          inString = true;
          i += 1;
          continue;
        }

        if (ch === "[" && input[i + 1] === "@") {
          // `[@Col]` is an implicit-this-row reference. Avoid rewriting `Table[@Col]` (already qualified).
          const prev = i > 0 ? (input[i - 1] ?? "") : "";
          if (prev && isIdentifierPart(prev)) {
            out += ch;
            i += 1;
            continue;
          }
          const start = i;
          let depth = 0;
          let j = i;
          while (j < input.length) {
            const c = input[j] ?? "";
            if (c === "[") {
              depth += 1;
              j += 1;
              continue;
            }
            if (c === "]") {
              if (depth > 0 && isEscapedBracket(input, j, depth)) {
                j += 2;
                continue;
              }
              depth = Math.max(0, depth - 1);
              j += 1;
              if (depth === 0) break;
              continue;
            }
            j += 1;
          }

          // Only rewrite if we found a matching closing bracket.
          if (depth === 0 && j > start) {
            const segment = input.slice(start, j);
            // Column names containing spaces are written in the implicit-this-row shorthand
            // using a nested bracket group: `[@[Total Amount]]`.
            if (segment.startsWith("[@[") && segment.endsWith("]]") && segment.length > 5) {
              const columnText = segment.slice(3, -2);
              out += `${tableName}[[#This Row],[${columnText}]]`;
            } else {
              out += tableName + segment;
            }
            changed = true;
            i = j;
            continue;
          }
        }

        out += ch;
        i += 1;
      }

      return changed ? out : null;
    };

    let reads = 0;
    const memo = new Map<string, SpreadsheetValue>();
    const stack = new Set<string>();

    const resolveSheetId = (token: string): string | null => {
      const resolved = this.resolveSheetIdByName(token);
      if (resolved) {
        // Avoid creating phantom sheets during preview evaluation. Some callers (e.g. tab completion
        // schema providers) may surface sheets that are not yet materialized in the DocumentController;
        // treating them as missing keeps preview evaluation side-effect free.
        if (!sheetExists(resolved)) return null;
        // Preserve the current sheet id casing when possible.
        if (resolved.toLowerCase() === sheetId.toLowerCase()) return sheetId;
        return resolved;
      }

      // `DocumentController` materializes sheets lazily; if the user refers to the current sheet
      // before it has been created, allow it through.
      const unquoted = (() => {
        const t = token.trim();
        const quoted = /^'((?:[^']|'')+)'$/.exec(t);
        if (quoted) return quoted[1]!.replace(/''/g, "'").trim();
        return t;
      })();
      if (unquoted && unquoted.toLowerCase() === sheetId.toLowerCase() && sheetExists(sheetId)) return sheetId;
      return null;
    };

    const getCellValue = (ref: string): SpreadsheetValue => {
      reads += 1;
      if (reads > MAX_CELL_READS) throw new Error("preview too large");

      const normalized = String(ref ?? "").replaceAll("$", "").trim();
      let targetSheet = sheetId;
      let addr = normalized;
      const bang = normalized.lastIndexOf("!");
      if (bang >= 0) {
        const sheetToken = normalized.slice(0, bang);
        const cellToken = normalized.slice(bang + 1);
        if (sheetToken && cellToken) {
          const resolved = resolveSheetId(sheetToken);
          if (!resolved) return "#REF!";
          targetSheet = resolved;
          addr = cellToken.trim();
        }
      }

      const normalizedAddr = addr.replaceAll("$", "").trim().toUpperCase();
      const key = `${targetSheet}:${normalizedAddr}`;
      if (memo.has(key)) return memo.get(key) as SpreadsheetValue;
      if (stack.has(key)) return "#REF!";

      stack.add(key);
      if (!sheetExists(targetSheet)) {
        stack.delete(key);
        return "#REF!";
      }

      // Use `peekCell` when available to avoid materializing sheets during preview evaluation.
      const state =
        typeof (this.document as any).peekCell === "function"
          ? ((this.document as any).peekCell(targetSheet, normalizedAddr) as { value: unknown; formula: string | null })
          : (this.document.getCell(targetSheet, normalizedAddr) as { value: unknown; formula: string | null });
      let value: SpreadsheetValue;
      if (state?.formula) {
        value = evaluateFormula(state.formula, getCellValue, {
          cellAddress: `${targetSheet}!${normalizedAddr}`,
          resolveNameToReference,
          maxRangeCells: MAX_CELL_READS,
        });
      } else {
        const rawValue = state?.value ?? null;
        value =
          rawValue == null || typeof rawValue === "number" || typeof rawValue === "string" || typeof rawValue === "boolean"
            ? (rawValue as SpreadsheetValue)
            : isRichTextValue(rawValue)
              ? (rawValue.text as SpreadsheetValue)
              : null;
      }
      stack.delete(key);
      memo.set(key, value);
      return value;
    };

    try {
      const evalExpr = rewriteImplicitThisRowReferences(trimmedExpr) ?? trimmedExpr;
      const value = evaluateFormula(`=${evalExpr}`, getCellValue, {
        cellAddress: `${sheetId}!${cellAddress}`,
        resolveNameToReference,
        resolveStructuredRefToReference,
        maxRangeCells: MAX_CELL_READS,
      });
      // Errors from the lightweight evaluator usually mean unsupported syntax / functions.
      // Treat them as "preview unavailable" so we don't show misleading `#NAME?` / `#VALUE!`
      // while users are typing or when the JS evaluator lags behind the full engine.
      if (typeof value === "string" && (value === "#NAME?" || value === "#VALUE!")) return "(preview unavailable)";
      return value;
    } catch {
      return "(preview unavailable)";
    }
  }

  private resolveSheetDisplayNameById(sheetId: string): string {
    const resolved = this.sheetNameResolver?.getSheetNameById(sheetId) ?? null;
    if (resolved) return resolved;
    const metaName = (this.document as any)?.getSheetMeta?.(sheetId)?.name;
    if (typeof metaName === "string" && metaName.trim() !== "") return metaName;
    return sheetId;
  }

  private clearComputedValuesByCoord(): void {
    this.computedValuesByCoord.clear();
    this.lastComputedValuesSheetId = null;
    this.lastComputedValuesSheetCache = null;
    // Clearing computed values affects the semantics of `getCellComputedValue` for formulas
    // (it may fall back to in-process evaluation until the engine repopulates the cache).
    this.computedValuesVersion += 1;
  }

  private getComputedValuesByCoordForSheet(sheetId: string): Map<number, SpreadsheetValue> | null {
    if (this.lastComputedValuesSheetId === sheetId) {
      return this.lastComputedValuesSheetCache;
    }

    const cache = this.computedValuesByCoord.get(sheetId) ?? null;
    this.lastComputedValuesSheetId = sheetId;
    this.lastComputedValuesSheetCache = cache;
    return cache;
  }

  private invalidateComputedValues(changes: unknown): void {
    if (!Array.isArray(changes)) return;
    const coordScratch = { row: 0, col: 0 };
    let lastSheetId: string | null = null;
    let lastSheetCache: Map<number, SpreadsheetValue> | null = null;
    let invalidated = false;
    for (const change of changes) {
      const ref = change as EngineCellRef;

      let sheetId = typeof ref.sheet === "string" ? ref.sheet : undefined;
      if (!sheetId && typeof ref.sheetId === "string") sheetId = ref.sheetId;
      if (!sheetId) sheetId = this.sheetId;

      let address = typeof ref.address === "string" ? ref.address : undefined;

      // Support "Sheet1!A1" style addresses if a sheet name was embedded.
      if (address) {
        // Use the last `!` so sheet-qualified refs with `!` in the sheet id (unlikely but possible)
        // still parse correctly (matches other A1 parsing in the codebase).
        const bang = address.lastIndexOf("!");
        if (bang >= 0) {
          const maybeSheet = address.slice(0, bang);
          const cell = address.slice(bang + 1);
          if (maybeSheet && cell) {
            sheetId = this.resolveSheetIdByName(maybeSheet) ?? maybeSheet;
            address = cell;
          }
        }
      }

      let row = isInteger(ref.row) ? ref.row : undefined;
      let col = isInteger(ref.col) ? ref.col : undefined;
      if ((row === undefined || col === undefined) && address) {
        if (parseA1CellRefIntoCoord(address, coordScratch)) {
          row = coordScratch.row;
          col = coordScratch.col;
        }
      }

      if (row !== undefined && col !== undefined && col >= 0 && col < COMPUTED_COORD_COL_STRIDE) {
        const key = row * COMPUTED_COORD_COL_STRIDE + col;
        if (sheetId !== lastSheetId) {
          lastSheetId = sheetId;
          lastSheetCache = this.computedValuesByCoord.get(sheetId) ?? null;
        }
        if (lastSheetCache?.delete(key)) invalidated = true;
      }
    }
    if (invalidated) this.computedValuesVersion += 1;
  }

  private applyComputedChanges(changes: unknown): void {
    if (!Array.isArray(changes)) return;
    let updated = false;
    const sheetCount = (this.document as any)?.model?.sheets?.size;
    const shouldInvalidate = (typeof sheetCount === "number" ? sheetCount : this.document.getSheetIds().length) <= 1;
    const chartDeltas: Array<{ sheetId: string; row: number; col: number }> | null =
      this.uiReady && this.chartStore.listCharts().some((chart) => chart.sheetId === this.sheetId) ? [] : null;

    const coordScratch = { row: 0, col: 0 };
    let lastSheetId: string | null = null;
    let lastSheetCache: Map<number, SpreadsheetValue> | null = null;

    let minRow = Infinity;
    let maxRow = -Infinity;
    let minCol = Infinity;
    let maxCol = -Infinity;
    let sawActiveSheet = false;

    for (const change of changes) {
      const ref = change as EngineCellRef;

      let sheetId = typeof ref.sheet === "string" ? ref.sheet : undefined;
      if (!sheetId && typeof ref.sheetId === "string") sheetId = ref.sheetId;
      if (!sheetId) sheetId = this.sheetId;

      let address = typeof ref.address === "string" ? ref.address : undefined;

      // Support "Sheet1!A1" style addresses if a sheet name was embedded.
      if (address) {
        // Use the last `!` so sheet-qualified refs with `!` in the sheet id (unlikely but possible)
        // still parse correctly (matches other A1 parsing in the codebase).
        const bang = address.lastIndexOf("!");
        if (bang >= 0) {
          const maybeSheet = address.slice(0, bang);
          const cell = address.slice(bang + 1);
          if (maybeSheet && cell) {
            sheetId = this.resolveSheetIdByName(maybeSheet) ?? maybeSheet;
            address = cell;
          }
        }
      }

      let row = isInteger(ref.row) ? ref.row : undefined;
      let col = isInteger(ref.col) ? ref.col : undefined;
      if ((row === undefined || col === undefined) && address) {
        if (parseA1CellRefIntoCoord(address, coordScratch)) {
          row = coordScratch.row;
          col = coordScratch.col;
        }
      }

      let value = ref.value;
      // Some engine implementations omit `value` entirely to represent an empty cell.
      // Treat missing values as null so we don't keep stale computed results around.
      if (value === undefined) value = null;
      if (value !== null && typeof value !== "number" && typeof value !== "string" && typeof value !== "boolean") {
        continue;
      }

      if (row === undefined || col === undefined) continue;

      // Only cache within the supported coordinate encoding range (Excel-style columns).
      if (row < 0 || col < 0 || col >= COMPUTED_COORD_COL_STRIDE) continue;

      if (sheetId !== lastSheetId) {
        lastSheetId = sheetId;
        lastSheetCache = this.computedValuesByCoord.get(sheetId) ?? null;
      }

      let sheetCache = lastSheetCache;
      if (!sheetCache) {
        sheetCache = new Map();
        this.computedValuesByCoord.set(sheetId, sheetCache);
        lastSheetCache = sheetCache;
        if (this.lastComputedValuesSheetId === sheetId) {
          this.lastComputedValuesSheetCache = sheetCache;
        }
      }

      const key = row * COMPUTED_COORD_COL_STRIDE + col;
      const prev = sheetCache.get(key);
      if (prev === value) continue;

      sheetCache.set(key, value);
      updated = true;
      if (chartDeltas) chartDeltas.push({ sheetId, row, col });

      if (shouldInvalidate && sheetId === this.sheetId) {
        // Only invalidate within the active grid limits. The engine may produce computed
        // results for cells outside the UI's configured grid bounds; those should not
        // trigger massive provider invalidations.
        const maxDocRows = this.limits.maxRows;
        const maxDocCols = this.limits.maxCols;
        if (row < maxDocRows && col < maxDocCols) {
          sawActiveSheet = true;
          minRow = Math.min(minRow, row);
          maxRow = Math.max(maxRow, row);
          minCol = Math.min(minCol, col);
          maxCol = Math.max(maxCol, col);
        }
      }
    }

    if (updated) {
      // Computed values can change asynchronously relative to user edits (and without bumping the
      // DocumentController's sheet content version). Keep derived caches (e.g. selection summary)
      // from going stale by tracking a separate version counter.
      this.computedValuesVersion += 1;
      // Keep the status/formula bar in sync once computed values arrive.
      if (this.uiReady) this.updateStatus();
      if (this.uiReady && shouldInvalidate && sawActiveSheet) {
        if (this.sharedGrid && this.sharedProvider) {
          this.sharedProvider.invalidateDocCells({
            startRow: minRow,
            endRow: maxRow + 1,
            startCol: minCol,
            endCol: maxCol + 1,
          });
        } else if (!this.sharedGrid) {
          // Ensure the legacy renderer repaints once computed values are available (engines may
          // produce them asynchronously relative to the DocumentController change event).
          this.refresh("scroll");
        }
      }

      if (this.uiReady && chartDeltas && chartDeltas.length > 0) {
        this.markChartsDirtyFromDeltas(chartDeltas);
        this.scheduleChartContentRefresh({ deltas: chartDeltas });
      }
    }
  }

  private computeCellValue(
    sheetId: string,
    cell: CellCoord,
    memo: Map<string, Map<number, SpreadsheetValue>>,
    stack: Map<string, Set<number>>,
    options: { useEngineCache: boolean },
    flags?: { sawFormula: boolean }
  ): SpreadsheetValue {
    if (options.useEngineCache) {
      const sheetCache = this.getComputedValuesByCoordForSheet(sheetId);
      if (sheetCache && cell.col >= 0 && cell.col < COMPUTED_COORD_COL_STRIDE && cell.row >= 0) {
        const key = cell.row * COMPUTED_COORD_COL_STRIDE + cell.col;
        const cached = sheetCache.get(key);
        // `computedValuesByCoord` never stores `undefined`; a missing entry always returns undefined.
        if (cached !== undefined) {
          if (flags) {
            const state = this.document.getCell(sheetId, cell) as { formula: string | null };
            if (state?.formula != null) flags.sawFormula = true;
          }
          return cached;
        }
      }
    }

    const state = this.document.getCell(sheetId, cell) as { value: unknown; formula: string | null };
    if (flags && state?.formula != null) flags.sawFormula = true;

    // Fast path: plain values do not participate in reference cycles and do not need to
    // be memoized per evaluation call. Avoid generating A1/key strings for them.
    if (state?.formula == null) {
      if (state?.value != null) {
        return isRichTextValue(state.value) ? state.value.text : (state.value as SpreadsheetValue);
      }
      return null;
    }

    const key = cell.row * EVAL_COORD_COL_STRIDE + cell.col;
    let sheetMemo = memo.get(sheetId);
    if (!sheetMemo) {
      sheetMemo = new Map();
      memo.set(sheetId, sheetMemo);
    }

    const cached = sheetMemo.get(key);
    if (cached !== undefined || sheetMemo.has(key)) return cached ?? null;

    let sheetStack = stack.get(sheetId);
    if (!sheetStack) {
      sheetStack = new Set();
      stack.set(sheetId, sheetStack);
    }

    if (sheetStack.has(key)) return "#REF!";

    sheetStack.add(key);
    const hasAiFunction = AI_FUNCTION_CALL_RE.test(state.formula);
    const address = hasAiFunction ? cellToA1(cell) : "";
    const cellAddress = hasAiFunction ? `${sheetId}!${address}` : undefined;
    // Lazily allocate the scratch coord. Many formulas are pure expressions (e.g. `=1+1`)
    // and never invoke the reference resolver callback.
    let coordScratch: { row: number; col: number } | null = null;

    const resolveNameToReference = (name: string): string | null => {
      const entry: any = this.searchWorkbook.getName(name);
      const range = entry?.range;
      if (
        !range ||
        typeof range.startRow !== "number" ||
        typeof range.startCol !== "number" ||
        typeof range.endRow !== "number" ||
        typeof range.endCol !== "number"
      ) {
        return null;
      }

      let a1 = "";
      try {
        a1 = rangeToA1(range);
      } catch {
        return null;
      }

      const sheetName = typeof entry?.sheetName === "string" && entry.sheetName.trim() ? entry.sheetName.trim() : "";
      const token = sheetName ? formatSheetNameForA1(sheetName) : "";
      const prefix = token ? `${token}!` : "";
      return `${prefix}${a1}`;
    };

    const resolveStructuredRefToReference = (refText: string): string | null => {
      const trimmed = String(refText ?? "").trim();
      if (!trimmed.includes("[") || trimmed.includes("!")) return null;

      const { references } = extractFormulaReferences(trimmed, undefined, undefined, { tables: this.searchWorkbook.tables as any });
      const first = references[0];
      if (!first) return null;
      if (first.start !== 0 || first.end !== trimmed.length) return null;

      const r = first.range;
      const sheet = typeof r.sheet === "string" && r.sheet.trim() ? r.sheet.trim() : sheetId;

      let a1 = "";
      try {
        a1 = rangeToA1({ startRow: r.startRow, endRow: r.endRow, startCol: r.startCol, endCol: r.endCol });
      } catch {
        return null;
      }

      const sheetToken = sheet ? formatSheetNameForA1(sheet) : "";
      const prefix = sheetToken ? `${sheetToken}!` : "";
      return `${prefix}${a1}`;
    };

    const value = evaluateFormula(state.formula, (ref) => {
      const normalized = ref.trim();
      let targetSheet = sheetId;
      let targetAddress = normalized;
      const bang = normalized.lastIndexOf("!");
      if (bang >= 0) {
        const maybeSheet = normalized.slice(0, bang);
        const addr = normalized.slice(bang + 1);
        if (maybeSheet && addr) {
          const resolved = this.resolveSheetIdByName(maybeSheet);
          if (!resolved) return "#REF!";
          targetSheet = resolved;
          targetAddress = addr.trim();
        }
      }
      // Avoid allocating a fresh `{row,col}` object for every reference evaluation.
      // (The scratch coord is safe because `computeCellValue` does not retain the object.)
      const coord = coordScratch ?? (coordScratch = { row: 0, col: 0 });
      coord.row = 0;
      coord.col = 0;
      parseA1CellRefIntoCoord(targetAddress, coord);
      return this.computeCellValue(targetSheet, coord, memo, stack, options);
    }, { ai: this.aiCellFunctions, cellAddress, resolveNameToReference, resolveStructuredRefToReference });

    sheetStack.delete(key);
    sheetMemo.set(key, value);
    return value;
  }

  private applyEdit(sheetId: string, cell: CellCoord, rawValue: string, options?: { label?: string }): void {
    if (this.isReadOnly()) {
      showCollabEditRejectedToast([
        { sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    // In collab/permissioned contexts, `DocumentController.canEditCell` may filter out
    // user edits entirely. Without a UX signal this can look like the UI "snapped back".
    const canEditCell = (this.document as any)?.canEditCell;
    if (typeof canEditCell === "function") {
      try {
        const allowed = Boolean(canEditCell.call(this.document, { sheetId, row: cell.row, col: cell.col }));
        if (!allowed) {
          showCollabEditRejectedToast([
            {
              sheetId,
              row: cell.row,
              col: cell.col,
              rejectionKind: "cell",
              rejectionReason: this.inferCollabEditRejectionReason({ sheetId, row: cell.row, col: cell.col }),
            },
          ]);
          return;
        }
      } catch {
        // Best-effort: fall through to attempting the edit.
      }
    }

    const label = options?.label ?? "Edit cell";
    const original = this.document.getCell(sheetId, cell) as { value: unknown; formula: string | null };
    if (rawValue.trim() === "") {
      this.document.clearCell(sheetId, cell, { label: options?.label ?? "Clear cell" });
      return;
    }

    // Preserve rich-text formatting runs when editing a rich-text cell with plain text
    // (but still allow formulas / leading apostrophes to override rich-text semantics).
    const trimmedStart = rawValue.trimStart();
    if (!trimmedStart.startsWith("=") && !rawValue.startsWith("'") && isRichTextValue(original?.value)) {
      const updated = applyPlainTextEdit(original.value, rawValue);
      if (original.formula == null && updated === original.value) {
        // No-op edit: keep rich runs without creating a history entry.
        return;
      }
      this.document.setCellValue(sheetId, cell, updated, { label });
      return;
    }
    this.document.setCellInput(sheetId, cell, rawValue, { label });
  }

  private inferCollabEditRejectionReason(cell: { sheetId: string; row: number; col: number }): "permission" | "encryption" | "unknown" {
    const session = this.collabSession;
    if (!session) return "permission";

    // Best-effort: infer missing-key encryption failures so the toast can be specific.
    try {
      const encryption = typeof (session as any).getEncryptionConfig === "function" ? (session as any).getEncryptionConfig() : null;
      if (!encryption || typeof encryption.keyForCell !== "function") return "permission";

      const key = encryption.keyForCell(cell) ?? null;
      const shouldEncrypt =
        typeof encryption.shouldEncryptCell === "function" ? Boolean(encryption.shouldEncryptCell(cell)) : key != null;

      // If the cell is already encrypted in the shared doc, or encryption is required by
      // config, and we don't have a key, treat this as an encryption rejection.
      const cellKey = makeCellKey(cell);
      const cellData = (session as any).cells?.get?.(cellKey) ?? null;
      const hasEnc = cellData && typeof cellData.get === "function" ? cellData.get("enc") !== undefined : false;
      if ((hasEnc || shouldEncrypt) && !key) return "encryption";
    } catch {
      // ignore
    }

    return "permission";
  }

  private commitFormulaBar(text: string, commit: FormulaBarCommit): void {
    this.endKeyboardRangeSelection();
    if (this.isReadOnly()) {
      this.cancelFormulaBar();
      return;
    }
    const target = this.formulaEditCell ?? { sheetId: this.sheetId, cell: { ...this.selection.active } };
    this.applyEdit(target.sheetId, target.cell, text);

    this.formulaEditCell = null;
    this.updateEditState();
    this.referencePreview = null;
    this.referenceHighlights = [];
    this.referenceHighlightsSource = [];
    this.updateEditState();

    if (this.sharedGrid) {
      this.syncSharedGridInteractionMode();
      this.sharedGrid.clearRangeSelection();
      this.sharedGrid.renderer.setReferenceHighlights(null);
    }

    // Restore selection to the original edit cell (sheet + cell), even if the user navigated
    // to another sheet while picking ranges. Navigation after commit (Enter/Tab) should be
    // relative to the original edit cell, not whatever cell/range was active during range-picking.
    //
    // When possible, preserve the existing selection ranges so Tab/Enter navigation cycles within
    // a multi-cell selection (Excel-like).
    if (target.sheetId !== this.sheetId) {
      this.activateCell(
        { sheetId: target.sheetId, row: target.cell.row, col: target.cell.col },
        { scrollIntoView: false, focus: false }
      );
    } else {
      const primaryRange = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0] ?? null;
      if (primaryRange && cellInRange(target.cell, primaryRange)) {
        this.selection = buildSelection(
          {
            ranges: this.selection.ranges,
            active: target.cell,
            anchor: this.selection.anchor,
            activeRangeIndex: this.selection.activeRangeIndex,
          },
          this.limits
        );
      } else {
        this.selection = setActiveCell(this.selection, target.cell, this.limits);
      }
    }

    if (commit.reason === "enter" || commit.reason === "tab") {
      const next = navigateSelectionByKey(
        this.selection,
        commit.reason === "enter" ? "Enter" : "Tab",
        { shift: commit.shift, primary: false },
        this.usedRangeProvider(),
        this.limits
      );
      if (next) this.selection = next;
    }

    this.ensureActiveCellVisible();
    this.scrollCellIntoView(this.selection.active);
    if (this.sharedGrid) this.syncSharedGridSelectionFromState({ scrollIntoView: false });
    this.refresh();
    this.focus();
  }

  private cancelFormulaBar(): void {
    this.endKeyboardRangeSelection();
    const target = this.formulaEditCell;
    this.formulaEditCell = null;
    this.updateEditState();
    this.referencePreview = null;
    this.referenceHighlights = [];
    this.referenceHighlightsSource = [];

    if (this.sharedGrid) {
      this.syncSharedGridInteractionMode();
      this.sharedGrid.clearRangeSelection();
      this.sharedGrid.renderer.setReferenceHighlights(null);
    }

    this.updateEditState();

    if (target) {
      // Restore the original edit location (sheet + cell).
      if (target.sheetId !== this.sheetId) {
        this.activateCell({ sheetId: target.sheetId, row: target.cell.row, col: target.cell.col });
      } else {
        const primaryRange = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0] ?? null;
        if (primaryRange && cellInRange(target.cell, primaryRange)) {
          this.selection = buildSelection(
            {
              ranges: this.selection.ranges,
              active: target.cell,
              anchor: this.selection.anchor,
              activeRangeIndex: this.selection.activeRangeIndex,
            },
            this.limits
          );
          this.ensureActiveCellVisible();
          const didScroll = this.scrollCellIntoView(this.selection.active);
          if (this.sharedGrid) this.syncSharedGridSelectionFromState({ scrollIntoView: false });
          else if (didScroll) this.ensureViewportMappingCurrent();
          this.renderSelection();
          this.updateStatus();
          this.focus();
        } else {
          this.activateCell({ sheetId: target.sheetId, row: target.cell.row, col: target.cell.col });
        }
      }
      this.renderReferencePreview();
      return;
    }

    this.ensureActiveCellVisible();
    const didScroll = this.scrollCellIntoView(this.selection.active);
    if (this.sharedGrid) this.syncSharedGridSelectionFromState({ scrollIntoView: false });
    else if (didScroll) this.ensureViewportMappingCurrent();
    this.renderReferencePreview();
    this.renderSelection();
    this.updateStatus();
    this.focus();
  }

  private computeReferenceHighlightsForSheet(
    sheetId: string,
    highlights: typeof this.referenceHighlightsSource
  ): Array<{ start: CellCoord; end: CellCoord; color: string; active: boolean }> {
    if (!highlights || highlights.length === 0) return [];

    const formulaSheetId = this.formulaEditCell?.sheetId ?? sheetId;

    return highlights
      .filter((h) => {
        const sheet = h.range.sheet;
        if (!sheet) {
          // Unqualified references (no sheet qualifier) are relative to the sheet containing the formula.
          // When the user is viewing another sheet while still editing the formula, don't render
          // misleading highlights on the active sheet.
          return formulaSheetId.toLowerCase() === sheetId.toLowerCase();
        }
        const resolved = this.getSheetIdByName(sheet);
        if (!resolved) return false;
        return resolved.toLowerCase() === sheetId.toLowerCase();
      })
      .map((h) => ({
        start: { row: h.range.startRow, col: h.range.startCol },
        end: { row: h.range.endRow, col: h.range.endCol },
        color: h.color,
        active: Boolean(h.active),
      }));
  }

  private renderReferencePreview(): void {
    if (this.sharedGrid) {
      // Reference previews are rendered via the shared grid's `rangeSelection` overlay.
      // The legacy implementation paints onto the content canvas, which would clobber cell text.
      return;
    }

    const ctx = this.referenceCtx;
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, this.referenceCanvas.width, this.referenceCanvas.height);
    ctx.restore();

    if (this.referenceHighlights.length === 0 && !this.referencePreview) return;
    this.ensureViewportMappingCurrent();

    const drawRangeOutline = (
      startRow: number,
      endRow: number,
      startCol: number,
      endCol: number,
      options: { color: string; lineWidth?: number; dash?: number[] }
    ) => {
      // Clip preview rendering to the visible viewport so dragging a range that
      // extends offscreen doesn't crash (and still provides visual feedback).
      const visibleStartRow = this.visibleRows.find((row) => row >= startRow && row <= endRow) ?? null;
      const visibleEndRow = (() => {
        for (let i = this.visibleRows.length - 1; i >= 0; i -= 1) {
          const row = this.visibleRows[i]!;
          if (row >= startRow && row <= endRow) return row;
        }
        return null;
      })();
      const visibleStartCol = this.visibleCols.find((col) => col >= startCol && col <= endCol) ?? null;
      const visibleEndCol = (() => {
        for (let i = this.visibleCols.length - 1; i >= 0; i -= 1) {
          const col = this.visibleCols[i]!;
          if (col >= startCol && col <= endCol) return col;
        }
        return null;
      })();
      if (visibleStartRow == null || visibleEndRow == null || visibleStartCol == null || visibleEndCol == null) return;

      const startRect = this.getCellRect({ row: visibleStartRow, col: visibleStartCol });
      const endRect = this.getCellRect({ row: visibleEndRow, col: visibleEndCol });
      if (!startRect || !endRect) return;

      const x = startRect.x;
      const y = startRect.y;
      const width = endRect.x + endRect.width - startRect.x;
      const height = endRect.y + endRect.height - startRect.y;
      if (width <= 0 || height <= 0) return;

      ctx.save();
      ctx.beginPath();
      ctx.rect(this.rowHeaderWidth, this.colHeaderHeight, this.viewportWidth(), this.viewportHeight());
      ctx.clip();
      ctx.strokeStyle = options.color;
      ctx.lineWidth = options.lineWidth ?? 2;
      ctx.setLineDash(options.dash ?? []);
      const inset = ctx.lineWidth / 2;
      if (width > ctx.lineWidth && height > ctx.lineWidth) {
        ctx.strokeRect(x + inset, y + inset, width - ctx.lineWidth, height - ctx.lineWidth);
      }
      ctx.restore();
    };

    // In formula-editing mode, the formula bar provides `referenceHighlights` with
    // per-reference colors and an `active` flag. Render all non-active refs as
    // dashed outlines, and the active one (if any) as a thicker solid outline.
    //
    // When not editing, only the single `referencePreview` (hover) is shown.
    for (const highlight of this.referenceHighlights) {
      if (highlight.active) continue;
      const startRow = Math.min(highlight.start.row, highlight.end.row);
      const endRow = Math.max(highlight.start.row, highlight.end.row);
      const startCol = Math.min(highlight.start.col, highlight.end.col);
      const endCol = Math.max(highlight.start.col, highlight.end.col);
      drawRangeOutline(startRow, endRow, startCol, endCol, { color: highlight.color, dash: [4, 3] });
    }

    for (const highlight of this.referenceHighlights) {
      if (!highlight.active) continue;
      const startRow = Math.min(highlight.start.row, highlight.end.row);
      const endRow = Math.max(highlight.start.row, highlight.end.row);
      const startCol = Math.min(highlight.start.col, highlight.end.col);
      const endCol = Math.max(highlight.start.col, highlight.end.col);
      drawRangeOutline(startRow, endRow, startCol, endCol, { color: highlight.color, lineWidth: 3 });
    }

    if (this.referenceHighlights.length === 0 && this.referencePreview) {
      const startRow = Math.min(this.referencePreview.start.row, this.referencePreview.end.row);
      const endRow = Math.max(this.referencePreview.start.row, this.referencePreview.end.row);
      const startCol = Math.min(this.referencePreview.start.col, this.referencePreview.end.col);
      const endCol = Math.max(this.referencePreview.start.col, this.referencePreview.end.col);
      drawRangeOutline(startRow, endRow, startCol, endCol, {
        color: resolveCssVar("--warning", { fallback: "CanvasText" }),
        dash: [4, 3],
      });
    }
  }

  private usedRangeProvider() {
    return {
      getUsedRange: () => this.computeUsedRange(),
      isCellEmpty: (cell: CellCoord) => {
        const state = this.document.getCell(this.sheetId, cell) as { value: unknown; formula: string | null };
        return state?.value == null && state?.formula == null;
      },
      // Shared-grid mode supports manual row/col hide/unhide but does not yet render outline-collapsed
      // (group) hidden state. Match Excel-like navigation by skipping user-hidden indices while
      // ignoring outline/filter-hidden ones.
      isRowHidden: (row: number) => this.isRowHidden(row),
      isColHidden: (col: number) => this.isColHidden(col),
    };
  }

  private computeUsedRange(): Range | null {
    return this.document.getUsedRange(this.sheetId);
  }

  /**
   * Clear the contents (values/formulas) of all selection ranges, preserving formats.
   *
   * This is intentionally separate from Delete-key handling so other UI surfaces
   * (command palette, context menus, extensions) can invoke it via `CommandRegistry`.
   */
  clearSelectionContents(): void {
    if (this.isReadOnly()) {
      const cell = this.selection.active;
      showCollabEditRejectedToast([
        { sheetId: this.sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
      ]);
      return;
    }
    if (this.isEditing()) return;
    this.clearSelectionContentsInternal();
    this.refresh();
  }

  private clearSelectionContentsInternal(): void {
    const used = this.computeUsedRange();
    if (!used) return;
    const label = t("command.edit.clearContents");
    for (const range of this.selection.ranges) {
      const clipped = intersectRanges(range, used);
      if (!clipped) continue;
      this.document.clearRange(
        this.sheetId,
        {
          start: { row: clipped.startRow, col: clipped.startCol },
          end: { row: clipped.endRow, col: clipped.endCol }
        },
        { label }
      );
    }
  }

  private getInlineEditSelectionRange(): Range | null {
    const range = this.selection.ranges[this.selection.activeRangeIndex] ?? this.selection.ranges[0];
    return range ? { ...range } : null;
  }
}

function isRichTextValue(
  value: unknown
): value is { text: string; runs?: Array<{ start: number; end: number; style?: Record<string, unknown> }> } {
  if (typeof value !== "object" || value == null) return false;
  const v = value as { text?: unknown; runs?: unknown };
  if (typeof v.text !== "string") return false;
  if (v.runs == null) return true;
  return Array.isArray(v.runs);
}

function intersectRanges(a: Range, b: Range): Range | null {
  const startRow = Math.max(a.startRow, b.startRow);
  const endRow = Math.min(a.endRow, b.endRow);
  const startCol = Math.max(a.startCol, b.startCol);
  const endCol = Math.min(a.endCol, b.endCol);
  if (startRow > endRow || startCol > endCol) return null;
  return { startRow, endRow, startCol, endCol };
}

function mod(n: number, m: number): number {
  if (!Number.isFinite(n) || !Number.isFinite(m) || m === 0) return 0;
  return ((n % m) + m) % m;
}

function coerceNumber(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed === "") return null;
    const num = Number(trimmed);
    return Number.isFinite(num) ? num : null;
  }
  return null;
}

function colToName(col: number): string {
  if (!Number.isFinite(col) || col < 0) return "";
  return colToNameA1(col);
}

function parseA1(a1: string): CellCoord {
  try {
    const { row0, col0 } = fromA1A1(a1);
    return { row: row0, col: col0 };
  } catch {
    return { row: 0, col: 0 };
  }
}

function parseA1CellRefIntoCoord(address: string, out: { row: number; col: number }): boolean {
  // Avoid regex + throwy parsing on hot paths. This parses a single A1 cell reference
  // (optionally containing `$` absolute markers) and writes the 0-based coord into `out`.
  //
  // Supported forms:
  // - A1
  // - $A$1
  // - A$1
  // - $A1
  //
  // Not supported:
  // - ranges (A1:B2)
  // - sheet-qualified refs (Sheet1!A1) (callers should split first)
  // - R1C1 / structured refs
  let start = 0;
  let end = address.length;

  // Trim ASCII whitespace without allocating.
  while (start < end && address.charCodeAt(start) <= 32) start += 1;
  while (end > start && address.charCodeAt(end - 1) <= 32) end -= 1;
  if (start >= end) return false;

  let i = start;
  // Optional leading `$`.
  if (address.charCodeAt(i) === 36) i += 1;

  // Column letters.
  let col1 = 0;
  let sawLetter = false;
  while (i < end) {
    const code = address.charCodeAt(i);
    let n = 0;
    if (code >= 65 && code <= 90) n = code - 64; // A-Z
    else if (code >= 97 && code <= 122) n = code - 96; // a-z
    else break;
    sawLetter = true;
    col1 = col1 * 26 + n;
    i += 1;
  }
  if (!sawLetter) return false;

  // Optional `$` before row digits.
  if (i < end && address.charCodeAt(i) === 36) i += 1;
  if (i >= end) return false;

  // Row digits (1-based, must start with 1-9).
  const firstDigit = address.charCodeAt(i);
  if (firstDigit < 49 || firstDigit > 57) return false;

  let row1 = 0;
  while (i < end) {
    const code = address.charCodeAt(i);
    if (code < 48 || code > 57) return false;
    row1 = row1 * 10 + (code - 48);
    i += 1;
  }

  // fromA1 returns 0-based row/col.
  out.row = row1 - 1;
  out.col = col1 - 1;
  return true;
}
