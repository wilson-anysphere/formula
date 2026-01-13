import type {
  CanvasGridImageResolver,
  CellRange,
  CellRange as GridCellRange,
  CellRichText,
  FillCommitEvent,
  GridAxisSizeChange
} from "@formula/grid";
import { MAX_GRID_ZOOM, MIN_GRID_ZOOM } from "@formula/grid";
import type { CellRange as FillEngineRange } from "@formula/fill-engine";
import type { DocumentController } from "../../document/documentController.js";
import type { DrawingObject, ImageStore } from "../../drawings/types";
import { DrawingOverlay, type ChartRenderer, type GridGeometry, type Viewport as DrawingsViewport } from "../../drawings/overlay";
import { showToast } from "../../extensions/ui.js";
import { showCollabEditRejectedToast } from "../../collab/editRejectionToast";
import { applyFillCommitToDocumentController } from "../../fill/applyFillCommit";
import { CellEditorOverlay, type EditorCommit } from "../../editor/cellEditorOverlay.js";
import { DesktopSharedGrid, type DesktopSharedGridCallbacks } from "../shared/desktopSharedGrid.js";
import { DocumentCellProvider } from "../shared/documentCellProvider.js";
import { applyPlainTextEdit } from "../text/rich-text/edit.js";

type ScrollState = { scrollX: number; scrollY: number };

const MAX_FILL_CELLS = 200_000;

function clamp(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return Math.min(max, Math.max(min, value));
}

type RichTextValue = CellRichText;

function isRichTextValue(value: unknown): value is RichTextValue {
  if (typeof value !== "object" || value == null) return false;
  const v = value as { text?: unknown; runs?: unknown };
  if (typeof v.text !== "string") return false;
  if (v.runs == null) return true;
  return Array.isArray(v.runs);
}

function isPlainObject(value: unknown): value is Record<string, any> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function parseImageCellPayload(value: unknown): { imageId: string; altText?: string } | null {
  if (!isPlainObject(value)) return null;
  const obj: any = value;

  let payload: any = null;
  if (typeof obj.type === "string") {
    if (obj.type.toLowerCase() !== "image") return null;
    payload = isPlainObject(obj.value) ? obj.value : null;
  } else {
    payload = obj;
  }
  if (!payload) return null;

  const imageId = payload.imageId ?? payload.image_id ?? payload.id;
  if (typeof imageId !== "string" || imageId.trim() === "") return null;

  const altTextRaw = payload.altText ?? payload.alt_text ?? payload.alt;
  const altText = typeof altTextRaw === "string" && altTextRaw.trim() !== "" ? altTextRaw : undefined;

  return { imageId, altText };
}

function focusWithoutScroll(el: HTMLElement): void {
  try {
    (el as any).focus?.({ preventScroll: true });
  } catch {
    el.focus?.();
  }
}

/**
 * Mounts a shared CanvasGrid-based renderer into an arbitrary container.
 *
 * This is used by the desktop split-view secondary pane.
 */
export class SecondaryGridView {
  readonly container: HTMLElement;
  readonly provider: DocumentCellProvider;
  readonly grid: DesktopSharedGrid;

  private disposed = false;
  private uiReady = false;
  private readonly ownsProvider: boolean;
  private readonly document: DocumentController;
  private readonly getSheetId: () => string;
  private readonly getComputedValue: (cell: { row: number; col: number }) => string | number | boolean | null;
  private readonly getDrawingObjects: (sheetId: string) => DrawingObject[];
  private readonly drawingsImages: ImageStore;
  private readonly getSelectedDrawingId?: () => number | null;
  private readonly onRequestRefresh?: () => void;
  private readonly onEditStateChange?: (isEditing: boolean) => void;
  private readonly headerRows = 1;
  private readonly headerCols = 1;
  private readonly editor: CellEditorOverlay;
  private editingCell: { row: number; col: number } | null = null;
  private suppressFocusRestoreOnNextCommandCommit = false;

  private sheetId: string;

  private readonly resizeObserver: ResizeObserver;
  private readonly disposeFns: Array<() => void> = [];

  private readonly persistDebounceMs: number;
  private pendingScroll: ScrollState | null = null;
  private scrollPersistTimer: number | null = null;
  private pendingZoom: number | null = null;
  private zoomPersistTimer: number | null = null;
  private lastZoom = 1;
  private suppressSelectionCallbacks = false;
  // Track axis sizing versions so we can invalidate cached drawings bounds when the
  // renderer's row/col sizes change during interactive resize drags.
  private rowsVersion = -1;
  private colsVersion = -1;

  private readonly persistScroll?: (scroll: ScrollState) => void;
  private readonly persistZoom?: (zoom: number) => void;

  private readonly drawingsCanvas: HTMLCanvasElement;
  private readonly drawingsOverlay: DrawingOverlay;
  private sheetViewFrozen: { rows: number; cols: number } = { rows: 0, cols: 0 };
  private drawingsRenderInProgress = false;
  private drawingsRenderQueued = false;
  // Rendering is synchronous, but some callers may trigger nested render requests (e.g. as
  // a side-effect of scroll/resize events). Keep a tiny re-entrancy/queue guard so we can
  // coalesce those into a single pass.

  constructor(options: {
    container: HTMLElement;
    /**
     * Optional provider to reuse (e.g. the primary shared-grid provider).
     *
     * When supplied, `document`/`getSheetId`/`showFormulas`/`getComputedValue` are
     * ignored for provider creation, but still accepted for convenience.
     */
    provider?: DocumentCellProvider;
    /**
     * Optional image resolver to reuse (e.g. the primary shared-grid image resolver).
     *
     * When supplied, in-cell image values (`CellData.image`) can render in the secondary pane
     * without requiring the pane to manage its own image store.
     */
    imageResolver?: CanvasGridImageResolver | null;
    document: DocumentController;
    getSheetId: () => string;
    rowCount: number;
    colCount: number;
    showFormulas: () => boolean;
    getComputedValue: (cell: { row: number; col: number }) => string | number | boolean | null;
    onSelectionChange?: (selection: { row: number; col: number } | null) => void;
    onSelectionRangeChange?: (range: CellRange | null) => void;
    getCommentMeta?: (row: number, col: number) => { resolved: boolean } | null;
    callbacks?: DesktopSharedGridCallbacks;
    initialScroll?: ScrollState;
    initialZoom?: number;
    persistScroll?: (scroll: ScrollState) => void;
    persistZoom?: (zoom: number) => void;
    persistDebounceMs?: number;
    /**
     * Drawing/picture objects for the currently visible sheet.
     */
    getDrawingObjects: (sheetId: string) => DrawingObject[];
    /**
     * Image store backing drawing objects.
     */
    images: ImageStore;
    /**
     * Optional chart renderer for drawing objects of kind `chart`.
     *
     * When omitted, charts render as placeholders in the secondary pane.
     */
    chartRenderer?: ChartRenderer;
    /**
     * Optional hook to render selection handles on a selected drawing object.
     */
    getSelectedDrawingId?: () => number | null;
    /**
     * Optional hook to refresh non-grid UI when the secondary pane mutates the document
     * (e.g. formula bar, charts, auditing overlays in the primary pane).
     */
    onRequestRefresh?: () => void;
    /**
     * Optional hook to notify when the secondary in-cell editor opens/closes.
     *
     * SpreadsheetApp's edit state does not include this secondary editor, so the desktop
     * shell uses this to keep UI state (status bar, shortcut gating, etc) in sync.
     */
    onEditStateChange?: (isEditing: boolean) => void;
  }) {
    this.container = options.container;
    this.document = options.document;
    this.getSheetId = options.getSheetId;
    this.getComputedValue = options.getComputedValue;
    this.getDrawingObjects = options.getDrawingObjects;
    this.drawingsImages = options.images;
    this.getSelectedDrawingId = options.getSelectedDrawingId;
    this.persistScroll = options.persistScroll;
    this.persistZoom = options.persistZoom;
    this.persistDebounceMs = options.persistDebounceMs ?? 150;
    this.sheetId = options.getSheetId();
    this.onRequestRefresh = options.onRequestRefresh;
    this.onEditStateChange = options.onEditStateChange;

    // Clear any placeholder content from the split-view scaffolding.
    this.container.replaceChildren();

    const gridCanvas = document.createElement("canvas");
    gridCanvas.className = "grid-canvas grid-canvas--base";
    gridCanvas.setAttribute("aria-hidden", "true");

    const contentCanvas = document.createElement("canvas");
    contentCanvas.className = "grid-canvas grid-canvas--content";
    contentCanvas.setAttribute("aria-hidden", "true");

    const drawingsCanvas = document.createElement("canvas");
    drawingsCanvas.className = "grid-canvas grid-canvas--drawings";
    drawingsCanvas.setAttribute("aria-hidden", "true");
    drawingsCanvas.classList.add("grid-canvas--shared-drawings");

    const selectionCanvas = document.createElement("canvas");
    selectionCanvas.className = "grid-canvas grid-canvas--selection";
    selectionCanvas.setAttribute("aria-hidden", "true");
    // CanvasGridRenderer assigns z-index inline (0/1/2). Ensure the selection layer still
    // sits above the drawings overlay canvas in shared-grid mode by applying the shared
    // stacking modifier (charts-overlay.css).
    selectionCanvas.classList.add("grid-canvas--shared-selection");

    this.container.appendChild(gridCanvas);
    this.container.appendChild(contentCanvas);
    this.container.appendChild(drawingsCanvas);
    this.container.appendChild(selectionCanvas);

    const vTrack = document.createElement("div");
    vTrack.setAttribute("aria-hidden", "true");
    vTrack.setAttribute("data-testid", "scrollbar-track-y-secondary");
    vTrack.className = "grid-scrollbar-track grid-scrollbar-track--vertical";

    const vThumb = document.createElement("div");
    vThumb.setAttribute("aria-hidden", "true");
    vThumb.setAttribute("data-testid", "scrollbar-thumb-y-secondary");
    vThumb.className = "grid-scrollbar-thumb";
    vTrack.appendChild(vThumb);
    this.container.appendChild(vTrack);

    const hTrack = document.createElement("div");
    hTrack.setAttribute("aria-hidden", "true");
    hTrack.setAttribute("data-testid", "scrollbar-track-x-secondary");
    hTrack.className = "grid-scrollbar-track grid-scrollbar-track--horizontal";

    const hThumb = document.createElement("div");
    hThumb.setAttribute("aria-hidden", "true");
    hThumb.setAttribute("data-testid", "scrollbar-thumb-x-secondary");
    hThumb.className = "grid-scrollbar-thumb";
    hTrack.appendChild(hThumb);
    this.container.appendChild(hTrack);

    // Editor overlay for in-place cell editing in the secondary pane.
    this.editor = new CellEditorOverlay(this.container, {
      onCommit: (commit) => {
        const suppressFocusRestore =
          commit.reason === "command" && this.suppressFocusRestoreOnNextCommandCommit;
        this.suppressFocusRestoreOnNextCommandCommit = false;
        this.editingCell = null;
        this.onEditStateChange?.(false);
        this.applyEdit(commit.cell, commit.value);
        if (commit.reason === "enter" || commit.reason === "tab") {
          this.advanceSelectionAfterEdit(commit);
        }
        this.onRequestRefresh?.();
        if (!suppressFocusRestore) focusWithoutScroll(this.container);
      },
      onCancel: () => {
        this.suppressFocusRestoreOnNextCommandCommit = false;
        this.editingCell = null;
        this.onEditStateChange?.(false);
        focusWithoutScroll(this.container);
      }
    });

    // Excel behavior: leaving in-cell editing (e.g. clicking the primary pane, ribbon, etc)
    // should commit the draft text.
    //
    // IMPORTANT: Avoid stealing focus back from whatever surface the user clicked. When the
    // editor blurs to an element other than the secondary grid root itself, suppress the
    // focus-restore logic that normally runs after a command commit.
    const onEditorBlur = (event: FocusEvent) => {
      if (!this.editor.isOpen()) return;
      const next = event.relatedTarget as Node | null;
      // Only restore focus to the grid when the blur target is the grid root itself. If focus
      // moved to any other element (including focusable overlays inside the container), avoid
      // stealing it back.
      this.suppressFocusRestoreOnNextCommandCommit = next !== this.container;
      this.editor.commit("command");
    };
    this.editor.element.addEventListener("blur", onEditorBlur);
    this.disposeFns.push(() => this.editor.element.removeEventListener("blur", onEditorBlur));

    this.ownsProvider = options.provider == null;
    this.provider =
      options.provider ??
      new DocumentCellProvider({
        document: options.document,
        getSheetId: options.getSheetId,
        headerRows: this.headerRows,
        headerCols: this.headerCols,
        rowCount: options.rowCount,
        colCount: options.colCount,
        showFormulas: options.showFormulas,
        getComputedValue: options.getComputedValue,
        getCommentMeta: options.getCommentMeta,
        cssVarRoot: this.container,
      });

    const externalCallbacks: DesktopSharedGridCallbacks = options.callbacks ?? {};

    this.grid = new DesktopSharedGrid({
      container: this.container,
      provider: this.provider,
      rowCount: options.rowCount,
      colCount: options.colCount,
      frozenRows: this.headerRows,
      frozenCols: this.headerCols,
      defaultRowHeight: 24,
      defaultColWidth: 100,
      imageResolver: options.imageResolver ?? null,
      enableResize: true,
      enableKeyboard: true,
      canvases: { grid: gridCanvas, content: contentCanvas, selection: selectionCanvas },
      scrollbars: { vTrack, vThumb, hTrack, hThumb },
      callbacks: {
        ...externalCallbacks,
        onScroll: (scroll, viewport) => {
          this.container.dataset.scrollX = String(scroll.x);
          this.container.dataset.scrollY = String(scroll.y);
          this.schedulePersistScroll({ scrollX: scroll.x, scrollY: scroll.y });

          const zoom = this.grid.renderer.getZoom();
          this.container.dataset.zoom = String(zoom);
          if (Math.abs(zoom - this.lastZoom) > 1e-6) {
            this.lastZoom = zoom;
            this.schedulePersistZoom(zoom);
          }

          if (!this.uiReady) {
            externalCallbacks.onScroll?.(scroll, viewport);
            return;
          }

          const rowsVersion = this.grid.renderer.scroll.rows.getVersion();
          const colsVersion = this.grid.renderer.scroll.cols.getVersion();
          if (rowsVersion !== this.rowsVersion || colsVersion !== this.colsVersion) {
            this.rowsVersion = rowsVersion;
            this.colsVersion = colsVersion;
            this.drawingsOverlay.invalidateSpatialIndex();
          }

          this.repositionEditor();
          void this.renderDrawings();
          externalCallbacks.onScroll?.(scroll, viewport);
        },
        onAxisSizeChange: (change) => {
          this.onAxisSizeChange(change);
          externalCallbacks.onAxisSizeChange?.(change);
        },
        onSelectionChange: (selection) => {
          if (!this.suppressSelectionCallbacks) {
            options.onSelectionChange?.(selection);
          }
          externalCallbacks.onSelectionChange?.(selection);
        },
        onSelectionRangeChange: (range) => {
          if (!this.suppressSelectionCallbacks) {
            options.onSelectionRangeChange?.(range);
          }
          externalCallbacks.onSelectionRangeChange?.(range);
        },
        onRequestCellEdit: (request) => {
          this.openEditor(request);
          externalCallbacks.onRequestCellEdit?.(request);
        },
        onFillCommit: (event) => {
          this.onFillCommit(event);
          void externalCallbacks.onFillCommit?.(event);
        },
      }
    });

    // CanvasGridRenderer caches decoded ImageBitmaps per imageId. When the underlying
    // workbook images change (e.g. applyState restore or future backend hydration),
    // invalidate cached bitmaps so the grid re-resolves.
    //
    // Additionally, re-render the drawings overlay when drawings/images change so pictures
    // inserted into the sheet show up in the secondary pane without requiring a scroll.
    const unsubscribeImages = this.document.on("change", (payload: any) => {
      if (this.disposed) return;
      const source = typeof payload?.source === "string" ? payload.source : "";
      // Axis resize updates originating from this pane already update the CanvasGridRenderer
      // interactively and explicitly trigger a drawings re-render via `onAxisSizeChange`.
      // Avoid double-rendering here (the first render can run before the spatial index is
      // invalidated, producing stale geometry).
      if (source === "secondaryGridAxis") {
        return;
      }
      if (source === "applyState") {
        this.grid.renderer.clearImageCache?.();
        this.drawingsOverlay.clearImageCache();
      } else {
        const deltas = Array.isArray(payload?.imageDeltas)
          ? payload.imageDeltas
          : Array.isArray(payload?.imagesDeltas)
            ? payload.imagesDeltas
            : [];
        for (const delta of deltas) {
          const imageId =
            typeof delta?.imageId === "string" ? delta.imageId : typeof delta?.id === "string" ? delta.id : null;
          if (!imageId) continue;
          this.grid.renderer.invalidateImage?.(imageId);
          this.drawingsOverlay.invalidateImage(imageId);
        }
      }

      // `sheetViewDeltas` can affect both drawings metadata (stored in sheet view state) and the
      // grid geometry that drawings are anchored to (frozen panes, row/col sizes).
      //
      // `syncSheetViewFromDocument()` already handles these deltas by applying view state into the
      // renderer, invalidating the drawing spatial index, and scheduling a re-render. Avoid
      // triggering a drawings render *before* that sync happens (this handler is registered before
      // the dedicated sheetView listener), otherwise the overlay can briefly render with stale
      // geometry (and unit tests can observe the stale pass).
      const sheetViewDeltas = Array.isArray(payload?.sheetViewDeltas) ? payload.sheetViewDeltas : [];
      if (sheetViewDeltas.length > 0) {
        const sheetId = this.getSheetId();
        if (sheetViewDeltas.some((delta: any) => String(delta?.sheetId ?? "") === sheetId)) {
          return;
        }
      }

      if (this.documentChangeAffectsDrawings(payload)) {
        // Drawings/pictures overlay caches sheet-space bounds inside its spatial index. Invalidate
        // so the next render recomputes even if `getDrawingObjects` returns a stable array reference.
        this.drawingsOverlay.invalidateSpatialIndex();
        void this.renderDrawings();
      }
    });
    this.disposeFns.push(() => unsubscribeImages());

    // Optional dedicated event streams (if/when DocumentController adds them).
    const unsubscribeDrawings = this.document.on("drawings", (payload: any) => {
      if (this.disposed) return;
      const sheetId = typeof payload?.sheetId === "string" ? payload.sheetId : null;
      if (sheetId && sheetId !== this.getSheetId()) return;
      this.drawingsOverlay.invalidateSpatialIndex();
      void this.renderDrawings();
    });
    this.disposeFns.push(() => unsubscribeDrawings());

    const unsubscribeImagesEvent = this.document.on("images", () => {
      if (this.disposed) return;
      this.drawingsOverlay.clearImageCache();
      void this.renderDrawings();
    });
    this.disposeFns.push(() => unsubscribeImagesEvent());

    // Excel behavior: clicking another cell while editing should commit the edit and move selection.
    // We rely on the editor's blur-to-commit handler above (which fires when DesktopSharedGrid
    // focuses the container during pointer interactions).

    // Match SpreadsheetApp header sizing so cell hit targets line up.
    this.grid.renderer.setColWidth(0, 48);
    this.grid.renderer.setRowHeight(0, 24);

    this.drawingsCanvas = drawingsCanvas;
    const geom: GridGeometry = {
      cellOriginPx: (cell) => {
        const headerWidth =
          this.headerCols > 0 ? this.grid.renderer.scroll.cols.totalSize(this.headerCols) : 0;
        const headerHeight =
          this.headerRows > 0 ? this.grid.renderer.scroll.rows.totalSize(this.headerRows) : 0;
        const gridRow = cell.row + this.headerRows;
        const gridCol = cell.col + this.headerCols;
        return {
          x: this.grid.renderer.scroll.cols.positionOf(gridCol) - headerWidth,
          y: this.grid.renderer.scroll.rows.positionOf(gridRow) - headerHeight,
        };
      },
      cellSizePx: (cell) => {
        const gridRow = cell.row + this.headerRows;
        const gridCol = cell.col + this.headerCols;
        return {
          width: this.grid.renderer.getColWidth(gridCol),
          height: this.grid.renderer.getRowHeight(gridRow),
        };
      },
    };
    this.drawingsOverlay = new DrawingOverlay(
      drawingsCanvas,
      this.drawingsImages,
      geom,
      options.chartRenderer,
      () => this.renderDrawings(),
      this.container,
    );

    // Initial sizing (ResizeObserver will keep it updated).
    this.resizeToContainer();

    const initialZoom = clamp(options.initialZoom ?? 1, MIN_GRID_ZOOM, MAX_GRID_ZOOM);
    if (initialZoom !== 1) {
      this.grid.renderer.setZoom(initialZoom);
      this.grid.syncScrollbars();
    }
    this.lastZoom = this.grid.renderer.getZoom();
    this.container.dataset.zoom = String(this.lastZoom);

    // Apply frozen panes + axis overrides from the DocumentController sheet view.
    this.syncSheetViewFromDocument();

    const initialScroll = options.initialScroll ?? { scrollX: 0, scrollY: 0 };
    if (initialScroll.scrollX !== 0 || initialScroll.scrollY !== 0) {
      this.grid.scrollTo(initialScroll.scrollX, initialScroll.scrollY);
    } else {
      // Still set the dataset so tests can read a stable baseline.
      const scroll = this.grid.getScroll();
      this.container.dataset.scrollX = String(scroll.x);
      this.container.dataset.scrollY = String(scroll.y);
    }

    // From this point on, viewport callbacks can safely render overlays.
    this.rowsVersion = this.grid.renderer.scroll.rows.getVersion();
    this.colsVersion = this.grid.renderer.scroll.cols.getVersion();
    this.uiReady = true;
    void this.renderDrawings();

    // Re-apply view state when the document emits sheet view deltas (freeze panes, row/col sizes).
    const unsubscribeSheetView = this.document.on("change", (payload: any) => {
      const deltas = Array.isArray(payload?.sheetViewDeltas) ? payload.sheetViewDeltas : [];
      if (deltas.length === 0) return;
      const source = typeof payload?.source === "string" ? payload.source : "";
      // Avoid redundant re-sync into the same secondary pane after local axis resize/auto-fit edits
      // (we update the renderer directly during the drag). Other panes (primary grid) will still
      // observe the change and sync.
      if (source === "secondaryGridAxis") return;
      const sheetId = this.getSheetId();
      if (!deltas.some((delta: any) => String(delta?.sheetId ?? "") === sheetId)) return;
      this.syncSheetViewFromDocument();
    });
    this.disposeFns.push(() => unsubscribeSheetView());
    this.resizeObserver = new ResizeObserver(() => this.resizeToContainer());
    this.resizeObserver.observe(this.container);
  }

  destroy(): void {
    this.disposed = true;
    // Cancel any in-flight drawing renders and release cached ImageBitmaps.
    this.drawingsOverlay?.destroy?.();
    this.drawingsRenderQueued = false;
    this.drawingsRenderInProgress = false;
    this.flushPersistence();
    this.resizeObserver.disconnect();
    for (const dispose of this.disposeFns) dispose();
    this.disposeFns.length = 0;
    // Ensure we don't drop in-progress edits when the split pane is torn down
    // (e.g. disabling split view while editing).
    this.editor.commit("command");
    this.editingCell = null;
    this.editor.close();
    // Remove the detached textarea + event listeners so a referenced SecondaryGridView
    // instance doesn't retain DOM subtrees after teardown (tests/hot reload/split toggling).
    try {
      this.editor.destroy();
    } catch {
      // ignore
    }
    this.grid.destroy();
    if (this.ownsProvider) {
      try {
        this.provider.dispose();
      } catch {
        // ignore
      }
    }
    // Release canvas backing stores even if the SecondaryGridView instance is still referenced
    // after destroy (tests, hot reload, split-pane toggling). Detached canvases can otherwise
    // retain multi-megabyte buffers.
    try {
      for (const canvas of Array.from(this.container.querySelectorAll("canvas"))) {
        try {
          canvas.width = 0;
          canvas.height = 0;
        } catch {
          // ignore
        }
      }
    } catch {
      // ignore
    }
    // Remove any DOM we created (the container stays in place).
    this.container.replaceChildren();
  }

  /**
   * Commit any in-progress cell edit without moving selection.
   *
   * This is intended for "command" entry points (Save/Open/Close/Quit) so pending
   * edits in the secondary pane are not lost.
   */
  commitPendingEditsForCommand(): void {
    if (!this.editor.isOpen()) return;
    this.editor.commit("command");
  }

  private openEditor(request: { row: number; col: number; initialKey?: string }): void {
    if (this.editor.isOpen()) return;
    if (request.row < this.headerRows || request.col < this.headerCols) return;
    const rect = this.grid.getCellRect(request.row, request.col);
    if (!rect) return;

    const sheetId = this.getSheetId();
    const cell = { row: request.row - this.headerRows, col: request.col - this.headerCols };
    // In collab read-only roles (viewer/commenter) and other permission-restricted scenarios,
    // DocumentController silently filters cell edits via `canEditCell`. Guard here so the
    // secondary-pane editor does not open only to have its commit no-op.
    //
    // SpreadsheetApp has similar guards for the primary grid.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const canEditCell = (this.document as any).canEditCell as
      | ((cell: { sheetId: string; row: number; col: number }) => boolean)
      | null
      | undefined;
    if (typeof canEditCell === "function") {
      let allowed = true;
      try {
        allowed = Boolean(canEditCell({ sheetId, row: cell.row, col: cell.col }));
      } catch {
        allowed = true;
      }
      if (!allowed) {
        showCollabEditRejectedToast([
          { sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
        ]);
        return;
      }
    }

    const initialValue = request.initialKey ?? this.getCellInputText(cell);
    this.editingCell = cell;
    this.editor.open(cell, rect, initialValue, { cursor: "end" });
    this.onEditStateChange?.(true);
  }

  private getCellInputText(cell: { row: number; col: number }): string {
    const state = this.document.getCell(this.getSheetId(), cell) as { value: unknown; formula: string | null };
    if (state?.formula != null) return state.formula;
    if (isRichTextValue(state?.value)) return state.value.text;
    if (parseImageCellPayload(state?.value)) return "";
    if (state?.value != null) return String(state.value);
    return "";
  }

  private applyEdit(cell: { row: number; col: number }, rawValue: string): void {
    const sheetId = this.getSheetId();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const canEditCell = (this.document as any).canEditCell as
      | ((cell: { sheetId: string; row: number; col: number }) => boolean)
      | null
      | undefined;
    if (typeof canEditCell === "function") {
      let allowed = true;
      try {
        allowed = Boolean(canEditCell({ sheetId, row: cell.row, col: cell.col }));
      } catch {
        allowed = true;
      }
      if (!allowed) {
        showCollabEditRejectedToast([
          { sheetId, row: cell.row, col: cell.col, rejectionKind: "cell", rejectionReason: "permission" },
        ]);
        return;
      }
    }
    const original = this.document.getCell(sheetId, cell) as { value: unknown; formula: string | null };

    const originalInput = (() => {
      if (!original) return "";
      if (original.formula != null) return original.formula;
      if (isRichTextValue(original.value)) return original.value.text;
      if (parseImageCellPayload(original.value)) return "";
      if (original.value != null) return String(original.value);
      return "";
    })();

    if (rawValue === originalInput) return;
    if (rawValue.trim() === "") {
      this.document.clearCell(sheetId, cell, { label: "Clear cell" });
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
      this.document.setCellValue(sheetId, cell, updated, { label: "Edit cell" });
      return;
    }

    this.document.setCellInput(sheetId, cell, rawValue, { label: "Edit cell" });
  }

  private repositionEditor(): void {
    if (!this.editor.isOpen()) return;
    const cell = this.editingCell;
    if (!cell) return;
    const gridRow = cell.row + this.headerRows;
    const gridCol = cell.col + this.headerCols;
    const rect = this.grid.getCellRect(gridRow, gridCol);
    if (rect) this.editor.reposition(rect);
  }

  private resizeToContainer(): void {
    if (this.disposed) return;
    const width = this.container.clientWidth;
    const height = this.container.clientHeight;
    const dpr = window.devicePixelRatio || 1;
    this.grid.resize(width, height, dpr);
    this.drawingsOverlay.resize(this.getDrawingsViewport({ width, height, dpr }));
    void this.renderDrawings();
  }

  /**
   * Sync frozen panes + row/col size overrides from `DocumentController.getSheetView()`
   * into the secondary pane renderer.
   */
  syncSheetViewFromDocument(): void {
    const sheetId = this.getSheetId();
    if (sheetId !== this.sheetId) {
      this.sheetId = sheetId;
    }

    const view = this.document.getSheetView(sheetId) as {
      frozenRows?: number;
      frozenCols?: number;
      colWidths?: Record<string, number>;
      rowHeights?: Record<string, number>;
    } | null;

    const maxRows = Math.max(0, this.grid.renderer.scroll.getCounts().rowCount - this.headerRows);
    const maxCols = Math.max(0, this.grid.renderer.scroll.getCounts().colCount - this.headerCols);

    const normalizeFrozen = (value: unknown, max: number): number => {
      const num = Number(value);
      if (!Number.isFinite(num)) return 0;
      return Math.max(0, Math.min(Math.trunc(num), max));
    };

    const frozenRows = normalizeFrozen(view?.frozenRows, maxRows);
    const frozenCols = normalizeFrozen(view?.frozenCols, maxCols);
    this.sheetViewFrozen = { rows: frozenRows, cols: frozenCols };
    this.grid.renderer.setFrozen(this.headerRows + frozenRows, this.headerCols + frozenCols);

    const zoom = this.grid.renderer.getZoom();

    const nextCols = new Map<number, number>();
    for (const [key, value] of Object.entries(view?.colWidths ?? {})) {
      const col = Number(key);
      if (!Number.isInteger(col) || col < 0) continue;
      const size = Number(value);
      if (!Number.isFinite(size) || size <= 0) continue;
      nextCols.set(col, size);
    }

    const nextRows = new Map<number, number>();
    for (const [key, value] of Object.entries(view?.rowHeights ?? {})) {
      const row = Number(key);
      if (!Number.isInteger(row) || row < 0) continue;
      const size = Number(value);
      if (!Number.isFinite(size) || size <= 0) continue;
      nextRows.set(row, size);
    }

    // Batch apply to avoid N-per-index invalidations when many overrides exist.
    const colSizes = new Map<number, number>();
    for (let i = 0; i < this.headerCols; i += 1) {
      colSizes.set(i, this.grid.renderer.getColWidth(i));
    }
    for (const [col, base] of nextCols) {
      if (col >= maxCols) continue;
      colSizes.set(col + this.headerCols, base * zoom);
    }

    const rowSizes = new Map<number, number>();
    for (let i = 0; i < this.headerRows; i += 1) {
      rowSizes.set(i, this.grid.renderer.getRowHeight(i));
    }
    for (const [row, base] of nextRows) {
      if (row >= maxRows) continue;
      rowSizes.set(row + this.headerRows, base * zoom);
    }

    this.grid.renderer.applyAxisSizeOverrides({ rows: rowSizes, cols: colSizes }, { resetUnspecified: true });
    // Axis size overrides mutate the underlying GridGeometry while the `DrawingOverlay` spatial
    // index is keyed by the stable `geom` object reference. Invalidate cached bounds so the
    // overlay recomputes drawing positions under the new row/col sizes.
    this.drawingsOverlay.invalidateSpatialIndex();
    // Keep row/col version tracking aligned with the renderer so the next scroll event doesn't
    // redundantly invalidate the drawings spatial index for the same axis-size update.
    this.rowsVersion = this.grid.renderer.scroll.rows.getVersion();
    this.colsVersion = this.grid.renderer.scroll.cols.getVersion();
    this.grid.syncScrollbars();
    const scroll = this.grid.getScroll();
    this.container.dataset.scrollX = String(scroll.x);
    this.container.dataset.scrollY = String(scroll.y);
    this.repositionEditor();
    void this.renderDrawings();
  }

  private onAxisSizeChange(change: GridAxisSizeChange): void {
    // The shared-grid renderer updates sizes interactively during resize drags; invalidate the
    // drawing spatial index so anchors recompute with the new geometry before re-rendering.
    this.drawingsOverlay.invalidateSpatialIndex();
    // The axis resize already updated the renderer sizes; sync our cached axis versions so the
    // next scroll doesn't redundantly invalidate for the same resize.
    this.rowsVersion = this.grid.renderer.scroll.rows.getVersion();
    this.colsVersion = this.grid.renderer.scroll.cols.getVersion();
    const sheetId = this.getSheetId();
    const baseSize = change.size / change.zoom;
    const baseDefault = change.defaultSize / change.zoom;
    const isDefault = Math.abs(baseSize - baseDefault) < 1e-6;
    // Tag sheet-view mutations originating from this secondary pane so we can avoid redundant
    // view re-syncs back into the same renderer instance (the pane is already updated during the
    // resize drag). Other panes (primary grid) still need to observe the change and sync.
    const source = "secondaryGridAxis";

    if (change.kind === "col") {
      const docCol = change.index - this.headerCols;
      if (docCol < 0) return;
      const label = change.source === "autoFit" ? "Autofit Column Width" : "Resize Column";
      if (isDefault) {
        this.document.resetColWidth(sheetId, docCol, { label, source });
      } else {
        this.document.setColWidth(sheetId, docCol, baseSize, { label, source });
      }

      // Keep axis version tracking aligned with the interactive renderer updates.
      this.rowsVersion = this.grid.renderer.scroll.rows.getVersion();
      this.colsVersion = this.grid.renderer.scroll.cols.getVersion();

      void this.renderDrawings();
      return;
    }

    const docRow = change.index - this.headerRows;
    if (docRow < 0) return;
    const label = change.source === "autoFit" ? "Autofit Row Height" : "Resize Row";
    if (isDefault) {
      this.document.resetRowHeight(sheetId, docRow, { label, source });
    } else {
      this.document.setRowHeight(sheetId, docRow, baseSize, { label, source });
    }
    void this.renderDrawings();
  }

  private onFillCommit(event: FillCommitEvent): void {
    const sheetId = this.getSheetId();
    const { sourceRange, targetRange, mode } = event;

    const prevSelection = this.grid.renderer.getSelection();
    const prevRanges = this.grid.renderer.getSelectionRanges().map((r) => ({ ...r }));
    const prevActiveIndex = this.grid.renderer.getActiveSelectionIndex();

    const toFillRange = (range: GridCellRange): FillEngineRange | null => {
      const startRow = Math.max(0, range.startRow - this.headerRows);
      const endRow = Math.max(0, range.endRow - this.headerRows);
      const startCol = Math.max(0, range.startCol - this.headerCols);
      const endCol = Math.max(0, range.endCol - this.headerCols);
      if (endRow <= startRow || endCol <= startCol) return null;
      return { startRow, endRow, startCol, endCol };
    };

    const source = toFillRange(sourceRange);
    const target = toFillRange(targetRange);
    if (!source || !target) return;

    // In collab read-only roles (viewer/commenter), block fill operations in the secondary pane.
    // DocumentController silently filters disallowed cell deltas via `canEditCell`, so guard here
    // to avoid a confusing "selection expanded but nothing happened" outcome.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const canEditCell = (this.document as any).canEditCell as
      | ((cell: { sheetId: string; row: number; col: number }) => boolean)
      | null
      | undefined;
    if (typeof canEditCell === "function") {
      let allowed = true;
      try {
        allowed = Boolean(canEditCell({ sheetId, row: source.startRow, col: source.startCol }));
      } catch {
        allowed = true;
      }
      if (!allowed) {
        showCollabEditRejectedToast([
          { sheetId, row: source.startRow, col: source.startCol, rejectionKind: "cell", rejectionReason: "permission" },
        ]);

        // DesktopSharedGrid will still expand selection to the dragged target range even if
        // we skip applying edits. Suppress selection sync callbacks and restore the prior
        // selection on the next microtask turn so split-view panes stay consistent.
        this.suppressSelectionCallbacks = true;
        queueMicrotask(() => {
          try {
            this.grid.setSelectionRanges(prevRanges, {
              activeIndex: prevActiveIndex,
              activeCell: prevSelection,
              scrollIntoView: false,
            });
          } finally {
            this.suppressSelectionCallbacks = false;
          }
        });
        return;
      }
    }

    const sourceCells = (source.endRow - source.startRow) * (source.endCol - source.startCol);
    const targetCells = (target.endRow - target.startRow) * (target.endCol - target.startCol);
    if (sourceCells > MAX_FILL_CELLS || targetCells > MAX_FILL_CELLS) {
      try {
        showToast(
          `Fill range too large (>${MAX_FILL_CELLS.toLocaleString()} cells). Select fewer cells and try again.`,
          "warning",
        );
      } catch {
        // `showToast` requires a #toast-root; unit tests don't always include it.
      }

      // DesktopSharedGrid will still expand selection to the dragged target range even if
      // we skip applying edits. Suppress selection sync callbacks and restore the prior
      // selection on the next microtask turn so split-view panes stay consistent.
      this.suppressSelectionCallbacks = true;
      queueMicrotask(() => {
        try {
          this.grid.setSelectionRanges(prevRanges, {
            activeIndex: prevActiveIndex,
            activeCell: prevSelection,
            scrollIntoView: false,
          });
        } finally {
          this.suppressSelectionCallbacks = false;
        }
      });
      return;
    }
    const fillCoordScratch = { row: 0, col: 0 };
    applyFillCommitToDocumentController({
      document: this.document,
      sheetId,
      sourceRange: source,
      targetRange: target,
      mode,
      getCellComputedValue: (row, col) => {
        fillCoordScratch.row = row;
        fillCoordScratch.col = col;
        return this.getComputedValue(fillCoordScratch) as any;
      },
    });

    // Ensure the secondary pane repaints immediately after the mutation. The primary
    // pane observes DocumentController changes via its shared provider.
    this.onRequestRefresh?.();
    this.grid.renderer.requestRender();
  }

  private advanceSelectionAfterEdit(commit: EditorCommit): void {
    if (commit.reason !== "enter" && commit.reason !== "tab") return;

    const renderer = this.grid.renderer;
    const selection = renderer.getSelection();
    const range = renderer.getSelectionRange();
    const ranges = renderer.getSelectionRanges();
    const activeIndex = renderer.getActiveSelectionIndex();
    const { rowCount, colCount } = renderer.scroll.getCounts();
    if (rowCount === 0 || colCount === 0) return;

    // Clamp navigation so we never land in the header region.
    const dataStartRow = this.headerRows >= rowCount ? 0 : this.headerRows;
    const dataStartCol = this.headerCols >= colCount ? 0 : this.headerCols;

    const rangeArea = (r: CellRange) => Math.max(0, r.endRow - r.startRow) * Math.max(0, r.endCol - r.startCol);

    // Excel-like behavior: when a multi-cell selection exists, Enter/Tab should move
    // *within* the selection range (wrapping) instead of collapsing selection.
    if (range && rangeArea(range) > 1) {
      const current = selection ?? { row: range.startRow, col: range.startCol };
      const activeRow = clamp(current.row, range.startRow, range.endRow - 1);
      const activeCol = clamp(current.col, range.startCol, range.endCol - 1);
      const backward = commit.shift;

      let nextRow = activeRow;
      let nextCol = activeCol;

      if (commit.reason === "tab") {
        if (!backward) {
          if (activeCol + 1 < range.endCol) {
            nextCol = activeCol + 1;
          } else if (activeRow + 1 < range.endRow) {
            nextRow = activeRow + 1;
            nextCol = range.startCol;
          } else {
            nextRow = range.startRow;
            nextCol = range.startCol;
          }
        } else {
          if (activeCol - 1 >= range.startCol) {
            nextCol = activeCol - 1;
          } else if (activeRow - 1 >= range.startRow) {
            nextRow = activeRow - 1;
            nextCol = range.endCol - 1;
          } else {
            nextRow = range.endRow - 1;
            nextCol = range.endCol - 1;
          }
        }
      } else {
        if (!backward) {
          if (activeRow + 1 < range.endRow) {
            nextRow = activeRow + 1;
          } else if (activeCol + 1 < range.endCol) {
            nextRow = range.startRow;
            nextCol = activeCol + 1;
          } else {
            nextRow = range.startRow;
            nextCol = range.startCol;
          }
        } else {
          if (activeRow - 1 >= range.startRow) {
            nextRow = activeRow - 1;
          } else if (activeCol - 1 >= range.startCol) {
            nextRow = range.endRow - 1;
            nextCol = activeCol - 1;
          } else {
            nextRow = range.endRow - 1;
            nextCol = range.endCol - 1;
          }
        }
      }

      nextRow = Math.max(dataStartRow, Math.min(rowCount - 1, nextRow));
      nextCol = Math.max(dataStartCol, Math.min(colCount - 1, nextCol));

      const updatedRanges = ranges.length === 0 ? [range] : ranges;
      const safeIndex = Math.max(0, Math.min(updatedRanges.length - 1, activeIndex));
      this.grid.setSelectionRanges(updatedRanges, {
        activeIndex: safeIndex,
        activeCell: { row: nextRow, col: nextCol },
      });
      return;
    }

    const current = selection ?? { row: commit.cell.row + this.headerRows, col: commit.cell.col + this.headerCols };
    let nextRow = current.row;
    let nextCol = current.col;

    if (commit.reason === "enter") {
      nextRow = current.row + (commit.shift ? -1 : 1);
    } else {
      nextCol = current.col + (commit.shift ? -1 : 1);
    }

    nextRow = Math.max(dataStartRow, Math.min(rowCount - 1, nextRow));
    nextCol = Math.max(dataStartCol, Math.min(colCount - 1, nextCol));

    this.grid.setSelectionRanges(
      [{ startRow: nextRow, endRow: nextRow + 1, startCol: nextCol, endCol: nextCol + 1 }],
      { activeIndex: 0, activeCell: { row: nextRow, col: nextCol } },
    );
  }

  private schedulePersistScroll(scroll: ScrollState): void {
    if (!this.persistScroll) return;
    this.pendingScroll = scroll;

    if (this.scrollPersistTimer != null) {
      window.clearTimeout(this.scrollPersistTimer);
    }
    this.scrollPersistTimer = window.setTimeout(() => {
      this.scrollPersistTimer = null;
      const pending = this.pendingScroll;
      this.pendingScroll = null;
      if (!pending) return;
      this.persistScroll?.(pending);
    }, this.persistDebounceMs);
  }

  private schedulePersistZoom(zoom: number): void {
    if (!this.persistZoom) return;
    this.pendingZoom = zoom;

    if (this.zoomPersistTimer != null) {
      window.clearTimeout(this.zoomPersistTimer);
    }
    this.zoomPersistTimer = window.setTimeout(() => {
      this.zoomPersistTimer = null;
      const pending = this.pendingZoom;
      this.pendingZoom = null;
      if (pending == null) return;
      this.persistZoom?.(pending);
    }, this.persistDebounceMs);
  }

  flushPersistence(): void {
    if (this.scrollPersistTimer != null) {
      window.clearTimeout(this.scrollPersistTimer);
      this.scrollPersistTimer = null;
    }
    if (this.zoomPersistTimer != null) {
      window.clearTimeout(this.zoomPersistTimer);
      this.zoomPersistTimer = null;
    }

    if (this.pendingScroll) {
      this.persistScroll?.(this.pendingScroll);
      this.pendingScroll = null;
    }
    if (this.pendingZoom != null) {
      this.persistZoom?.(this.pendingZoom);
      this.pendingZoom = null;
    }
  }

  private documentChangeAffectsDrawings(payload: any): boolean {
    if (!payload || typeof payload !== "object") return false;

    const source = typeof payload?.source === "string" ? payload.source : "";
    // Axis resize events originating from this secondary pane are already handled by the
    // `onAxisSizeChange` callback (which invalidates drawing geometry and triggers a render).
    // Skip the redundant render work that would otherwise be triggered by the resulting
    // `sheetViewDeltas` change event.
    if (source === "secondaryGridAxis") return false;
    // Applying a new document snapshot can replace the drawing layer entirely.
    if (source === "applyState") return true;
    // Some integrations may publish drawings/images updates with a dedicated source tag.
    if (source === "drawings" || source === "images") return true;

    const sheetId = this.getSheetId();
    const matchesSheet = (delta: any): boolean => String(delta?.sheetId ?? "") === sheetId;
    const touchesSheet = (deltas: any): boolean => Array.isArray(deltas) && deltas.some(matchesSheet);

    // NOTE: sheet view deltas (frozen panes / row+col sizes / drawing metadata) are synced via
    // `syncSheetViewFromDocument()` in a dedicated listener, which also triggers a drawings render.
    // Avoid treating them as drawing-affecting here to prevent redundant renders with stale geometry
    // during axis resize / sheet-view sync flows.

    // Sheet meta/order changes can change the active sheet or invalidate cached geometry.
    if (touchesSheet(payload?.sheetMetaDeltas) || payload?.sheetOrderDelta) return true;

    if (
      touchesSheet(payload?.drawingsDeltas) ||
      touchesSheet(payload?.drawingDeltas) ||
      touchesSheet(payload?.sheetDrawingsDeltas) ||
      touchesSheet(payload?.sheetDrawingDeltas)
    ) {
      return true;
    }

    // Image updates may be workbook-wide; re-render so any referenced bitmaps refresh.
    // NOTE: DocumentController always includes `imageDeltas: []` in change payloads, even when no
    // images changed. Guard on non-empty arrays so unrelated changes don't trigger extra drawings
    // renders.
    const imageDeltas = Array.isArray(payload?.imageDeltas)
      ? payload.imageDeltas
      : Array.isArray(payload?.imagesDeltas)
        ? payload.imagesDeltas
        : null;
    if (imageDeltas && imageDeltas.length > 0) return true;

    if (payload?.drawingsChanged === true || payload?.imagesChanged === true) return true;

    return false;
  }

  private getDrawingsViewport(override?: { width: number; height: number; dpr: number }): DrawingsViewport {
    const scroll = this.grid.getScroll();
    const width = override?.width ?? this.container.clientWidth;
    const height = override?.height ?? this.container.clientHeight;
    const dpr = override?.dpr ?? (window.devicePixelRatio || 1);
    const zoom = this.grid.renderer.getZoom();
    const viewport = this.grid.renderer.scroll.getViewportState();
    const headerWidth = this.headerCols > 0 ? this.grid.renderer.scroll.cols.totalSize(this.headerCols) : 0;
    const headerHeight = this.headerRows > 0 ? this.grid.renderer.scroll.rows.totalSize(this.headerRows) : 0;
    // Clamp offsets to the visible viewport so extreme split ratios (very narrow panes) don't
    // produce header offsets larger than the canvas size.
    const headerOffsetX = Math.min(headerWidth, width);
    const headerOffsetY = Math.min(headerHeight, height);
    return {
      scrollX: scroll.x,
      scrollY: scroll.y,
      width,
      height,
      dpr,
      zoom,
      frozenRows: this.sheetViewFrozen.rows,
      frozenCols: this.sheetViewFrozen.cols,
      headerOffsetX,
      headerOffsetY,
      frozenWidthPx: viewport.frozenWidth,
      frozenHeightPx: viewport.frozenHeight,
    };
  }

  private renderDrawings(): void {
    if (this.disposed) return;
    if (this.drawingsRenderInProgress) {
      this.drawingsRenderQueued = true;
      return;
    }

    this.drawingsRenderInProgress = true;
    try {
      do {
        if (this.disposed) return;
        this.drawingsRenderQueued = false;
        const sheetId = this.getSheetId();
        const objects = this.getDrawingObjects(sheetId);
        const viewport = this.getDrawingsViewport();
        const selectedId = this.getSelectedDrawingId?.() ?? null;
        this.drawingsOverlay.setSelectedId(selectedId);
        // `DrawingOverlay.render()` is synchronous, but some unit tests stub it with an async mock
        // (returning a Promise). Handle both so rejected promises don't surface as unhandled
        // rejections during scroll/resize-driven repaint storms.
        const result = this.drawingsOverlay.render(objects, viewport) as unknown;
        if (typeof (result as { then?: unknown } | null)?.then === "function") {
          void Promise.resolve(result).catch(() => {
            // Best-effort: drawing overlay should never break grid interactions.
          });
        }
      } while (this.drawingsRenderQueued);
    } catch {
      // Best-effort: drawing overlay should never break grid interactions.
    } finally {
      this.drawingsRenderInProgress = false;
    }
  }
}
