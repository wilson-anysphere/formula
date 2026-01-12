import type { CellRange, CellRange as GridCellRange, CellRichText, FillCommitEvent, GridAxisSizeChange } from "@formula/grid";
import type { CellRange as FillEngineRange } from "@formula/fill-engine";
import type { DocumentController } from "../../document/documentController.js";
import { applyFillCommitToDocumentController } from "../../fill/applyFillCommit";
import { CellEditorOverlay } from "../../editor/cellEditorOverlay.js";
import { DesktopSharedGrid, type DesktopSharedGridCallbacks } from "../shared/desktopSharedGrid.js";
import { DocumentCellProvider } from "../shared/documentCellProvider.js";
import { applyPlainTextEdit } from "../text/rich-text/edit.js";
import { navigateSelectionByKey } from "../../selection/navigation.js";
import { buildSelection } from "../../selection/selection.js";
import type { GridLimits, Range, SelectionState, UsedRangeProvider } from "../../selection/types";

type ScrollState = { scrollX: number; scrollY: number };

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

  private readonly document: DocumentController;
  private readonly getSheetId: () => string;
  private readonly getComputedValue: (cell: { row: number; col: number }) => string | number | boolean | null;
  private readonly onRequestRefresh?: () => void;
  private readonly headerRows = 1;
  private readonly headerCols = 1;
  private readonly editor: CellEditorOverlay;
  private editingCell: { row: number; col: number } | null = null;

  private sheetId: string;

  private readonly resizeObserver: ResizeObserver;
  private readonly disposeFns: Array<() => void> = [];

  private readonly persistDebounceMs: number;
  private pendingScroll: ScrollState | null = null;
  private scrollPersistTimer: number | null = null;
  private pendingZoom: number | null = null;
  private zoomPersistTimer: number | null = null;
  private lastZoom = 1;

  private readonly persistScroll?: (scroll: ScrollState) => void;
  private readonly persistZoom?: (zoom: number) => void;

  constructor(options: {
    container: HTMLElement;
    /**
     * Optional provider to reuse (e.g. the primary shared-grid provider).
     *
     * When supplied, `document`/`getSheetId`/`showFormulas`/`getComputedValue` are
     * ignored for provider creation, but still accepted for convenience.
     */
    provider?: DocumentCellProvider;
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
     * Optional hook to refresh non-grid UI when the secondary pane mutates the document
     * (e.g. formula bar, charts, auditing overlays in the primary pane).
     */
    onRequestRefresh?: () => void;
  }) {
    this.container = options.container;
    this.document = options.document;
    this.getSheetId = options.getSheetId;
    this.getComputedValue = options.getComputedValue;
    this.persistScroll = options.persistScroll;
    this.persistZoom = options.persistZoom;
    this.persistDebounceMs = options.persistDebounceMs ?? 150;
    this.sheetId = options.getSheetId();
    this.onRequestRefresh = options.onRequestRefresh;

    // Clear any placeholder content from the split-view scaffolding.
    this.container.replaceChildren();

    const gridCanvas = document.createElement("canvas");
    gridCanvas.className = "grid-canvas grid-canvas--base";
    gridCanvas.setAttribute("aria-hidden", "true");

    const contentCanvas = document.createElement("canvas");
    contentCanvas.className = "grid-canvas grid-canvas--content";
    contentCanvas.setAttribute("aria-hidden", "true");

    const selectionCanvas = document.createElement("canvas");
    selectionCanvas.className = "grid-canvas grid-canvas--selection";
    selectionCanvas.setAttribute("aria-hidden", "true");

    this.container.appendChild(gridCanvas);
    this.container.appendChild(contentCanvas);
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
        this.editingCell = null;
        this.applyEdit(commit.cell, commit.value);
        if (commit.reason !== "command") {
          const gridSelection = this.grid.renderer.getSelection();
          const gridRanges = this.grid.renderer.getSelectionRanges();
          const activeIndex = this.grid.renderer.getActiveSelectionIndex();
          const counts = this.grid.renderer.scroll.getCounts();
          const limits: GridLimits = {
            maxRows: Math.max(1, counts.rowCount - this.headerRows),
            maxCols: Math.max(1, counts.colCount - this.headerCols),
          };

          const docRanges: Range[] =
            gridRanges.length > 0
              ? gridRanges.map((r) => ({
                  startRow: Math.max(0, r.startRow - this.headerRows),
                  endRow: Math.max(0, r.endRow - this.headerRows - 1),
                  startCol: Math.max(0, r.startCol - this.headerCols),
                  endCol: Math.max(0, r.endCol - this.headerCols - 1),
                }))
              : [
                  {
                    startRow: commit.cell.row,
                    endRow: commit.cell.row,
                    startCol: commit.cell.col,
                    endCol: commit.cell.col,
                  },
                ];

          const activeCell = gridSelection
            ? { row: Math.max(0, gridSelection.row - this.headerRows), col: Math.max(0, gridSelection.col - this.headerCols) }
            : { ...commit.cell };

          const selectionState: SelectionState = buildSelection(
            {
              ranges: docRanges,
              active: activeCell,
              anchor: activeCell,
              activeRangeIndex: Math.max(0, Math.min(activeIndex, docRanges.length - 1)),
            },
            limits
          );

          const sheetId = this.getSheetId();
          const data: UsedRangeProvider = {
            getUsedRange: () => this.document.getUsedRange(sheetId),
            isCellEmpty: (cell) => {
              const state = this.document.getCell(sheetId, cell) as { value: unknown; formula: string | null } | null;
              return state?.value == null && state?.formula == null;
            },
            // Shared-grid mode currently doesn't hide outline rows/cols; match SpreadsheetApp behavior.
            isRowHidden: () => false,
            isColHidden: () => false,
          };

          const next = navigateSelectionByKey(
            selectionState,
            commit.reason === "enter" ? "Enter" : "Tab",
            { shift: commit.shift, primary: false },
            data,
            limits
          );

          if (next) {
            const nextGridRanges: CellRange[] = next.ranges.map((r) => ({
              startRow: r.startRow + this.headerRows,
              endRow: r.endRow + this.headerRows + 1,
              startCol: r.startCol + this.headerCols,
              endCol: r.endCol + this.headerCols + 1,
            }));
            const nextActive = { row: next.active.row + this.headerRows, col: next.active.col + this.headerCols };
            this.grid.setSelectionRanges(nextGridRanges, { activeIndex: next.activeRangeIndex, activeCell: nextActive });
          }
        }
        this.onRequestRefresh?.();
        focusWithoutScroll(this.container);
      },
      onCancel: () => {
        this.editingCell = null;
        focusWithoutScroll(this.container);
      }
    });

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
        getCommentMeta: options.getCommentMeta
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

          this.repositionEditor();
          externalCallbacks.onScroll?.(scroll, viewport);
        },
        onAxisSizeChange: (change) => {
          this.onAxisSizeChange(change);
          externalCallbacks.onAxisSizeChange?.(change);
        },
        onSelectionChange: (selection) => {
          options.onSelectionChange?.(selection);
          externalCallbacks.onSelectionChange?.(selection);
        },
        onSelectionRangeChange: (range) => {
          options.onSelectionRangeChange?.(range);
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

    // Clicking the canvas should behave like clicking a spreadsheet surface: focus the pane so
    // keyboard navigation/editing works immediately.
    const onPointerDown = () => focusWithoutScroll(this.container);
    this.container.addEventListener("pointerdown", onPointerDown);
    this.disposeFns.push(() => this.container.removeEventListener("pointerdown", onPointerDown));

    // Match SpreadsheetApp header sizing so cell hit targets line up.
    this.grid.renderer.setColWidth(0, 48);
    this.grid.renderer.setRowHeight(0, 24);

    // Initial sizing (ResizeObserver will keep it updated).
    this.resizeToContainer();

    const initialZoom = clamp(options.initialZoom ?? 1, 0.25, 4);
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

    // Re-apply view state when the document emits sheet view deltas (freeze panes, row/col sizes).
    const unsubscribeSheetView = this.document.on("change", (payload: any) => {
      const deltas = Array.isArray(payload?.sheetViewDeltas) ? payload.sheetViewDeltas : [];
      if (deltas.length === 0) return;
      const sheetId = this.getSheetId();
      if (!deltas.some((delta: any) => String(delta?.sheetId ?? "") === sheetId)) return;
      this.syncSheetViewFromDocument();
    });
    this.disposeFns.push(() => unsubscribeSheetView());
    this.resizeObserver = new ResizeObserver(() => this.resizeToContainer());
    this.resizeObserver.observe(this.container);
  }

  destroy(): void {
    this.flushPersistence();
    this.resizeObserver.disconnect();
    for (const dispose of this.disposeFns) dispose();
    this.disposeFns.length = 0;
    this.editingCell = null;
    this.editor.close();
    this.grid.destroy();
    // Remove any DOM we created (the container stays in place).
    this.container.replaceChildren();
  }

  private openEditor(request: { row: number; col: number; initialKey?: string }): void {
    if (this.editor.isOpen()) return;
    if (request.row < this.headerRows || request.col < this.headerCols) return;
    const rect = this.grid.getCellRect(request.row, request.col);
    if (!rect) return;

    const cell = { row: request.row - this.headerRows, col: request.col - this.headerCols };
    const initialValue = request.initialKey ?? this.getCellInputText(cell);
    this.editingCell = cell;
    this.editor.open(cell, rect, initialValue, { cursor: "end" });
  }

  private getCellInputText(cell: { row: number; col: number }): string {
    const state = this.document.getCell(this.getSheetId(), cell) as { value: unknown; formula: string | null };
    if (state?.formula != null) return state.formula;
    if (isRichTextValue(state?.value)) return state.value.text;
    if (state?.value != null) return String(state.value);
    return "";
  }

  private applyEdit(cell: { row: number; col: number }, rawValue: string): void {
    const sheetId = this.getSheetId();
    const original = this.document.getCell(sheetId, cell) as { value: unknown; formula: string | null };
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
    const width = this.container.clientWidth;
    const height = this.container.clientHeight;
    const dpr = window.devicePixelRatio || 1;
    this.grid.resize(width, height, dpr);
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

    this.grid.syncScrollbars();
    const scroll = this.grid.getScroll();
    this.container.dataset.scrollX = String(scroll.x);
    this.container.dataset.scrollY = String(scroll.y);
    this.repositionEditor();
  }

  private onAxisSizeChange(change: GridAxisSizeChange): void {
    const sheetId = this.getSheetId();
    const baseSize = change.size / change.zoom;
    const baseDefault = change.defaultSize / change.zoom;
    const isDefault = Math.abs(baseSize - baseDefault) < 1e-6;

    if (change.kind === "col") {
      const docCol = change.index - this.headerCols;
      if (docCol < 0) return;
      const label = change.source === "autoFit" ? "Autofit Column Width" : "Resize Column";
      if (isDefault) {
        this.document.resetColWidth(sheetId, docCol, { label });
      } else {
        this.document.setColWidth(sheetId, docCol, baseSize, { label });
      }
      return;
    }

    const docRow = change.index - this.headerRows;
    if (docRow < 0) return;
    const label = change.source === "autoFit" ? "Autofit Row Height" : "Resize Row";
    if (isDefault) {
      this.document.resetRowHeight(sheetId, docRow, { label });
    } else {
      this.document.setRowHeight(sheetId, docRow, baseSize, { label });
    }
  }

  private onFillCommit(event: FillCommitEvent): void {
    const sheetId = this.getSheetId();
    const { sourceRange, targetRange, mode } = event;

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

    applyFillCommitToDocumentController({
      document: this.document,
      sheetId,
      sourceRange: source,
      targetRange: target,
      mode,
      getCellComputedValue: (row, col) => this.getComputedValue({ row, col }) as any,
    });

    // Ensure the secondary pane repaints immediately after the mutation. The primary
    // pane observes DocumentController changes via its shared provider.
    this.onRequestRefresh?.();
    this.grid.renderer.requestRender();
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
}
