import type {
  CellProvider,
  CellRange,
  FillCommitEvent,
  FillMode,
  GridAxisSizeChange,
  GridPerfStats,
  GridViewportState,
  ScrollToCellAlign
} from "@formula/grid";
import { CanvasGridRenderer, computeScrollbarThumb, resolveGridThemeFromCssVars } from "@formula/grid";

export type DesktopGridInteractionMode = "default" | "rangeSelection";

export interface DesktopSharedGridCallbacks {
  onScroll?: (scroll: { x: number; y: number }, viewport: GridViewportState) => void;
  onSelectionChange?: (selection: { row: number; col: number } | null) => void;
  onSelectionRangeChange?: (range: CellRange | null) => void;
  onRequestCellEdit?: (request: { row: number; col: number; initialKey?: string }) => void;
  onAxisSizeChange?: (change: GridAxisSizeChange) => void;

  onRangeSelectionStart?: (range: CellRange) => void;
  onRangeSelectionChange?: (range: CellRange) => void;
  onRangeSelectionEnd?: (range: CellRange) => void;

  /**
   * Called when the user commits a fill handle drag.
   *
   * The `targetRange` excludes the `sourceRange` (it only includes the cells that
   * should be written).
   */
  onFillCommit?: (event: FillCommitEvent) => void | Promise<void>;
}

type ResizeHit = { kind: "col"; index: number } | { kind: "row"; index: number };

type ResizeDragState =
  | { kind: "col"; index: number; startClient: number; startSize: number }
  | { kind: "row"; index: number; startClient: number; startSize: number };

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function clampIndex(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return clamp(Math.trunc(value), min, max);
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

function toA1Address(row0: number, col0: number): string {
  return `${toColumnName(col0)}${row0 + 1}`;
}

function rangesEqual(a: CellRange | null, b: CellRange | null): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return a.startRow === b.startRow && a.endRow === b.endRow && a.startCol === b.startCol && a.endCol === b.endCol;
}

function formatCellDisplayText(value: string | number | boolean | null): string {
  if (value === null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function describeCell(
  selection: { row: number; col: number } | null,
  range: CellRange | null,
  provider: CellProvider,
  headerRows: number,
  headerCols: number
): string {
  if (!selection) return "No cell selected.";

  const row0 = selection.row - headerRows;
  const col0 = selection.col - headerCols;
  const address =
    row0 >= 0 && col0 >= 0 ? toA1Address(row0, col0) : `row ${selection.row + 1}, column ${selection.col + 1}`;

  const cell = provider.getCell(selection.row, selection.col);
  const valueText = formatCellDisplayText(cell?.value ?? null);
  const valueDescription = valueText.trim() === "" ? "blank" : valueText;

  let selectionDescription = "none";
  if (range) {
    const startRow0 = range.startRow - headerRows;
    const startCol0 = range.startCol - headerCols;
    const endRow0 = range.endRow - headerRows - 1;
    const endCol0 = range.endCol - headerCols - 1;
    if (startRow0 >= 0 && startCol0 >= 0 && endRow0 >= 0 && endCol0 >= 0) {
      const start = toA1Address(startRow0, startCol0);
      const end = toA1Address(endRow0, endCol0);
      selectionDescription = start === end ? start : `${start}:${end}`;
    } else {
      selectionDescription = `row ${range.startRow + 1}, column ${range.startCol + 1}`;
    }
  }

  return `Active cell ${address}, value ${valueDescription}. Selection ${selectionDescription}.`;
}

export class DesktopSharedGrid {
  readonly renderer: CanvasGridRenderer;

  private readonly container: HTMLElement;
  private readonly provider: CellProvider;
  private readonly callbacks: DesktopSharedGridCallbacks;

  private readonly gridCanvas: HTMLCanvasElement;
  private readonly contentCanvas: HTMLCanvasElement;
  private readonly selectionCanvas: HTMLCanvasElement;

  private readonly vTrack: HTMLDivElement;
  private readonly vThumb: HTMLDivElement;
  private readonly hTrack: HTMLDivElement;
  private readonly hThumb: HTMLDivElement;

  private readonly frozenRows: number;
  private readonly frozenCols: number;
  private readonly headerRows: number;
  private readonly headerCols: number;

  private interactionMode: DesktopGridInteractionMode = "default";

  private selectionAnchor: { row: number; col: number } | null = null;
  private keyboardAnchor: { row: number; col: number } | null = null;
  private selectionPointerId: number | null = null;
  private transientRange: CellRange | null = null;
  private lastPointerViewport: { x: number; y: number } | null = null;
  private autoScrollFrame: number | null = null;

  private dragMode: "selection" | "fillHandle" | null = null;
  private fillHandleState: {
    source: CellRange;
    target: CellRange;
    mode: FillMode;
    previewTarget: CellRange | null;
    endCell: { row: number; col: number };
  } | null = null;

  private resizePointerId: number | null = null;
  private resizeDrag: ResizeDragState | null = null;

  private lastAnnounced: { selection: { row: number; col: number } | null; range: CellRange | null } = {
    selection: null,
    range: null
  };

  private readonly a11yStatusId: string;
  private readonly a11yStatusEl: HTMLDivElement;

  private disposeFns: Array<() => void> = [];

  constructor(options: {
    container: HTMLElement;
    provider: CellProvider;
    rowCount: number;
    colCount: number;
    canvases: { grid: HTMLCanvasElement; content: HTMLCanvasElement; selection: HTMLCanvasElement };
    scrollbars: { vTrack: HTMLDivElement; vThumb: HTMLDivElement; hTrack: HTMLDivElement; hThumb: HTMLDivElement };
    frozenRows?: number;
    frozenCols?: number;
    defaultRowHeight?: number;
    defaultColWidth?: number;
    prefetchOverscanRows?: number;
    prefetchOverscanCols?: number;
    enableResize?: boolean;
    enableKeyboard?: boolean;
    enableWheel?: boolean;
    callbacks?: DesktopSharedGridCallbacks;
  }) {
    this.container = options.container;
    this.provider = options.provider;
    this.callbacks = options.callbacks ?? {};

    this.gridCanvas = options.canvases.grid;
    this.contentCanvas = options.canvases.content;
    this.selectionCanvas = options.canvases.selection;

    this.vTrack = options.scrollbars.vTrack;
    this.vThumb = options.scrollbars.vThumb;
    this.hTrack = options.scrollbars.hTrack;
    this.hThumb = options.scrollbars.hThumb;

    this.frozenRows = options.frozenRows ?? 0;
    this.frozenCols = options.frozenCols ?? 0;
    this.headerRows = this.frozenRows > 0 ? 1 : 0;
    this.headerCols = this.frozenCols > 0 ? 1 : 0;

    this.renderer = new CanvasGridRenderer({
      provider: options.provider,
      rowCount: options.rowCount,
      colCount: options.colCount,
      defaultRowHeight: options.defaultRowHeight,
      defaultColWidth: options.defaultColWidth,
      prefetchOverscanRows: options.prefetchOverscanRows,
      prefetchOverscanCols: options.prefetchOverscanCols
    });

    // Match the React CanvasGrid behaviour: enable resize unless explicitly disabled.
    const enableResize = options.enableResize ?? true;

    this.a11yStatusId = `desktop-grid-a11y-${Math.random().toString(36).slice(2)}`;
    this.a11yStatusEl = document.createElement("div");
    this.a11yStatusEl.id = this.a11yStatusId;
    this.a11yStatusEl.dataset.testid = "canvas-grid-a11y-status";
    this.a11yStatusEl.setAttribute("role", "status");
    this.a11yStatusEl.setAttribute("aria-live", "polite");
    this.a11yStatusEl.setAttribute("aria-atomic", "true");
    this.a11yStatusEl.style.position = "absolute";
    this.a11yStatusEl.style.width = "1px";
    this.a11yStatusEl.style.height = "1px";
    this.a11yStatusEl.style.padding = "0";
    this.a11yStatusEl.style.margin = "-1px";
    this.a11yStatusEl.style.overflow = "hidden";
    this.a11yStatusEl.style.clip = "rect(0, 0, 0, 0)";
    this.a11yStatusEl.style.whiteSpace = "nowrap";
    this.a11yStatusEl.style.border = "0";
    this.a11yStatusEl.textContent = describeCell(null, null, this.provider, this.headerRows, this.headerCols);
    this.container.appendChild(this.a11yStatusEl);

    this.container.setAttribute("role", "grid");
    this.container.setAttribute("aria-rowcount", String(options.rowCount));
    this.container.setAttribute("aria-colcount", String(options.colCount));
    this.container.setAttribute("aria-multiselectable", "true");
    this.container.setAttribute("aria-describedby", this.a11yStatusId);
    this.container.style.touchAction = "none";

    // Attempt to resolve theme CSS vars if the host app defines them.
    this.renderer.setTheme(resolveGridThemeFromCssVars(this.container));

    this.renderer.attach({
      grid: this.gridCanvas,
      content: this.contentCanvas,
      selection: this.selectionCanvas
    });
    this.renderer.setFrozen(this.frozenRows, this.frozenCols);
    this.renderer.setFillHandleEnabled(this.interactionMode === "default" && Boolean(this.callbacks.onFillCommit));

    // Ensure the selection canvas captures pointer events (background/content are paint-only).
    this.gridCanvas.style.pointerEvents = "none";
    this.contentCanvas.style.pointerEvents = "none";
    this.selectionCanvas.style.pointerEvents = "auto";

    const enableWheel = options.enableWheel ?? true;
    const enableKeyboard = options.enableKeyboard ?? true;

    if (enableWheel) this.installWheelHandler();
    if (enableKeyboard) this.installKeyboardHandler();
    this.installPointerHandlers({ enableResize });
    this.installScrollbarHandlers();
    this.syncScrollbars();
  }

  destroy(): void {
    for (const dispose of this.disposeFns) dispose();
    this.disposeFns = [];
    this.stopAutoScroll();
    this.renderer.destroy();
    this.a11yStatusEl.remove();
  }

  setInteractionMode(mode: DesktopGridInteractionMode): void {
    this.interactionMode = mode;
    if (mode !== "default") {
      this.cancelFillHandleDrag();
    }
    this.renderer.setFillHandleEnabled(mode === "default" && Boolean(this.callbacks.onFillCommit));
  }

  /**
   * Cancel an in-progress fill handle drag (if any).
   *
   * Returns true when a fill drag was active and was canceled.
   */
  cancelFillHandleDrag(): boolean {
    if (this.dragMode !== "fillHandle") return false;

    this.dragMode = null;
    this.fillHandleState = null;
    this.lastPointerViewport = null;
    this.selectionAnchor = null;
    this.stopAutoScroll();

    const pointerId = this.selectionPointerId;
    this.selectionPointerId = null;

    this.renderer.setFillPreviewRange(null);
    this.selectionCanvas.style.cursor = "default";

    if (pointerId !== null) {
      try {
        this.selectionCanvas.releasePointerCapture?.(pointerId);
      } catch {
        // Ignore capture release failures.
      }
    }

    return true;
  }

  getScroll(): { x: number; y: number } {
    return this.renderer.scroll.getScroll();
  }

  scrollTo(x: number, y: number): void {
    this.renderer.setScroll(x, y);
    this.syncScrollbars();
    this.emitScroll();
  }

  scrollBy(dx: number, dy: number): void {
    this.renderer.scrollBy(dx, dy);
    this.syncScrollbars();
    this.emitScroll();
  }

  scrollToCell(row: number, col: number, opts?: { align?: ScrollToCellAlign; padding?: number }): void {
    this.renderer.scrollToCell(row, col, opts);
    this.syncScrollbars();
    this.emitScroll();
  }

  getCellRect(row: number, col: number): { x: number; y: number; width: number; height: number } | null {
    return this.renderer.getCellRect(row, col);
  }

  getPerfStats(): Readonly<GridPerfStats> {
    return this.renderer.getPerfStats();
  }

  setPerfStatsEnabled(enabled: boolean): void {
    this.renderer.setPerfStatsEnabled(enabled);
  }

  setSelectionRanges(ranges: CellRange[] | null, opts?: { activeIndex?: number; activeCell?: { row: number; col: number } | null }): void {
    this.keyboardAnchor = null;
    this.transientRange = null;
    this.renderer.setRangeSelection(null);

    const prevSelection = this.renderer.getSelection();
    const prevRange = this.renderer.getSelectionRange();

    this.renderer.setSelectionRanges(ranges, { activeIndex: opts?.activeIndex, activeCell: opts?.activeCell ?? undefined });

    const nextSelection = this.renderer.getSelection();
    const nextRange = this.renderer.getSelectionRange();

    this.announceSelection(nextSelection, nextRange);
    this.emitSelectionChange(prevSelection, nextSelection);
    this.emitSelectionRangeChange(prevRange, nextRange);

    if (nextSelection) {
      this.renderer.scrollToCell(nextSelection.row, nextSelection.col, { align: "auto", padding: 8 });
      this.syncScrollbars();
      this.emitScroll();
    }
  }

  setRangeSelection(range: CellRange | null): void {
    this.transientRange = range;
    this.renderer.setRangeSelection(range);
    this.announceSelection(this.renderer.getSelection(), range);
  }

  clearRangeSelection(): void {
    this.transientRange = null;
    this.renderer.setRangeSelection(null);
    this.announceSelection(this.renderer.getSelection(), this.renderer.getSelectionRange());
  }

  private announceSelection(selection: { row: number; col: number } | null, range: CellRange | null): void {
    if (
      (this.lastAnnounced.selection?.row ?? null) === (selection?.row ?? null) &&
      (this.lastAnnounced.selection?.col ?? null) === (selection?.col ?? null) &&
      rangesEqual(this.lastAnnounced.range, range)
    ) {
      return;
    }

    this.lastAnnounced = { selection, range };
    this.a11yStatusEl.textContent = describeCell(selection, range, this.provider, this.headerRows, this.headerCols);
  }

  private emitScroll(): void {
    this.callbacks.onScroll?.(this.renderer.scroll.getScroll(), this.renderer.getViewportState());
  }

  private emitSelectionChange(
    prev: { row: number; col: number } | null,
    next: { row: number; col: number } | null
  ): void {
    if ((prev?.row ?? null) === (next?.row ?? null) && (prev?.col ?? null) === (next?.col ?? null)) return;
    this.callbacks.onSelectionChange?.(next);
  }

  private emitSelectionRangeChange(prev: CellRange | null, next: CellRange | null): void {
    if (rangesEqual(prev, next)) return;
    this.callbacks.onSelectionRangeChange?.(next);
  }

  private installWheelHandler(): void {
    const onWheel = (event: WheelEvent) => {
      const target = event.target as HTMLElement | null;
      if (target?.closest?.('[data-testid="comments-panel"]')) return;
      if (event.ctrlKey) return;

      let deltaX = event.deltaX;
      let deltaY = event.deltaY;

      if (event.deltaMode === 1) {
        const line = 16;
        deltaX *= line;
        deltaY *= line;
      } else if (event.deltaMode === 2) {
        const viewport = this.renderer.scroll.getViewportState();
        deltaX *= viewport.width;
        deltaY *= viewport.height;
      }

      if (event.shiftKey && deltaX === 0) {
        deltaX = deltaY;
        deltaY = 0;
      }

      if (deltaX === 0 && deltaY === 0) return;

      event.preventDefault();
      this.renderer.scrollBy(deltaX, deltaY);
      this.syncScrollbars();
      this.emitScroll();
    };

    this.container.addEventListener("wheel", onWheel, { passive: false });
    this.disposeFns.push(() => this.container.removeEventListener("wheel", onWheel));
  }

  private installKeyboardHandler(): void {
    const onKeyDown = (event: KeyboardEvent) => {
      if (this.interactionMode === "rangeSelection") return;

      const renderer = this.renderer;
      const selection = renderer.getSelection();
      const { rowCount, colCount } = renderer.scroll.getCounts();
      if (rowCount === 0 || colCount === 0) return;

      const active =
        selection ??
        ({
          row: clampIndex(this.headerRows, 0, rowCount - 1),
          col: clampIndex(this.headerCols, 0, colCount - 1)
        } as const);

      const ctrlOrMeta = event.ctrlKey || event.metaKey;
      const dataStartRow = this.headerRows >= rowCount ? 0 : this.headerRows;
      const dataStartCol = this.headerCols >= colCount ? 0 : this.headerCols;

      const applySelectionRange = (range: CellRange) => {
        this.keyboardAnchor = null;

        const prevSelection = renderer.getSelection();
        const prevRange = renderer.getSelectionRange();

        renderer.setSelectionRange(range, { activeCell: prevSelection ?? active });

        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();

        this.announceSelection(nextSelection, nextRange);
        if (nextSelection) {
          renderer.scrollToCell(nextSelection.row, nextSelection.col, { align: "auto", padding: 8 });
          this.syncScrollbars();
          this.emitScroll();
        }

        this.emitSelectionChange(prevSelection, nextSelection);
        this.emitSelectionRangeChange(prevRange, nextRange);
      };

      if (event.key === "F2") {
        event.preventDefault();
        if (!selection) renderer.setSelection(active);
        this.callbacks.onRequestCellEdit?.({ row: active.row, col: active.col });
        return;
      }

      // Shift+Space selects the entire row, like Excel.
      if (!ctrlOrMeta && !event.altKey && event.shiftKey && (event.code === "Space" || event.key === " ")) {
        event.preventDefault();

        const startCol = this.headerCols >= colCount ? 0 : this.headerCols;
        applySelectionRange({
          startRow: active.row,
          endRow: active.row + 1,
          startCol,
          endCol: colCount
        });
        return;
      }

      // Ctrl/Cmd+Space selects the entire column, like Excel.
      if (ctrlOrMeta && !event.altKey && (event.code === "Space" || event.key === " ")) {
        event.preventDefault();

        const startRow = this.headerRows >= rowCount ? 0 : this.headerRows;
        applySelectionRange({
          startRow,
          endRow: rowCount,
          startCol: active.col,
          endCol: active.col + 1
        });
        return;
      }

      // Ctrl/Cmd+A selects all cells.
      if (ctrlOrMeta && !event.altKey && event.key.toLowerCase() === "a") {
        event.preventDefault();

        const startRow = this.headerRows >= rowCount ? 0 : this.headerRows;
        const startCol = this.headerCols >= colCount ? 0 : this.headerCols;
        applySelectionRange({
          startRow,
          endRow: rowCount,
          startCol,
          endCol: colCount
        });
        return;
      }

      const isPrintable = event.key.length === 1 && !ctrlOrMeta && !event.altKey;
      if (isPrintable) {
        event.preventDefault();
        if (!selection) renderer.setSelection(active);
        this.callbacks.onRequestCellEdit?.({ row: active.row, col: active.col, initialKey: event.key });
        return;
      }

      const prevSelection = renderer.getSelection();
      const prevRange = renderer.getSelectionRange();

      const rangeArea = (r: CellRange) => Math.max(0, r.endRow - r.startRow) * Math.max(0, r.endCol - r.startCol);

      // Excel-like behavior: Tab/Enter moves the active cell *within* the current selection range
      // (wrapping) instead of collapsing selection.
      if ((event.key === "Tab" || event.key === "Enter") && prevRange && rangeArea(prevRange) > 1) {
        event.preventDefault();
        this.keyboardAnchor = null;

        const current = prevSelection ?? { row: prevRange.startRow, col: prevRange.startCol };
        const activeRow = clamp(current.row, prevRange.startRow, prevRange.endRow - 1);
        const activeCol = clamp(current.col, prevRange.startCol, prevRange.endCol - 1);
        const backward = event.shiftKey;

        let nextRow = activeRow;
        let nextCol = activeCol;

        if (event.key === "Tab") {
          if (!backward) {
            if (activeCol + 1 < prevRange.endCol) {
              nextCol = activeCol + 1;
            } else if (activeRow + 1 < prevRange.endRow) {
              nextRow = activeRow + 1;
              nextCol = prevRange.startCol;
            } else {
              nextRow = prevRange.startRow;
              nextCol = prevRange.startCol;
            }
          } else {
            if (activeCol - 1 >= prevRange.startCol) {
              nextCol = activeCol - 1;
            } else if (activeRow - 1 >= prevRange.startRow) {
              nextRow = activeRow - 1;
              nextCol = prevRange.endCol - 1;
            } else {
              nextRow = prevRange.endRow - 1;
              nextCol = prevRange.endCol - 1;
            }
          }
        } else {
          if (!backward) {
            if (activeRow + 1 < prevRange.endRow) {
              nextRow = activeRow + 1;
            } else if (activeCol + 1 < prevRange.endCol) {
              nextRow = prevRange.startRow;
              nextCol = activeCol + 1;
            } else {
              nextRow = prevRange.startRow;
              nextCol = prevRange.startCol;
            }
          } else {
            if (activeRow - 1 >= prevRange.startRow) {
              nextRow = activeRow - 1;
            } else if (activeCol - 1 >= prevRange.startCol) {
              nextRow = prevRange.endRow - 1;
              nextCol = activeCol - 1;
            } else {
              nextRow = prevRange.endRow - 1;
              nextCol = prevRange.endCol - 1;
            }
          }
        }

        const ranges = renderer.getSelectionRanges();
        const activeIndex = renderer.getActiveSelectionIndex();
        renderer.setSelectionRanges(ranges, { activeIndex, activeCell: { row: nextRow, col: nextCol } });

        renderer.scrollToCell(nextRow, nextCol, { align: "auto", padding: 8 });
        this.syncScrollbars();
        this.emitScroll();

        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();

        this.announceSelection(nextSelection, nextRange);
        this.emitSelectionChange(prevSelection, nextSelection);
        return;
      }

      let nextRow = active.row;
      let nextCol = active.col;
      let handled = true;

      const viewport = renderer.scroll.getViewportState();
      const pageRows = Math.max(1, viewport.main.rows.end - viewport.main.rows.start);
      const pageCols = Math.max(1, viewport.main.cols.end - viewport.main.cols.start);

      switch (event.key) {
        case "ArrowUp":
          nextRow = ctrlOrMeta ? dataStartRow : active.row - 1;
          break;
        case "ArrowDown":
          nextRow = ctrlOrMeta ? rowCount - 1 : active.row + 1;
          break;
        case "ArrowLeft":
          nextCol = ctrlOrMeta ? dataStartCol : active.col - 1;
          break;
        case "ArrowRight":
          nextCol = ctrlOrMeta ? colCount - 1 : active.col + 1;
          break;
        case "PageUp":
          if (event.altKey) {
            nextCol = active.col - pageCols;
          } else {
            nextRow = active.row - pageRows;
          }
          break;
        case "PageDown":
          if (event.altKey) {
            nextCol = active.col + pageCols;
          } else {
            nextRow = active.row + pageRows;
          }
          break;
        case "Home":
          if (ctrlOrMeta) {
            nextRow = dataStartRow;
            nextCol = dataStartCol;
          } else {
            nextCol = dataStartCol;
          }
          break;
        case "End":
          if (ctrlOrMeta) {
            nextRow = rowCount - 1;
            nextCol = colCount - 1;
          } else {
            nextCol = colCount - 1;
          }
          break;
        case "Enter":
          nextRow = active.row + (event.shiftKey ? -1 : 1);
          break;
        case "Tab":
          nextCol = active.col + (event.shiftKey ? -1 : 1);
          break;
        default:
          handled = false;
      }

      if (!handled) return;
      event.preventDefault();

      nextRow = Math.max(dataStartRow, Math.min(rowCount - 1, nextRow));
      nextCol = Math.max(dataStartCol, Math.min(colCount - 1, nextCol));

      const extendSelection = event.shiftKey && event.key !== "Tab" && event.key !== "Enter";

      if (extendSelection) {
        const anchor = this.keyboardAnchor ?? prevSelection ?? active;
        if (!this.keyboardAnchor) this.keyboardAnchor = anchor;

        const range: CellRange = {
          startRow: Math.min(anchor.row, nextRow),
          endRow: Math.max(anchor.row, nextRow) + 1,
          startCol: Math.min(anchor.col, nextCol),
          endCol: Math.max(anchor.col, nextCol) + 1
        };

        const ranges = renderer.getSelectionRanges();
        const activeIndex = renderer.getActiveSelectionIndex();
        const updatedRanges = ranges.length === 0 ? [range] : ranges;
        updatedRanges[Math.min(activeIndex, updatedRanges.length - 1)] = range;
        renderer.setSelectionRanges(updatedRanges, { activeIndex, activeCell: { row: nextRow, col: nextCol } });
      } else {
        this.keyboardAnchor = null;
        renderer.setSelection({ row: nextRow, col: nextCol });
      }

      renderer.scrollToCell(nextRow, nextCol, { align: "auto", padding: 8 });
      this.syncScrollbars();
      this.emitScroll();

      const nextSelection = renderer.getSelection();
      const nextRange = renderer.getSelectionRange();

      this.announceSelection(nextSelection, nextRange);
      this.emitSelectionChange(prevSelection, nextSelection);
      this.emitSelectionRangeChange(prevRange, nextRange);
    };

    this.container.addEventListener("keydown", onKeyDown);
    this.disposeFns.push(() => this.container.removeEventListener("keydown", onKeyDown));
  }

  private stopAutoScroll(): void {
    if (this.autoScrollFrame == null) return;
    cancelAnimationFrame(this.autoScrollFrame);
    this.autoScrollFrame = null;
  }

  private scheduleAutoScroll(): void {
    if (this.autoScrollFrame != null) return;

    const clamp01 = (value: number) => Math.max(0, Math.min(1, value));
    const edge = 28;
    const maxSpeed = 24;

    const tick = () => {
      this.autoScrollFrame = null;
      if (this.selectionPointerId == null) return;
      const point = this.lastPointerViewport;
      if (!point) return;

      const renderer = this.renderer;
      const viewport = renderer.scroll.getViewportState();
      if (viewport.width <= 0 || viewport.height <= 0) return;

      const leftThreshold = viewport.frozenWidth + edge;
      const topThreshold = viewport.frozenHeight + edge;

      const leftFactor = clamp01((leftThreshold - point.x) / edge);
      const rightFactor = clamp01((point.x - (viewport.width - edge)) / edge);
      const topFactor = clamp01((topThreshold - point.y) / edge);
      const bottomFactor = clamp01((point.y - (viewport.height - edge)) / edge);

      const dx = viewport.maxScrollX > 0 ? (rightFactor - leftFactor) * maxSpeed : 0;
      const dy = viewport.maxScrollY > 0 ? (bottomFactor - topFactor) * maxSpeed : 0;

      if (dx === 0 && dy === 0) return;

      const before = renderer.scroll.getScroll();
      renderer.scrollBy(dx, dy);
      const after = renderer.scroll.getScroll();
      this.syncScrollbars();
      this.emitScroll();

      if (before.x === after.x && before.y === after.y) return;

      const clampedX = Math.max(0, Math.min(viewport.width, point.x));
      const clampedY = Math.max(0, Math.min(viewport.height, point.y));
      const picked = renderer.pickCellAt(clampedX, clampedY);
      if (picked) {
        if (this.dragMode === "fillHandle") this.applyFillHandleDrag(picked);
        else this.applyDragRange(picked);
      }

      this.autoScrollFrame = requestAnimationFrame(tick);
    };

    this.autoScrollFrame = requestAnimationFrame(tick);
  }

  private applyDragRange(picked: { row: number; col: number }): void {
    const renderer = this.renderer;
    const anchor = this.selectionAnchor;
    if (!anchor) return;

    const range: CellRange = {
      startRow: Math.min(anchor.row, picked.row),
      endRow: Math.max(anchor.row, picked.row) + 1,
      startCol: Math.min(anchor.col, picked.col),
      endCol: Math.max(anchor.col, picked.col) + 1
    };

    if (this.interactionMode === "rangeSelection") {
      const prev = this.transientRange;
      if (rangesEqual(prev, range)) return;
      this.transientRange = range;
      renderer.setRangeSelection(range);
      this.announceSelection(renderer.getSelection(), range);
      this.callbacks.onRangeSelectionChange?.(range);
      return;
    }

    const prevRange = renderer.getSelectionRange();
    if (rangesEqual(prevRange, range)) return;

    const ranges = renderer.getSelectionRanges();
    const activeIndex = renderer.getActiveSelectionIndex();
    const updatedRanges = ranges.length === 0 ? [range] : ranges;
    updatedRanges[Math.min(activeIndex, updatedRanges.length - 1)] = range;
    renderer.setSelectionRanges(updatedRanges, { activeIndex });

    const nextSelection = renderer.getSelection();
    const nextRange = renderer.getSelectionRange();
    this.announceSelection(nextSelection, nextRange);
    this.callbacks.onSelectionRangeChange?.(nextRange ?? range);
  }

  private computeFillDeltaRange(source: CellRange, union: CellRange): CellRange | null {
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
  }

  private computeFillTarget(source: CellRange, picked: { row: number; col: number }, direction: "up" | "down" | "left" | "right"): CellRange {
    const { rowCount, colCount } = this.renderer.scroll.getCounts();
    const dataStartRow = this.headerRows >= rowCount ? 0 : this.headerRows;
    const dataStartCol = this.headerCols >= colCount ? 0 : this.headerCols;

    const clampRow = (row: number) => Math.max(dataStartRow, Math.min(row, rowCount));
    const clampCol = (col: number) => Math.max(dataStartCol, Math.min(col, colCount));

    if (direction === "down") {
      const endRow = clampRow(Math.max(source.endRow, picked.row + 1));
      return { startRow: source.startRow, endRow, startCol: source.startCol, endCol: source.endCol };
    }

    if (direction === "up") {
      const startRow = clampRow(Math.min(source.startRow, picked.row));
      return { startRow, endRow: source.endRow, startCol: source.startCol, endCol: source.endCol };
    }

    if (direction === "right") {
      const endCol = clampCol(Math.max(source.endCol, picked.col + 1));
      return { startRow: source.startRow, endRow: source.endRow, startCol: source.startCol, endCol };
    }

    const startCol = clampCol(Math.min(source.startCol, picked.col));
    return { startRow: source.startRow, endRow: source.endRow, startCol, endCol: source.endCol };
  }

  private applyFillHandleDrag(picked: { row: number; col: number }): void {
    const state = this.fillHandleState;
    if (!state) return;

    const srcTop = state.source.startRow;
    const srcBottom = state.source.endRow - 1;
    const srcLeft = state.source.startCol;
    const srcRight = state.source.endCol - 1;

    const rowExtension = picked.row < srcTop ? picked.row - srcTop : picked.row > srcBottom ? picked.row - srcBottom : 0;
    const colExtension = picked.col < srcLeft ? picked.col - srcLeft : picked.col > srcRight ? picked.col - srcRight : 0;

    const unionRange = (() => {
      if (rowExtension === 0 && colExtension === 0) return state.source;

      const axis =
        rowExtension !== 0 && colExtension !== 0
          ? Math.abs(rowExtension) >= Math.abs(colExtension)
            ? "vertical"
            : "horizontal"
          : rowExtension !== 0
            ? "vertical"
            : "horizontal";

      const direction =
        axis === "vertical" ? (rowExtension > 0 ? "down" : "up") : colExtension > 0 ? "right" : "left";

      return this.computeFillTarget(state.source, picked, direction);
    })();

    const endCell = {
      row: clamp(picked.row, unionRange.startRow, unionRange.endRow - 1),
      col: clamp(picked.col, unionRange.startCol, unionRange.endCol - 1)
    };

    if (rangesEqual(unionRange, state.target) && endCell.row === state.endCell.row && endCell.col === state.endCell.col) {
      return;
    }

    const previewTarget = this.computeFillDeltaRange(state.source, unionRange);
    this.fillHandleState = { ...state, target: unionRange, previewTarget, endCell };
    this.renderer.setFillPreviewRange(unionRange);
  }

  private getViewportPoint(event: { clientX: number; clientY: number }): { x: number; y: number } {
    const rect = this.selectionCanvas.getBoundingClientRect();
    return { x: event.clientX - rect.left, y: event.clientY - rect.top };
  }

  private getResizeHit(viewportX: number, viewportY: number): ResizeHit | null {
    const renderer = this.renderer;
    const viewport = renderer.scroll.getViewportState();
    const { rowCount, colCount } = renderer.scroll.getCounts();
    if (rowCount === 0 || colCount === 0) return null;

    const frozenWidthClamped = Math.min(viewport.frozenWidth, viewport.width);
    const frozenHeightClamped = Math.min(viewport.frozenHeight, viewport.height);

    const inHeaderRow = viewport.frozenRows > 0 && viewportY >= 0 && viewportY <= frozenHeightClamped;
    const inRowHeaderCol = viewport.frozenCols > 0 && viewportX >= 0 && viewportX <= frozenWidthClamped;

    const absScrollX = viewport.frozenWidth + viewport.scrollX;
    const absScrollY = viewport.frozenHeight + viewport.scrollY;

    const colAxis = renderer.scroll.cols;
    const rowAxis = renderer.scroll.rows;

    const RESIZE_HIT_RADIUS_PX = 4;

    let best: (ResizeHit & { distance: number }) | null = null;

    if (inHeaderRow) {
      const inFrozenCols = viewportX < frozenWidthClamped;
      const sheetX = inFrozenCols ? viewportX : absScrollX + (viewportX - frozenWidthClamped);
      const minCol = inFrozenCols ? 0 : viewport.frozenCols;
      const maxColInclusive = inFrozenCols ? viewport.frozenCols - 1 : colCount - 1;

      if (maxColInclusive >= minCol) {
        const col = colAxis.indexAt(sheetX, { min: minCol, maxInclusive: maxColInclusive });
        const colStart = colAxis.positionOf(col);
        const colEnd = colStart + colAxis.getSize(col);

        const distToStart = Math.abs(sheetX - colStart);
        const distToEnd = Math.abs(sheetX - colEnd);

        if (distToStart <= RESIZE_HIT_RADIUS_PX && col > 0) {
          best = { kind: "col", index: col - 1, distance: distToStart };
        } else if (distToEnd <= RESIZE_HIT_RADIUS_PX) {
          best = { kind: "col", index: col, distance: distToEnd };
        }
      }
    }

    if (inRowHeaderCol) {
      const inFrozenRows = viewportY < frozenHeightClamped;
      const sheetY = inFrozenRows ? viewportY : absScrollY + (viewportY - frozenHeightClamped);
      const minRow = inFrozenRows ? 0 : viewport.frozenRows;
      const maxRowInclusive = inFrozenRows ? viewport.frozenRows - 1 : rowCount - 1;

      if (maxRowInclusive >= minRow) {
        const row = rowAxis.indexAt(sheetY, { min: minRow, maxInclusive: maxRowInclusive });
        const rowStart = rowAxis.positionOf(row);
        const rowEnd = rowStart + rowAxis.getSize(row);

        const distToStart = Math.abs(sheetY - rowStart);
        const distToEnd = Math.abs(sheetY - rowEnd);

        let candidate: (ResizeHit & { distance: number }) | null = null;
        if (distToStart <= RESIZE_HIT_RADIUS_PX && row > 0) {
          candidate = { kind: "row", index: row - 1, distance: distToStart };
        } else if (distToEnd <= RESIZE_HIT_RADIUS_PX) {
          candidate = { kind: "row", index: row, distance: distToEnd };
        }

        if (candidate && (!best || candidate.distance < best.distance)) {
          best = candidate;
        }
      }
    }

    return best ? { kind: best.kind, index: best.index } : null;
  }

  private installPointerHandlers(options: { enableResize: boolean }): void {
    const selectionCanvas = this.selectionCanvas;

    const MIN_COL_WIDTH = 24;
    const MIN_ROW_HEIGHT = 16;

    const onPointerDown = (event: PointerEvent) => {
      const renderer = this.renderer;
      const point = this.getViewportPoint(event);
      this.lastPointerViewport = point;

      // Excel/Sheets behavior: right-clicking inside an existing selection keeps the
      // selection intact; right-clicking outside moves the active cell to the clicked
      // cell. We intentionally only support sheet cells for now (not row/col header
      // context menus).
      if (event.pointerType === "mouse" && event.button !== 0) {
        const picked = renderer.pickCellAt(point.x, point.y);
        if (!picked) return;

        const isHeaderCell = (this.headerRows > 0 && picked.row < this.headerRows) || (this.headerCols > 0 && picked.col < this.headerCols);
        if (isHeaderCell) return;

        const prevSelection = renderer.getSelection();
        const prevRange = renderer.getSelectionRange();

        const ranges = renderer.getSelectionRanges();
        const inSelection = ranges.some(
          (range) =>
            picked.row >= range.startRow &&
            picked.row < range.endRow &&
            picked.col >= range.startCol &&
            picked.col < range.endCol
        );
        if (!inSelection) {
          renderer.setSelection(picked);
        }

        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();
        this.announceSelection(nextSelection, nextRange);
        this.emitSelectionChange(prevSelection, nextSelection);
        this.emitSelectionRangeChange(prevRange, nextRange);

        // Best-effort: keep focus on the grid so keyboard navigation continues.
        try {
          (this.container as any).focus?.({ preventScroll: true });
        } catch {
          this.container.focus?.();
        }
        return;
      }

      event.preventDefault();
      this.keyboardAnchor = null;
      this.dragMode = null;
      this.fillHandleState = null;
      renderer.setFillPreviewRange(null);

      if (options.enableResize) {
        const hit = this.getResizeHit(point.x, point.y);
        if (hit) {
          this.resizePointerId = event.pointerId;
          selectionCanvas.setPointerCapture?.(event.pointerId);

          if (hit.kind === "col") {
            this.resizeDrag = { kind: "col", index: hit.index, startClient: event.clientX, startSize: renderer.getColWidth(hit.index) };
            selectionCanvas.style.cursor = "col-resize";
          } else {
            this.resizeDrag = { kind: "row", index: hit.index, startClient: event.clientY, startSize: renderer.getRowHeight(hit.index) };
            selectionCanvas.style.cursor = "row-resize";
          }
          return;
        }
      }

      if (this.interactionMode === "default" && this.callbacks.onFillCommit) {
        const handleRect = renderer.getFillHandleRect();
        if (
          handleRect &&
          point.x >= handleRect.x &&
          point.x <= handleRect.x + handleRect.width &&
          point.y >= handleRect.y &&
          point.y <= handleRect.y + handleRect.height
        ) {
          const source = renderer.getSelectionRange();
          if (source) {
            this.selectionPointerId = event.pointerId;
            this.dragMode = "fillHandle";
            this.selectionAnchor = null;
            const mode: FillMode = event.altKey ? "formulas" : event.metaKey || event.ctrlKey ? "copy" : "series";
            this.fillHandleState = {
              source,
              target: source,
              mode,
              previewTarget: null,
              endCell: { row: source.endRow - 1, col: source.endCol - 1 }
            };
            selectionCanvas.setPointerCapture?.(event.pointerId);

            // Focus the container so keyboard shortcuts (e.g. Escape) still work while dragging.
            try {
              (this.container as any).focus?.({ preventScroll: true });
            } catch {
              this.container.focus?.();
            }

            this.transientRange = null;
            renderer.setRangeSelection(null);

            renderer.setFillPreviewRange(source);
            this.scheduleAutoScroll();
            return;
          }
        }
      }

      const picked = renderer.pickCellAt(point.x, point.y);
      if (!picked) return;

      if (this.interactionMode === "rangeSelection") {
        this.selectionPointerId = event.pointerId;
        this.dragMode = "selection";
        selectionCanvas.setPointerCapture?.(event.pointerId);

        this.selectionAnchor = picked;
        const range: CellRange = { startRow: picked.row, endRow: picked.row + 1, startCol: picked.col, endCol: picked.col + 1 };
        this.transientRange = range;
        renderer.setRangeSelection(range);
        this.announceSelection(renderer.getSelection(), range);
        this.callbacks.onRangeSelectionStart?.(range);
        this.scheduleAutoScroll();
        return;
      }

      // Focus the container so keyboard navigation continues after mouse interaction.
      try {
        (this.container as any).focus?.({ preventScroll: true });
      } catch {
        this.container.focus?.();
      }

      this.transientRange = null;
      renderer.setRangeSelection(null);

      const prevSelection = renderer.getSelection();
      const prevRange = renderer.getSelectionRange();

      const isAdditive = event.metaKey || event.ctrlKey;
      const isExtend = event.shiftKey;

      const { rowCount, colCount } = renderer.scroll.getCounts();
      const viewport = renderer.scroll.getViewportState();
      const dataStartRow = this.headerRows >= rowCount ? 0 : this.headerRows;
      const dataStartCol = this.headerCols >= colCount ? 0 : this.headerCols;

      const applyHeaderRange = (range: CellRange, activeCell: { row: number; col: number }) => {
        if (isAdditive) {
          const existing = renderer.getSelectionRanges();
          const nextRanges = [...existing, range];
          renderer.setSelectionRanges(nextRanges, { activeIndex: nextRanges.length - 1, activeCell });
          return;
        }

        if (isExtend && prevSelection) {
          const existing = renderer.getSelectionRanges();
          const activeIndex = renderer.getActiveSelectionIndex();
          const updatedRanges = existing.length === 0 ? [range] : existing;
          updatedRanges[Math.min(activeIndex, updatedRanges.length - 1)] = range;
          renderer.setSelectionRanges(updatedRanges, { activeIndex, activeCell });
          return;
        }

        renderer.setSelectionRange(range, { activeCell });
      };

      const isCornerHeader =
        this.headerRows > 0 && this.headerCols > 0 && picked.row < this.headerRows && picked.col < this.headerCols;
      const isColumnHeader = this.headerRows > 0 && picked.row < this.headerRows && picked.col >= this.headerCols;
      const isRowHeader = this.headerCols > 0 && picked.col < this.headerCols && picked.row >= this.headerRows;

      if (isCornerHeader || isColumnHeader || isRowHeader) {
        if (isCornerHeader) {
          const range: CellRange = { startRow: dataStartRow, endRow: rowCount, startCol: dataStartCol, endCol: colCount };
          const activeCell =
            prevSelection ??
            ({
              row: Math.max(dataStartRow, viewport.main.rows.start),
              col: Math.max(dataStartCol, viewport.main.cols.start)
            } as const);
          applyHeaderRange(range, activeCell);
        } else if (isColumnHeader) {
          const anchorCol = prevSelection ? clamp(prevSelection.col, dataStartCol, colCount - 1) : picked.col;
          const startCol = isExtend && prevSelection ? Math.min(anchorCol, picked.col) : picked.col;
          const endCol = (isExtend && prevSelection ? Math.max(anchorCol, picked.col) : picked.col) + 1;

          const range: CellRange = {
            startRow: dataStartRow,
            endRow: rowCount,
            startCol,
            endCol: Math.min(colCount, endCol)
          };

          const baseRow = prevSelection ? prevSelection.row : Math.max(dataStartRow, viewport.main.rows.start);
          applyHeaderRange(range, { row: baseRow, col: picked.col });
        } else {
          const anchorRow = prevSelection ? clamp(prevSelection.row, dataStartRow, rowCount - 1) : picked.row;
          const startRow = isExtend && prevSelection ? Math.min(anchorRow, picked.row) : picked.row;
          const endRow = (isExtend && prevSelection ? Math.max(anchorRow, picked.row) : picked.row) + 1;

          const range: CellRange = {
            startRow,
            endRow: Math.min(rowCount, endRow),
            startCol: dataStartCol,
            endCol: colCount
          };

          const baseCol = prevSelection ? prevSelection.col : Math.max(dataStartCol, viewport.main.cols.start);
          applyHeaderRange(range, { row: picked.row, col: baseCol });
        }

        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();

        this.announceSelection(nextSelection, nextRange);
        this.emitSelectionChange(prevSelection, nextSelection);
        this.emitSelectionRangeChange(prevRange, nextRange);
        this.scheduleAutoScroll();
        return;
      }

      this.selectionPointerId = event.pointerId;
      this.dragMode = "selection";
      selectionCanvas.setPointerCapture?.(event.pointerId);

      if (isAdditive) {
        this.selectionAnchor = picked;
        renderer.addSelectionRange({ startRow: picked.row, endRow: picked.row + 1, startCol: picked.col, endCol: picked.col + 1 });
      } else if (isExtend && prevSelection) {
        this.selectionAnchor = prevSelection;
        const range: CellRange = {
          startRow: Math.min(prevSelection.row, picked.row),
          endRow: Math.max(prevSelection.row, picked.row) + 1,
          startCol: Math.min(prevSelection.col, picked.col),
          endCol: Math.max(prevSelection.col, picked.col) + 1
        };
        const ranges = renderer.getSelectionRanges();
        const activeIndex = renderer.getActiveSelectionIndex();
        const updatedRanges = ranges.length === 0 ? [range] : ranges;
        updatedRanges[Math.min(activeIndex, updatedRanges.length - 1)] = range;
        renderer.setSelectionRanges(updatedRanges, { activeIndex });
      } else {
        this.selectionAnchor = picked;
        renderer.setSelection(picked);
      }

      const nextSelection = renderer.getSelection();
      const nextRange = renderer.getSelectionRange();

      this.announceSelection(nextSelection, nextRange);
      this.emitSelectionChange(prevSelection, nextSelection);
      this.emitSelectionRangeChange(prevRange, nextRange);

      this.scheduleAutoScroll();
    };

    const onPointerMove = (event: PointerEvent) => {
      const renderer = this.renderer;

      if (this.resizePointerId != null) {
        if (event.pointerId !== this.resizePointerId) return;
        const drag = this.resizeDrag;
        if (!drag) return;
        event.preventDefault();

        if (drag.kind === "col") {
          const delta = event.clientX - drag.startClient;
          renderer.setColWidth(drag.index, Math.max(MIN_COL_WIDTH, drag.startSize + delta));
        } else {
          const delta = event.clientY - drag.startClient;
          renderer.setRowHeight(drag.index, Math.max(MIN_ROW_HEIGHT, drag.startSize + delta));
        }

        this.syncScrollbars();
        this.emitScroll();
        return;
      }

      if (this.selectionPointerId == null) return;
      if (event.pointerId !== this.selectionPointerId) return;
      event.preventDefault();

      const point = this.getViewportPoint(event);
      this.lastPointerViewport = point;

      const picked = renderer.pickCellAt(point.x, point.y);
      if (!picked) return;
      if (this.dragMode === "fillHandle") this.applyFillHandleDrag(picked);
      else this.applyDragRange(picked);
      this.scheduleAutoScroll();
    };

    const endDrag = (event: PointerEvent) => {
      const renderer = this.renderer;
      if (this.resizePointerId != null && event.pointerId === this.resizePointerId) {
        const drag = this.resizeDrag;
        this.resizePointerId = null;
        this.resizeDrag = null;
        selectionCanvas.style.cursor = "default";
        try {
          selectionCanvas.releasePointerCapture?.(event.pointerId);
        } catch {
          // Ignore.
        }

        if (drag) {
          const endSize = drag.kind === "col" ? this.renderer.getColWidth(drag.index) : this.renderer.getRowHeight(drag.index);
          const defaultSize = drag.kind === "col" ? this.renderer.scroll.cols.defaultSize : this.renderer.scroll.rows.defaultSize;
          if (endSize !== drag.startSize) {
            this.callbacks.onAxisSizeChange?.({
              kind: drag.kind,
              index: drag.index,
              size: endSize,
              previousSize: drag.startSize,
              defaultSize,
              zoom: this.renderer.getZoom(),
              source: "resize"
            });
          }
        }
      }

      if (this.selectionPointerId == null) return;
      if (event.pointerId !== this.selectionPointerId) return;

      const prevSelection = renderer.getSelection();
      const prevRange = renderer.getSelectionRange();
      const dragMode = this.dragMode;
      this.dragMode = null;

      this.selectionPointerId = null;
      this.selectionAnchor = null;
      this.lastPointerViewport = null;
      this.stopAutoScroll();

      if (dragMode === "fillHandle") {
        const state = this.fillHandleState;
        this.fillHandleState = null;
        renderer.setFillPreviewRange(null);

        const shouldCommit = event.type === "pointerup";
        if (state && shouldCommit && !rangesEqual(state.source, state.target)) {
          const targetRange = this.computeFillDeltaRange(state.source, state.target);
          if (targetRange) {
            const commitResult = this.callbacks.onFillCommit?.({
              sourceRange: state.source,
              targetRange,
              mode: state.mode
            });
            void Promise.resolve(commitResult).catch(() => {
              // Consumers own commit error handling; swallow to avoid unhandled rejections.
            });

            const ranges = renderer.getSelectionRanges();
            const activeIndex = renderer.getActiveSelectionIndex();
            const updatedRanges = ranges.length === 0 ? [state.target] : [...ranges];
            updatedRanges[Math.min(activeIndex, updatedRanges.length - 1)] = state.target;
            renderer.setSelectionRanges(updatedRanges, { activeIndex, activeCell: state.endCell });
          }
        }

        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();
        this.announceSelection(nextSelection, nextRange);
        this.emitSelectionChange(prevSelection, nextSelection);
        this.emitSelectionRangeChange(prevRange, nextRange);
      } else {
        this.fillHandleState = null;
        renderer.setFillPreviewRange(null);

        if (this.interactionMode === "rangeSelection") {
          const range = this.transientRange;
          if (range) this.callbacks.onRangeSelectionEnd?.(range);
        }
      }

      selectionCanvas.style.cursor = "default";

      try {
        selectionCanvas.releasePointerCapture?.(event.pointerId);
      } catch {
        // Ignore.
      }
    };

    const onPointerHover = (event: PointerEvent) => {
      if (this.resizePointerId != null || this.selectionPointerId != null) return;
      const point = this.getViewportPoint(event);

      if (options.enableResize) {
        const hit = this.getResizeHit(point.x, point.y);
        if (hit) {
          selectionCanvas.style.cursor = hit.kind === "col" ? "col-resize" : "row-resize";
          return;
        }
      }

      if (this.interactionMode === "default" && this.callbacks.onFillCommit) {
        const handleRect = this.renderer.getFillHandleRect();
        if (
          handleRect &&
          point.x >= handleRect.x &&
          point.x <= handleRect.x + handleRect.width &&
          point.y >= handleRect.y &&
          point.y <= handleRect.y + handleRect.height
        ) {
          selectionCanvas.style.cursor = "crosshair";
          return;
        }
      }

      selectionCanvas.style.cursor = "default";
    };

    const onPointerLeave = () => {
      if (this.resizePointerId != null || this.selectionPointerId != null) return;
      selectionCanvas.style.cursor = "default";
    };

    const onDoubleClick = (event: MouseEvent) => {
      if (this.interactionMode === "rangeSelection") return;
      const renderer = this.renderer;
      const point = this.getViewportPoint(event);

      if (options.enableResize) {
        const hit = this.getResizeHit(point.x, point.y);
        if (hit) {
          event.preventDefault();
          const prevSize = hit.kind === "col" ? renderer.getColWidth(hit.index) : renderer.getRowHeight(hit.index);
          const defaultSize = hit.kind === "col" ? renderer.scroll.cols.defaultSize : renderer.scroll.rows.defaultSize;
          const nextSize =
            hit.kind === "col" ? renderer.autoFitCol(hit.index, { maxWidth: 500 }) : renderer.autoFitRow(hit.index, { maxHeight: 500 });
          this.syncScrollbars();
          this.emitScroll();

          if (nextSize !== prevSize) {
            this.callbacks.onAxisSizeChange?.({
              kind: hit.kind,
              index: hit.index,
              size: nextSize,
              previousSize: prevSize,
              defaultSize,
              zoom: renderer.getZoom(),
              source: "autoFit"
            });
          }
          return;
        }
      }

      const picked = renderer.pickCellAt(point.x, point.y);
      if (!picked) return;
      this.callbacks.onRequestCellEdit?.({ row: picked.row, col: picked.col });
    };

    selectionCanvas.addEventListener("pointerdown", onPointerDown);
    selectionCanvas.addEventListener("pointermove", onPointerMove);
    selectionCanvas.addEventListener("pointermove", onPointerHover);
    selectionCanvas.addEventListener("pointerleave", onPointerLeave);
    selectionCanvas.addEventListener("pointerup", endDrag);
    selectionCanvas.addEventListener("pointercancel", endDrag);
    selectionCanvas.addEventListener("dblclick", onDoubleClick);

    this.disposeFns.push(() => {
      selectionCanvas.removeEventListener("pointerdown", onPointerDown);
      selectionCanvas.removeEventListener("pointermove", onPointerMove);
      selectionCanvas.removeEventListener("pointermove", onPointerHover);
      selectionCanvas.removeEventListener("pointerleave", onPointerLeave);
      selectionCanvas.removeEventListener("pointerup", endDrag);
      selectionCanvas.removeEventListener("pointercancel", endDrag);
      selectionCanvas.removeEventListener("dblclick", onDoubleClick);
    });
  }

  private installScrollbarHandlers(): void {
    const onVThumbDown = (event: PointerEvent) => {
      event.preventDefault();
      event.stopPropagation();

      const renderer = this.renderer;
      const trackRect = this.vTrack.getBoundingClientRect();
      const thumbRect = this.vThumb.getBoundingClientRect();
      const maxScroll = renderer.scroll.getViewportState().maxScrollY;
      if (maxScroll === 0) return;

      const grabOffset = event.clientY - thumbRect.top;
      const thumbTravel = Math.max(0, trackRect.height - thumbRect.height);

      const onMove = (move: PointerEvent) => {
        if (move.pointerId !== event.pointerId) return;
        move.preventDefault();
        const pointerPos = move.clientY;
        const thumbOffset = pointerPos - trackRect.top - grabOffset;
        const clamped = clamp(thumbOffset, 0, thumbTravel);
        const nextScroll = thumbTravel === 0 ? 0 : (clamped / thumbTravel) * maxScroll;
        renderer.setScroll(renderer.scroll.getScroll().x, nextScroll);
        this.syncScrollbars();
        this.emitScroll();
      };

      const onUp = (up: PointerEvent) => {
        if (up.pointerId !== event.pointerId) return;
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
      };

      window.addEventListener("pointermove", onMove, { passive: false });
      window.addEventListener("pointerup", onUp, { passive: false });
    };

    this.vThumb.addEventListener("pointerdown", onVThumbDown, { passive: false });
    this.disposeFns.push(() => this.vThumb.removeEventListener("pointerdown", onVThumbDown));

    const onHThumbDown = (event: PointerEvent) => {
      event.preventDefault();
      event.stopPropagation();

      const renderer = this.renderer;
      const trackRect = this.hTrack.getBoundingClientRect();
      const thumbRect = this.hThumb.getBoundingClientRect();
      const maxScroll = renderer.scroll.getViewportState().maxScrollX;
      if (maxScroll === 0) return;

      const grabOffset = event.clientX - thumbRect.left;
      const thumbTravel = Math.max(0, trackRect.width - thumbRect.width);

      const onMove = (move: PointerEvent) => {
        if (move.pointerId !== event.pointerId) return;
        move.preventDefault();
        const pointerPos = move.clientX;
        const thumbOffset = pointerPos - trackRect.left - grabOffset;
        const clamped = clamp(thumbOffset, 0, thumbTravel);
        const nextScroll = thumbTravel === 0 ? 0 : (clamped / thumbTravel) * maxScroll;
        renderer.setScroll(nextScroll, renderer.scroll.getScroll().y);
        this.syncScrollbars();
        this.emitScroll();
      };

      const onUp = (up: PointerEvent) => {
        if (up.pointerId !== event.pointerId) return;
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
      };

      window.addEventListener("pointermove", onMove, { passive: false });
      window.addEventListener("pointerup", onUp, { passive: false });
    };

    this.hThumb.addEventListener("pointerdown", onHThumbDown, { passive: false });
    this.disposeFns.push(() => this.hThumb.removeEventListener("pointerdown", onHThumbDown));

    const onVTrackDown = (event: PointerEvent) => {
      if (event.target !== this.vTrack) return;
      const renderer = this.renderer;
      event.preventDefault();
      event.stopPropagation();

      const viewport = renderer.scroll.getViewportState();
      const maxScrollY = viewport.maxScrollY;
      const trackRect = this.vTrack.getBoundingClientRect();

      const thumb = computeScrollbarThumb({
        scrollPos: renderer.scroll.getScroll().y,
        viewportSize: Math.max(0, viewport.height - viewport.frozenHeight),
        contentSize: Math.max(0, viewport.totalHeight - viewport.frozenHeight),
        trackSize: trackRect.height
      });

      const thumbTravel = Math.max(0, trackRect.height - thumb.size);
      if (thumbTravel === 0 || maxScrollY === 0) return;

      const pointerPos = event.clientY - trackRect.top;
      const targetOffset = pointerPos - thumb.size / 2;
      const clampedOffset = clamp(targetOffset, 0, thumbTravel);
      const nextScroll = (clampedOffset / thumbTravel) * maxScrollY;

      renderer.setScroll(renderer.scroll.getScroll().x, nextScroll);
      this.syncScrollbars();
      this.emitScroll();
    };

    this.vTrack.addEventListener("pointerdown", onVTrackDown, { passive: false });
    this.disposeFns.push(() => this.vTrack.removeEventListener("pointerdown", onVTrackDown));

    const onHTrackDown = (event: PointerEvent) => {
      if (event.target !== this.hTrack) return;
      const renderer = this.renderer;
      event.preventDefault();
      event.stopPropagation();

      const viewport = renderer.scroll.getViewportState();
      const maxScrollX = viewport.maxScrollX;
      const trackRect = this.hTrack.getBoundingClientRect();

      const thumb = computeScrollbarThumb({
        scrollPos: renderer.scroll.getScroll().x,
        viewportSize: Math.max(0, viewport.width - viewport.frozenWidth),
        contentSize: Math.max(0, viewport.totalWidth - viewport.frozenWidth),
        trackSize: trackRect.width
      });

      const thumbTravel = Math.max(0, trackRect.width - thumb.size);
      if (thumbTravel === 0 || maxScrollX === 0) return;

      const pointerPos = event.clientX - trackRect.left;
      const targetOffset = pointerPos - thumb.size / 2;
      const clampedOffset = clamp(targetOffset, 0, thumbTravel);
      const nextScroll = (clampedOffset / thumbTravel) * maxScrollX;

      renderer.setScroll(nextScroll, renderer.scroll.getScroll().y);
      this.syncScrollbars();
      this.emitScroll();
    };

    this.hTrack.addEventListener("pointerdown", onHTrackDown, { passive: false });
    this.disposeFns.push(() => this.hTrack.removeEventListener("pointerdown", onHTrackDown));
  }

  syncScrollbars(): void {
    const renderer = this.renderer;
    const viewport = renderer.scroll.getViewportState();
    const scroll = renderer.scroll.getScroll();

    const maxX = viewport.maxScrollX;
    const maxY = viewport.maxScrollY;
    const showH = maxX > 0;
    const showV = maxY > 0;

    const padding = 2;
    const thickness = 10;

    this.vTrack.style.display = showV ? "block" : "none";
    this.hTrack.style.display = showH ? "block" : "none";

    const frozenWidth = Math.min(viewport.frozenWidth, viewport.width);
    const frozenHeight = Math.min(viewport.frozenHeight, viewport.height);

    if (showV) {
      this.vTrack.style.right = `${padding}px`;
      this.vTrack.style.top = `${frozenHeight + padding}px`;
      this.vTrack.style.bottom = `${(showH ? thickness : 0) + padding}px`;
      this.vTrack.style.width = `${thickness}px`;

      const trackSize = this.vTrack.getBoundingClientRect().height;
      const thumb = computeScrollbarThumb({
        scrollPos: scroll.y,
        viewportSize: Math.max(0, viewport.height - frozenHeight),
        contentSize: Math.max(0, viewport.totalHeight - frozenHeight),
        trackSize
      });

      this.vThumb.style.height = `${thumb.size}px`;
      this.vThumb.style.transform = `translateY(${thumb.offset}px)`;
    }

    if (showH) {
      this.hTrack.style.left = `${frozenWidth + padding}px`;
      this.hTrack.style.right = `${(showV ? thickness : 0) + padding}px`;
      this.hTrack.style.bottom = `${padding}px`;
      this.hTrack.style.height = `${thickness}px`;

      const trackSize = this.hTrack.getBoundingClientRect().width;
      const thumb = computeScrollbarThumb({
        scrollPos: scroll.x,
        viewportSize: Math.max(0, viewport.width - frozenWidth),
        contentSize: Math.max(0, viewport.totalWidth - frozenWidth),
        trackSize
      });

      this.hThumb.style.width = `${thumb.size}px`;
      this.hThumb.style.transform = `translateX(${thumb.offset}px)`;
    }
  }

  resize(width: number, height: number, devicePixelRatio: number): void {
    this.renderer.resize(width, height, devicePixelRatio);
    this.syncScrollbars();
    this.emitScroll();
  }
}
