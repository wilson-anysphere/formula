import React, { useCallback, useEffect, useId, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState } from "react";
import type { CellProvider, CellRange } from "../model/CellProvider";
import type { GridPresence } from "../presence/types";
import {
  CanvasGridRenderer,
  type CanvasGridImageResolver,
  type GridPerfStats
} from "../rendering/CanvasGridRenderer";
import type { GridTheme } from "../theme/GridTheme";
import { resolveGridTheme } from "../theme/GridTheme";
import { resolveGridThemeFromCssVars } from "../theme/resolveThemeFromCssVars";
import { computeFillPreview, hitTestSelectionHandle, type FillDragPreview, type FillMode } from "../interaction/fillHandle";
import { clampZoom } from "../utils/zoomMath";
import { computeScrollbarThumb } from "../virtualization/scrollbarMath";
import type { GridViewportState } from "../virtualization/VirtualScrollManager";
import { wheelDeltaToPixels } from "./wheelDeltaToPixels";
import { describeActiveCellLabel, describeCellForA11y, SR_ONLY_STYLE } from "../a11y/a11y";

export type ScrollToCellAlign = "auto" | "start" | "center" | "end";

export interface GridApi {
  scrollTo(x: number, y: number): void;
  scrollBy(deltaX: number, deltaY: number): void;
  getScroll(): { x: number; y: number };
  setZoom(zoom: number): void;
  getZoom(): number;
  /**
   * Invalidate a single decoded image in the renderer cache.
   *
   * The next paint for any visible cells referencing the image will re-resolve and decode it.
   */
  invalidateImage(imageId: string): void;
  /** Clear all decoded images from the renderer cache. */
  clearImageCache(): void;
  setFrozen(frozenRows: number, frozenCols: number): void;
  setRowHeight(row: number, height: number): void;
  setColWidth(col: number, width: number): void;
  resetRowHeight(row: number): void;
  resetColWidth(col: number): void;
  /**
   * Apply many row/column size overrides in one batch, triggering at most one full invalidation.
   *
   * Sizes are specified in CSS pixels at the current zoom.
   *
   * When `resetUnspecified` is set, any existing overrides for the provided axes that are *not*
   * present in the new maps are cleared.
   */
  applyAxisSizeOverrides(
    overrides: { rows?: ReadonlyMap<number, number>; cols?: ReadonlyMap<number, number> },
    options?: { resetUnspecified?: boolean }
  ): void;
  getRowHeight(row: number): number;
  getColWidth(col: number): number;
  setSelection(row: number, col: number): void;
  setSelectionRange(range: CellRange | null): void;
  getSelectionRange(): CellRange | null;
  setSelectionRanges(ranges: CellRange[] | null, opts?: { activeIndex?: number }): void;
  getSelectionRanges(): CellRange[];
  getActiveSelectionRangeIndex(): number;
  clearSelection(): void;
  getSelection(): { row: number; col: number } | null;
  getPerfStats(): Readonly<GridPerfStats> | null;
  setPerfStatsEnabled(enabled: boolean): void;
  scrollToCell(row: number, col: number, opts?: { align?: ScrollToCellAlign; padding?: number }): void;
  getCellRect(row: number, col: number): { x: number; y: number; width: number; height: number } | null;
  /**
   * Returns the fill-handle rect for the active selection range, in viewport coordinates
   * (clipped to the visible viewport).
   *
   * Returns `null` when the handle is not visible, e.g. offscreen, behind frozen
   * rows/cols, or when `interactionMode !== "default"`.
   */
  getFillHandleRect(): { x: number; y: number; width: number; height: number } | null;
  getViewportState(): GridViewportState | null;
  /**
   * Set a transient range selection overlay.
   *
   * This does not affect the primary grid selection; it's intended for
   * formula-bar range picking UX.
   */
  setRangeSelection(range: CellRange | null): void;
  /**
   * Set the list of ranges that should be highlighted as formula references.
   *
   * This is intended to match Excel's UX while editing a formula: each referenced
   * range is outlined on the grid using the same color as the formula text.
   */
  setReferenceHighlights(highlights: Array<{ range: CellRange; color: string; active?: boolean }> | null): void;
  setRemotePresences(presences: GridPresence[] | null): void;
  renderImmediately(): void;
}

export type GridInteractionMode = "default" | "rangeSelection";

export interface FillCommitEvent {
  sourceRange: CellRange;
  targetRange: CellRange;
  mode: FillMode;
}

export interface CanvasGridProps {
  provider: CellProvider;
  rowCount: number;
  colCount: number;
  /**
   * Optional resolver for `CellData.image` payloads.
   *
   * When provided, the {@link CanvasGridRenderer} can draw in-cell images directly on the
   * canvas content layer (no DOM overlays).
   */
  imageResolver?: CanvasGridImageResolver | null;
  /**
   * Number of header rows at the top of the grid.
   *
   * Header rows are treated as non-data UI (e.g. column labels). Selection ranges
   * produced by header interactions exclude the header region and instead target
   * the data region `[headerRows, rowCount)`.
   *
   * The default selection overlay also follows the selection range, so header
   * cells are not highlighted when selecting a full row/column via headers.
   *
   * Note: header rows are typically also frozen via `frozenRows` for best UX,
   * but this is not required.
   */
  headerRows?: number;
  /**
   * Number of header columns at the left side of the grid.
   *
   * Header columns are treated as non-data UI (e.g. row labels). Selection ranges
   * produced by header interactions exclude the header region and instead target
   * the data region `[headerCols, colCount)`.
   *
   * The default selection overlay also follows the selection range, so header
   * cells are not highlighted when selecting a full row/column via headers.
   *
   * Note: header columns are typically also frozen via `frozenCols` for best UX,
   * but this is not required.
   */
  headerCols?: number;
  frozenRows?: number;
  frozenCols?: number;
  theme?: Partial<GridTheme>;
  /**
   * Default font family used for cell text when `CellStyle.fontFamily` is unset.
   *
   * When omitted, the renderer uses its default system UI font stack.
   */
  defaultCellFontFamily?: string;
  /**
   * Default font family used for header cell text when `CellStyle.fontFamily` is unset.
   *
   * Defaults to `defaultCellFontFamily` when omitted.
   */
  defaultHeaderFontFamily?: string;
  defaultRowHeight?: number;
  defaultColWidth?: number;
  zoom?: number;
  onZoomChange?: (zoom: number) => void;
  /**
   * Fired whenever the grid's effective scroll position changes.
   *
   * Useful for positioning DOM overlays (tooltips, editors, etc.) since the grid
   * does not use a native scroll container.
   */
  onScroll?: (scroll: { x: number; y: number }, viewport: GridViewportState) => void;
  /**
   * Touch interaction strategy.
   *
   * - `"auto"` (default): single-finger drags pan; taps select.
   * - `"select"`: touch behaves like mouse (drag selects).
   * - `"pan"`: touch always pans (taps do not select).
   */
  touchMode?: "pan" | "select" | "auto";
  enableResize?: boolean;
  /**
   * Fired when a row height or column width is updated via an interactive resize
   * drag or via auto-fit (double-click / double-tap on a resize handle).
   *
   * This callback intentionally only fires for user-driven interactions so host
   * applications can persist the new size and create a single undo entry.
   */
  onAxisSizeChange?: (change: GridAxisSizeChange) => void;
  /**
   * How many extra rows beyond the visible viewport to prefetch.
   *
   * This reduces flicker/blank cells when using async (engine-backed) providers
   * by warming the cache ahead of fast scrolls.
   */
  prefetchOverscanRows?: number;
  /**
   * How many extra columns beyond the visible viewport to prefetch.
   *
   * This reduces flicker/blank cells when using async (engine-backed) providers
   * by warming the cache ahead of fast scrolls.
   */
  prefetchOverscanCols?: number;
  remotePresences?: GridPresence[] | null;
  apiRef?: React.Ref<GridApi>;
  onSelectionChange?: (cell: { row: number; col: number } | null) => void;
  onSelectionRangeChange?: (range: CellRange | null) => void;
  interactionMode?: GridInteractionMode;
  onRangeSelectionStart?: (range: CellRange) => void;
  onRangeSelectionChange?: (range: CellRange) => void;
  onRangeSelectionEnd?: (range: CellRange) => void;
  onFillHandleChange?: (args: { source: CellRange; target: CellRange }) => void;
  /**
   * Called when the user finishes dragging the selection fill handle.
   *
   * `target` is the full selection range after the drag, including the original
   * `source` range.
   */
  onFillHandleCommit?: (args: { source: CellRange; target: CellRange }) => void | Promise<void>;
  onFillPreviewChange?: (previewRange: CellRange | null) => void;
  onFillCommit?: (event: FillCommitEvent) => void | Promise<void>;
  onRequestCellEdit?: (request: { row: number; col: number; initialKey?: string }) => void;
  style?: React.CSSProperties;
  ariaLabel?: string;
  ariaLabelledBy?: string;
}

export type GridAxisSizeChange =
  | {
      kind: "col";
      index: number;
      /** New size in CSS pixels. */
      size: number;
      /** Size before the interaction began. */
      previousSize: number;
      /** Default size for this axis (used when the override is cleared). */
      defaultSize: number;
      /** Current grid zoom factor. */
      zoom: number;
      source: "resize" | "autoFit";
    }
  | {
      kind: "row";
      index: number;
      /** New size in CSS pixels. */
      size: number;
      /** Size before the interaction began. */
      previousSize: number;
      /** Default size for this axis (used when the override is cleared). */
      defaultSize: number;
      /** Current grid zoom factor. */
      zoom: number;
      source: "resize" | "autoFit";
    };

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function clampIndex(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return min;
  return clamp(Math.trunc(value), min, max);
}

function partialThemeEqual(a: Partial<GridTheme>, b: Partial<GridTheme>): boolean {
  const aKeys = Object.keys(a) as Array<keyof GridTheme>;
  const bKeys = Object.keys(b) as Array<keyof GridTheme>;
  if (aKeys.length !== bKeys.length) return false;
  for (const key of aKeys) {
    if (a[key] !== b[key]) return false;
  }
  return true;
}

type ResizeHit = { kind: "col"; index: number } | { kind: "row"; index: number };

type ResizeDragState =
  | { kind: "col"; index: number; startClient: number; startSize: number }
  | { kind: "row"; index: number; startClient: number; startSize: number };

export function CanvasGrid(props: CanvasGridProps): React.ReactElement {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const gridCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const contentCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const selectionCanvasRef = useRef<HTMLCanvasElement | null>(null);

  const vTrackRef = useRef<HTMLDivElement | null>(null);
  const vThumbRef = useRef<HTMLDivElement | null>(null);
  const hTrackRef = useRef<HTMLDivElement | null>(null);
  const hThumbRef = useRef<HTMLDivElement | null>(null);

  const scrollbarLayoutRef = useRef<{ inset: number; corner: number }>({ inset: 0, corner: 0 });
  const scrollbarThumbRef = useRef<{ vSize: number | null; vOffset: number | null; hSize: number | null; hOffset: number | null }>({
    vSize: null,
    vOffset: null,
    hSize: null,
    hOffset: null
  });
  const scrollbarThumbScratchRef = useRef({
    v: { size: 0, offset: 0 },
    h: { size: 0, offset: 0 }
  });

  const rendererRef = useRef<CanvasGridRenderer | null>(null);
  const onSelectionChangeRef = useRef(props.onSelectionChange);
  const onSelectionRangeChangeRef = useRef(props.onSelectionRangeChange);
  const onRangeSelectionStartRef = useRef(props.onRangeSelectionStart);
  const onRangeSelectionChangeRef = useRef(props.onRangeSelectionChange);
  const onRangeSelectionEndRef = useRef(props.onRangeSelectionEnd);
  const onRequestCellEditRef = useRef(props.onRequestCellEdit);
  const onFillHandleChangeRef = useRef(props.onFillHandleChange);
  const onFillHandleCommitRef = useRef(props.onFillHandleCommit);
  const onFillPreviewChangeRef = useRef(props.onFillPreviewChange);
  const onFillCommitRef = useRef(props.onFillCommit);
  const onZoomChangeRef = useRef(props.onZoomChange);
  const onAxisSizeChangeRef = useRef(props.onAxisSizeChange);
  const onScrollRef = useRef(props.onScroll);
  const touchModeRef = useRef<CanvasGridProps["touchMode"]>(props.touchMode ?? "auto");

  onSelectionChangeRef.current = props.onSelectionChange;
  onSelectionRangeChangeRef.current = props.onSelectionRangeChange;
  onRangeSelectionStartRef.current = props.onRangeSelectionStart;
  onRangeSelectionChangeRef.current = props.onRangeSelectionChange;
  onRangeSelectionEndRef.current = props.onRangeSelectionEnd;
  onRequestCellEditRef.current = props.onRequestCellEdit;
  onFillHandleChangeRef.current = props.onFillHandleChange;
  onFillHandleCommitRef.current = props.onFillHandleCommit;
  onFillPreviewChangeRef.current = props.onFillPreviewChange;
  onFillCommitRef.current = props.onFillCommit;
  onZoomChangeRef.current = props.onZoomChange;
  onAxisSizeChangeRef.current = props.onAxisSizeChange;
  onScrollRef.current = props.onScroll;
  touchModeRef.current = props.touchMode ?? "auto";

  const selectionAnchorRef = useRef<{ row: number; col: number } | null>(null);
  const keyboardAnchorRef = useRef<{ row: number; col: number } | null>(null);
  const selectionPointerIdRef = useRef<number | null>(null);
  const selectionCanvasViewportOriginRef = useRef<{ left: number; top: number } | null>(null);
  const transientRangeRef = useRef<CellRange | null>(null);
  const lastPointerViewportRef = useRef<{ x: number; y: number } | null>(null);
  const autoScrollFrameRef = useRef<number | null>(null);
  const resizePointerIdRef = useRef<number | null>(null);
  const resizeDragRef = useRef<ResizeDragState | null>(null);
  const lastResizeClickRef = useRef<{ time: number; hit: ResizeHit } | null>(null);
  const dragModeRef = useRef<"selection" | "fillHandle" | null>(null);
  const cancelFillHandleDragRef = useRef<(() => void) | null>(null);
  const fillHandleStateRef = useRef<{
    source: CellRange;
    target: CellRange;
    mode: FillMode;
    previewTarget: CellRange;
  } | null>(null);
  const lastEmittedScrollRef = useRef<{ x: number; y: number } | null>(null);

  const sanitizeHeaderCount = (value: number | undefined, max: number) => {
    if (typeof value !== "number" || !Number.isFinite(value)) return 0;
    return Math.max(0, Math.min(max, Math.floor(value)));
  };

  const headersControlled = props.headerRows !== undefined || props.headerCols !== undefined;
  const headersControlledRef = useRef(headersControlled);
  headersControlledRef.current = headersControlled;
  const headerRows = sanitizeHeaderCount(props.headerRows, props.rowCount);
  const headerCols = sanitizeHeaderCount(props.headerCols, props.colCount);
  const headerRowsRef = useRef(headerRows);
  const headerColsRef = useRef(headerCols);
  headerRowsRef.current = headerRows;
  headerColsRef.current = headerCols;

  const rowCountRef = useRef(props.rowCount);
  const colCountRef = useRef(props.colCount);
  rowCountRef.current = props.rowCount;
  colCountRef.current = props.colCount;

  const frozenRows = props.frozenRows ?? 0;
  const frozenCols = props.frozenCols ?? 0;

  const providerRef = useRef(props.provider);
  providerRef.current = props.provider;
  const prefetchOverscanRows = props.prefetchOverscanRows ?? 10;
  const prefetchOverscanCols = props.prefetchOverscanCols ?? 5;
  const interactionMode = props.interactionMode ?? "default";
  const fillHandleEnabled = interactionMode === "default" && (props.onFillCommit != null || props.onFillHandleCommit != null);
  const interactionModeRef = useRef<GridInteractionMode>(interactionMode);
  interactionModeRef.current = interactionMode;

  const enableResizeRef = useRef(props.enableResize ?? false);
  enableResizeRef.current = props.enableResize ?? false;

  const zoomControlledRef = useRef(props.zoom !== undefined);
  zoomControlledRef.current = props.zoom !== undefined;
  const zoomGestureRef = useRef<number | null>(null);

  const [uncontrolledZoom, setUncontrolledZoom] = useState(() => clampZoom(props.zoom ?? 1));
  const zoom = props.zoom !== undefined ? clampZoom(props.zoom) : uncontrolledZoom;
  const zoomRef = useRef(zoom);
  zoomRef.current = zoom;

  const statusId = useId();
  const activeCellId = useMemo(() => `formula-grid-active-cell-${statusId.replace(/:/g, "")}`, [statusId]);
  const [cssTheme, setCssTheme] = useState<Partial<GridTheme>>({});
  const resolvedTheme = useMemo(() => resolveGridTheme(cssTheme, props.theme), [cssTheme, props.theme]);
  const [a11yStatusText, setA11yStatusText] = useState<string>(() =>
    describeCellForA11y({
      selection: null,
      range: null,
      provider: providerRef.current,
      headerRows: headerRowsRef.current,
      headerCols: headerColsRef.current
    })
  );
  const [a11yActiveCell, setA11yActiveCell] = useState<{ row: number; col: number; label: string } | null>(null);

  const announceSelection = useCallback((selection: { row: number; col: number } | null, range: CellRange | null) => {
    const text = describeCellForA11y({
      selection,
      range,
      provider: providerRef.current,
      headerRows: headerRowsRef.current,
      headerCols: headerColsRef.current
    });
    setA11yStatusText((prev) => (prev === text ? prev : text));

    setA11yActiveCell((prev) => {
      if (!selection) return prev === null ? prev : null;

      const label = describeActiveCellLabel(
        selection,
        providerRef.current,
        headerRowsRef.current,
        headerColsRef.current
      );
      if (!label) return prev === null ? prev : null;

      if (prev && prev.row === selection.row && prev.col === selection.col && prev.label === label) return prev;
      return { row: selection.row, col: selection.col, label };
    });
  }, []);

  const rendererFactory = useMemo(
    () =>
      () =>
        new CanvasGridRenderer({
          provider: props.provider,
          rowCount: props.rowCount,
          colCount: props.colCount,
          defaultCellFontFamily: props.defaultCellFontFamily,
          defaultHeaderFontFamily: props.defaultHeaderFontFamily,
          defaultRowHeight: props.defaultRowHeight,
          defaultColWidth: props.defaultColWidth,
          prefetchOverscanRows,
          prefetchOverscanCols,
          imageResolver: props.imageResolver ?? null
        }),
      [
        props.provider,
        props.rowCount,
        props.colCount,
        props.defaultCellFontFamily,
        props.defaultHeaderFontFamily,
        props.defaultRowHeight,
        props.defaultColWidth,
        props.imageResolver,
        prefetchOverscanRows,
        prefetchOverscanCols
      ]
  );

  const maybeEmitScroll = (scroll: { x: number; y: number }, viewport: GridViewportState) => {
    const last = lastEmittedScrollRef.current;
    if (last && last.x === scroll.x && last.y === scroll.y) return;
    // Don't emit on the first observation; treat it as the baseline.
    if (!last) {
      lastEmittedScrollRef.current = { x: scroll.x, y: scroll.y };
      return;
    }

    lastEmittedScrollRef.current = { x: scroll.x, y: scroll.y };
    onScrollRef.current?.({ x: scroll.x, y: scroll.y }, viewport);
  };

  const syncScrollbars = () => {
    const renderer = rendererRef.current;
    if (!renderer) return;

    const viewport = renderer.scroll.getViewportState();
    const scroll = renderer.scroll.getScroll();

    maybeEmitScroll(scroll, viewport);

    const vTrack = vTrackRef.current;
    const vThumb = vThumbRef.current;
    const hTrack = hTrackRef.current;
    const hThumb = hThumbRef.current;

    if (!vTrack || !vThumb || !hTrack || !hThumb) return;
    const minThumbSize = 24 * zoomRef.current;

    // Avoid layout reads during continuous scroll; track sizing is deterministic from the
    // viewport dimensions + the same inset/corner values used to style the tracks.
    const { inset, corner } = scrollbarLayoutRef.current;
    const vTrackSize = Math.max(0, viewport.height - inset - corner);
    const hTrackSize = Math.max(0, viewport.width - inset - corner);

    const frozenHeight = viewport.frozenHeight;
    const frozenWidth = viewport.frozenWidth;

    const vThumbMetrics = computeScrollbarThumb({
      scrollPos: scroll.y,
      viewportSize: Math.max(0, viewport.height - frozenHeight),
      contentSize: Math.max(0, viewport.totalHeight - frozenHeight),
      trackSize: vTrackSize,
      minThumbSize,
      out: scrollbarThumbScratchRef.current.v
    });

    const prevThumb = scrollbarThumbRef.current;
    if (prevThumb.vSize !== vThumbMetrics.size) {
      vThumb.style.height = `${vThumbMetrics.size}px`;
      prevThumb.vSize = vThumbMetrics.size;
    }
    if (prevThumb.vOffset !== vThumbMetrics.offset) {
      vThumb.style.transform = `translateY(${vThumbMetrics.offset}px)`;
      prevThumb.vOffset = vThumbMetrics.offset;
    }

    const hThumbMetrics = computeScrollbarThumb({
      scrollPos: scroll.x,
      viewportSize: Math.max(0, viewport.width - frozenWidth),
      contentSize: Math.max(0, viewport.totalWidth - frozenWidth),
      trackSize: hTrackSize,
      minThumbSize,
      out: scrollbarThumbScratchRef.current.h
    });

    if (prevThumb.hSize !== hThumbMetrics.size) {
      hThumb.style.width = `${hThumbMetrics.size}px`;
      prevThumb.hSize = hThumbMetrics.size;
    }
    if (prevThumb.hOffset !== hThumbMetrics.offset) {
      hThumb.style.transform = `translateX(${hThumbMetrics.offset}px)`;
      prevThumb.hOffset = hThumbMetrics.offset;
    }
  };

  const setZoomInternal = (nextZoom: number, options?: { anchorX?: number; anchorY?: number; force?: boolean }) => {
    const clamped = clampZoom(nextZoom);
    if (zoomControlledRef.current && !options?.force) {
      zoomGestureRef.current = clamped;
      onZoomChangeRef.current?.(clamped);
      return;
    }

    zoomGestureRef.current = null;
    const prev = zoomRef.current;
    zoomRef.current = clamped;
    if (!zoomControlledRef.current) {
      setUncontrolledZoom(clamped);
    }

    const anchor =
      options && (options.anchorX !== undefined || options.anchorY !== undefined)
        ? { anchorX: options.anchorX, anchorY: options.anchorY }
        : undefined;
    rendererRef.current?.setZoom(clamped, anchor);
    syncScrollbars();
    if (!options?.force && prev !== clamped) {
      onZoomChangeRef.current?.(clamped);
    }
  };

  useEffect(() => {
    zoomGestureRef.current = null;
    if (props.zoom === undefined) return;
    const clamped = clampZoom(props.zoom);
    setUncontrolledZoom(clamped);
    setZoomInternal(clamped, { force: true });
  }, [props.zoom]);

  useLayoutEffect(() => {
    // Keep track-size inputs in sync with the committed DOM styles. Avoid mutating refs during
    // render (React concurrent mode), and ensure `syncScrollbars` always uses values that match the
    // currently-rendered scrollbar layout.
    const inset = 2 * zoom;
    const thickness = 10 * zoom;
    const gap = 4 * zoom;
    const corner = inset + thickness + gap;
    scrollbarLayoutRef.current = { inset, corner };

    syncScrollbars();
  }, [zoom]);

  useImperativeHandle(
    props.apiRef,
    (): GridApi => ({
      scrollTo: (x, y) => {
        const renderer = rendererRef.current;
        if (!renderer) return;
        renderer.setScroll(x, y);
        syncScrollbars();
      },
      scrollBy: (dx, dy) => {
        const renderer = rendererRef.current;
        if (!renderer) return;
        renderer.scrollBy(dx, dy);
        syncScrollbars();
      },
      getScroll: () => rendererRef.current?.scroll.getScroll() ?? { x: 0, y: 0 },
      setZoom: (nextZoom) => setZoomInternal(nextZoom),
      getZoom: () => zoomRef.current,
      invalidateImage: (imageId) => {
        rendererRef.current?.invalidateImage(imageId);
      },
      clearImageCache: () => {
        rendererRef.current?.clearImageCache();
      },
      setFrozen: (rows, cols) => {
        const renderer = rendererRef.current;
        if (!renderer) return;
        renderer.setFrozen(rows, cols);
        syncScrollbars();
      },
      setRowHeight: (row, height) => {
        const renderer = rendererRef.current;
        if (!renderer) return;
        renderer.setRowHeight(row, height);
        syncScrollbars();
      },
      setColWidth: (col, width) => {
        const renderer = rendererRef.current;
        if (!renderer) return;
        renderer.setColWidth(col, width);
        syncScrollbars();
      },
      resetRowHeight: (row) => {
        const renderer = rendererRef.current;
        if (!renderer) return;
        renderer.resetRowHeight(row);
        syncScrollbars();
      },
      resetColWidth: (col) => {
        const renderer = rendererRef.current;
        if (!renderer) return;
        renderer.resetColWidth(col);
        syncScrollbars();
      },
      applyAxisSizeOverrides: (overrides, options) => {
        const renderer = rendererRef.current;
        if (!renderer) return;
        renderer.applyAxisSizeOverrides(overrides, options);
        syncScrollbars();
      },
      getRowHeight: (row) => rendererRef.current?.getRowHeight(row) ?? (props.defaultRowHeight ?? 21) * zoomRef.current,
      getColWidth: (col) => rendererRef.current?.getColWidth(col) ?? (props.defaultColWidth ?? 100) * zoomRef.current,
      setSelection: (row, col) => {
        const renderer = rendererRef.current;
        if (!renderer) return;

        const prevSelection = renderer.getSelection();
        const prevRange = renderer.getSelectionRange();
        renderer.setSelection({ row, col });
        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();

        announceSelection(nextSelection, nextRange);
        if (nextSelection) {
          renderer.scrollToCell(nextSelection.row, nextSelection.col, { align: "auto" });
          syncScrollbars();
        }

        if (
          (prevSelection?.row ?? null) !== (nextSelection?.row ?? null) ||
          (prevSelection?.col ?? null) !== (nextSelection?.col ?? null)
        ) {
          onSelectionChangeRef.current?.(nextSelection);
        }

        if (
          (prevRange?.startRow ?? null) !== (nextRange?.startRow ?? null) ||
          (prevRange?.endRow ?? null) !== (nextRange?.endRow ?? null) ||
          (prevRange?.startCol ?? null) !== (nextRange?.startCol ?? null) ||
          (prevRange?.endCol ?? null) !== (nextRange?.endCol ?? null)
        ) {
          onSelectionRangeChangeRef.current?.(nextRange);
        }
      },
      setSelectionRange: (range) => {
        const renderer = rendererRef.current;
        if (!renderer) return;

        const prevSelection = renderer.getSelection();
        const prevRange = renderer.getSelectionRange();
        renderer.setSelectionRange(range);
        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();

        announceSelection(nextSelection, nextRange);
        if (nextSelection) {
          renderer.scrollToCell(nextSelection.row, nextSelection.col, { align: "auto" });
          syncScrollbars();
        }

        if (
          (prevSelection?.row ?? null) !== (nextSelection?.row ?? null) ||
          (prevSelection?.col ?? null) !== (nextSelection?.col ?? null)
        ) {
          onSelectionChangeRef.current?.(nextSelection);
        }

        if (
          (prevRange?.startRow ?? null) !== (nextRange?.startRow ?? null) ||
          (prevRange?.endRow ?? null) !== (nextRange?.endRow ?? null) ||
          (prevRange?.startCol ?? null) !== (nextRange?.startCol ?? null) ||
          (prevRange?.endCol ?? null) !== (nextRange?.endCol ?? null)
        ) {
          onSelectionRangeChangeRef.current?.(nextRange);
        }
      },
      getSelectionRange: () => rendererRef.current?.getSelectionRange() ?? null,
      setSelectionRanges: (ranges, opts) => {
        const renderer = rendererRef.current;
        if (!renderer) return;

        const prevSelection = renderer.getSelection();
        const prevRange = renderer.getSelectionRange();
        renderer.setSelectionRanges(ranges, { activeIndex: opts?.activeIndex });
        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();

        announceSelection(nextSelection, nextRange);
        if (nextSelection) {
          renderer.scrollToCell(nextSelection.row, nextSelection.col, { align: "auto" });
          syncScrollbars();
        }

        if (
          (prevSelection?.row ?? null) !== (nextSelection?.row ?? null) ||
          (prevSelection?.col ?? null) !== (nextSelection?.col ?? null)
        ) {
          onSelectionChangeRef.current?.(nextSelection);
        }

        if (
          (prevRange?.startRow ?? null) !== (nextRange?.startRow ?? null) ||
          (prevRange?.endRow ?? null) !== (nextRange?.endRow ?? null) ||
          (prevRange?.startCol ?? null) !== (nextRange?.startCol ?? null) ||
          (prevRange?.endCol ?? null) !== (nextRange?.endCol ?? null)
        ) {
          onSelectionRangeChangeRef.current?.(nextRange);
        }
      },
      getSelectionRanges: () => rendererRef.current?.getSelectionRanges() ?? [],
      getActiveSelectionRangeIndex: () => rendererRef.current?.getActiveSelectionIndex() ?? 0,
      clearSelection: () => {
        const renderer = rendererRef.current;
        const prevSelection = renderer?.getSelection() ?? null;
        const prevRange = renderer?.getSelectionRange() ?? null;
        renderer?.setSelectionRanges(null);
        renderer?.setRangeSelection(null);
        announceSelection(null, null);
        if (prevSelection) onSelectionChangeRef.current?.(null);
        if (prevRange) onSelectionRangeChangeRef.current?.(null);
      },
      getSelection: () => rendererRef.current?.getSelection() ?? null,
      getPerfStats: () => rendererRef.current?.getPerfStats() ?? null,
      setPerfStatsEnabled: (enabled) => rendererRef.current?.setPerfStatsEnabled(enabled),
      scrollToCell: (row, col, opts) => {
        const renderer = rendererRef.current;
        if (!renderer) return;
        renderer.scrollToCell(row, col, opts);
        syncScrollbars();
      },
      getCellRect: (row, col) => rendererRef.current?.getCellRect(row, col) ?? null,
      getFillHandleRect: () => rendererRef.current?.getFillHandleRect() ?? null,
      getViewportState: () => rendererRef.current?.getViewportState() ?? null,
      setRangeSelection: (range) => rendererRef.current?.setRangeSelection(range),
      setReferenceHighlights: (highlights) => rendererRef.current?.setReferenceHighlights(highlights),
      setRemotePresences: (presences) => rendererRef.current?.setRemotePresences(presences),
      renderImmediately: () => rendererRef.current?.renderImmediately()
    }),
    [props.apiRef]
  );

  useLayoutEffect(() => {
    const container = containerRef.current;
    const gridCanvas = gridCanvasRef.current;
    const contentCanvas = contentCanvasRef.current;
    const selectionCanvas = selectionCanvasRef.current;
    if (!container || !gridCanvas || !contentCanvas || !selectionCanvas) return;

    const nextCssTheme = resolveGridThemeFromCssVars(container);
    setCssTheme((prev) => (partialThemeEqual(prev, nextCssTheme) ? prev : nextCssTheme));

    const renderer = rendererFactory();
    renderer.setHeaders(headersControlled ? headerRows : null, headersControlled ? headerCols : null);
    rendererRef.current = renderer;

    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.setFrozen(frozenRows, frozenCols);
    renderer.setFillHandleEnabled(interactionModeRef.current === "default");
    renderer.setZoom(zoomRef.current);

    const resize = () => {
      const rect = container.getBoundingClientRect();
      const dpr = window.devicePixelRatio || 1;
      renderer.resize(rect.width, rect.height, dpr);
      if (selectionCanvasViewportOriginRef.current) {
        const canvasRect = selectionCanvas.getBoundingClientRect();
        selectionCanvasViewportOriginRef.current = { left: canvasRect.left, top: canvasRect.top };
      }
      syncScrollbars();
    };

    resize();

    const ro = new ResizeObserver(resize);
    ro.observe(container);

    return () => {
      ro.disconnect();
      renderer.destroy();
      rendererRef.current = null;
    };
  }, [rendererFactory, frozenRows, frozenCols]);

  useEffect(() => {
    const renderer = rendererRef.current;
    if (!renderer) return;
    renderer.setHeaders(headersControlled ? headerRows : null, headersControlled ? headerCols : null);
  }, [headersControlled, headerRows, headerCols]);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const refreshTheme = () => {
      const next = resolveGridThemeFromCssVars(container);
      setCssTheme((prev) => (partialThemeEqual(prev, next) ? prev : next));
    };

    const observers: MutationObserver[] = [];
    if (typeof MutationObserver !== "undefined") {
      const observer = new MutationObserver(() => refreshTheme());
      observer.observe(container, { attributes: true, attributeFilter: ["style", "class"] });
      observers.push(observer);

      const root = container.ownerDocument?.documentElement;
      if (root && root !== container) {
        const rootObserver = new MutationObserver(() => refreshTheme());
        rootObserver.observe(root, { attributes: true, attributeFilter: ["style", "class"] });
        observers.push(rootObserver);
      }

      const body = container.ownerDocument?.body;
      if (body && body !== container && body !== root) {
        const bodyObserver = new MutationObserver(() => refreshTheme());
        bodyObserver.observe(body, { attributes: true, attributeFilter: ["style", "class"] });
        observers.push(bodyObserver);
      }
    }

    const canMatchMedia = typeof window !== "undefined" && typeof window.matchMedia === "function";
    const mqlDark = canMatchMedia ? window.matchMedia("(prefers-color-scheme: dark)") : null;
    const mqlContrast = canMatchMedia ? window.matchMedia("(prefers-contrast: more)") : null;
    const mqlForcedColors = canMatchMedia ? window.matchMedia("(forced-colors: active)") : null;
    const onMediaChange = () => refreshTheme();

    const attachMediaListener = (mql: MediaQueryList | null) => {
      if (!mql) return () => {};
      const legacy = mql as unknown as {
        addListener?: (listener: () => void) => void;
        removeListener?: (listener: () => void) => void;
      };

      if (typeof (mql as any).addEventListener === "function") {
        mql.addEventListener("change", onMediaChange);
        return () => mql.removeEventListener("change", onMediaChange);
      }

      legacy.addListener?.(onMediaChange);
      return () => legacy.removeListener?.(onMediaChange);
    };

    const detachDark = attachMediaListener(mqlDark);
    const detachContrast = attachMediaListener(mqlContrast);
    const detachForced = attachMediaListener(mqlForcedColors);

    // Apply the CSS-driven theme immediately (important when the grid mounts under a non-default
    // prefers-color-scheme/contrast setting).
    refreshTheme();

    return () => {
      for (const observer of observers) observer.disconnect();
      detachDark();
      detachContrast();
      detachForced();
    };
  }, []);

  useLayoutEffect(() => {
    rendererRef.current?.setTheme(resolvedTheme);
  }, [resolvedTheme]);

  useEffect(() => {
    const provider = props.provider;
    if (!provider.subscribe) return;

    return provider.subscribe((update) => {
      const renderer = rendererRef.current;
      if (!renderer) return;
      const selection = renderer.getSelection();
      if (!selection) return;

      if (update.type === "invalidateAll") {
        announceSelection(selection, renderer.getSelectionRange());
        return;
      }

      const { range } = update;
      if (
        selection.row >= range.startRow &&
        selection.row < range.endRow &&
        selection.col >= range.startCol &&
        selection.col < range.endCol
      ) {
        announceSelection(selection, renderer.getSelectionRange());
      }
    });
  }, [props.provider, announceSelection]);

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;
    const renderer = rendererRef.current;
    if (!renderer) return;

    const onWheel = (event: WheelEvent) => {
      const renderer = rendererRef.current;
      if (!renderer) return;

      const viewport = renderer.scroll.getViewportState();
      const lineHeight = 16 * zoomRef.current;
      const pageWidth = Math.max(0, viewport.width - viewport.frozenWidth);
      const pageHeight = Math.max(0, viewport.height - viewport.frozenHeight);

      // On most browsers, trackpad pinch-to-zoom is surfaced as a ctrl+wheel event.
      // Treat ctrl/meta+wheel as a zoom gesture (like common spreadsheet apps) instead of scrolling
      // the page or the grid.
      if (event.ctrlKey || event.metaKey) {
        const delta = wheelDeltaToPixels(event.deltaY, event.deltaMode, { lineHeight, pageSize: viewport.height });
        if (delta === 0) return;

        event.preventDefault();

        // Avoid layout reads in the hot zoom path: when the wheel event targets the grid canvases
        // (positioned at 0,0 in the container), `offsetX/offsetY` are already viewport coords.
        const target = event.target;
        const useOffsets = target === container || target instanceof HTMLCanvasElement;
        const rect = useOffsets ? null : container.getBoundingClientRect();
        const anchorX = useOffsets ? event.offsetX : event.clientX - rect!.left;
        const anchorY = useOffsets ? event.offsetY : event.clientY - rect!.top;

        // `delta` is in pixels (after normalization). Use an exponential scale so
        // small trackpad deltas feel smooth while mouse wheels produce ~10% steps.
        const zoomFactor = Math.exp(-delta * 0.001);
        const baseZoom = zoomGestureRef.current ?? zoomRef.current;
        setZoomInternal(baseZoom * zoomFactor, { anchorX, anchorY });
        return;
      }

      let deltaX = wheelDeltaToPixels(event.deltaX, event.deltaMode, { lineHeight, pageSize: pageWidth });
      let deltaY = wheelDeltaToPixels(event.deltaY, event.deltaMode, { lineHeight, pageSize: pageHeight });

      // Common spreadsheet UX: shift+wheel scrolls horizontally.
      if (event.shiftKey) {
        deltaX += deltaY;
        deltaY = 0;
      }

      if (deltaX === 0 && deltaY === 0) return;

      event.preventDefault();
      renderer.scrollBy(deltaX, deltaY);
      syncScrollbars();
    };

    container.addEventListener("wheel", onWheel, { passive: false });
    return () => {
      container.removeEventListener("wheel", onWheel);
    };
  }, []);

  useEffect(() => {
    const selectionCanvas = selectionCanvasRef.current;
    if (!selectionCanvas) return;

    const clamp01 = (value: number) => Math.max(0, Math.min(1, value));
    const RESIZE_HIT_RADIUS_PX = 4;
    const MIN_COL_WIDTH = 24;
    const MIN_ROW_HEIGHT = 16;
    const AUTO_FIT_MAX_COL_WIDTH = 500;
    const AUTO_FIT_MAX_ROW_HEIGHT = 500;
    const DOUBLE_CLICK_MS = 500;
    const isMacPlatform = (() => {
      try {
        const platform = typeof navigator !== "undefined" ? navigator.platform : "";
        return /Mac|iPhone|iPad|iPod/.test(platform);
      } catch {
        return false;
      }
    })();

    const nowMs = () =>
      typeof performance !== "undefined" && typeof performance.now === "function" ? performance.now() : Date.now();

    const TOUCH_PAN_THRESHOLD_PX = 4;

    type TouchPanState = {
      pointerId: number;
      startClientX: number;
      startClientY: number;
      lastClientX: number;
      lastClientY: number;
      moved: boolean;
    };

    type TouchPinchState = { startDistance: number; startZoom: number };

    const touchPointers = new Map<number, { clientX: number; clientY: number }>();
    let touchPan: TouchPanState | null = null;
    let touchPinch: TouchPinchState | null = null;
    let touchTapDisabled = false;

    const stopAutoScroll = () => {
      if (autoScrollFrameRef.current === null) return;
      cancelAnimationFrame(autoScrollFrameRef.current);
      autoScrollFrameRef.current = null;
    };

    const cacheViewportOrigin = () => {
      const rect = selectionCanvas.getBoundingClientRect();
      const origin = { left: rect.left, top: rect.top };
      selectionCanvasViewportOriginRef.current = origin;
      return origin;
    };

    const clearViewportOrigin = () => {
      selectionCanvasViewportOriginRef.current = null;
    };

    const dragViewportPointScratch = { x: 0, y: 0 };
    const hoverViewportPointScratch = { x: 0, y: 0 };
    const pickCellScratch = { row: 0, col: 0 };
    const fillHandlePointerCellScratch = { row: 0, col: 0 };
    const selectionDragRangeScratch: CellRange = { startRow: 0, endRow: 1, startCol: 0, endCol: 1 };
    const fillDragPreviewScratch: FillDragPreview = {
      axis: "vertical",
      unionRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
      targetRange: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 }
    };

    // Cache the last picked cell during pointer-driven selection / fill-handle drags so we can
    // skip redundant work (and allocations) when high-frequency pointermove events stay within
    // the same cell.
    let lastDragPickedRow = Number.NaN;
    let lastDragPickedCol = Number.NaN;
    let lastFillHandlePointerRow = Number.NaN;
    let lastFillHandlePointerCol = Number.NaN;

    const getViewportPoint = (event: { clientX: number; clientY: number }, out?: { x: number; y: number }) => {
      const origin = selectionCanvasViewportOriginRef.current;
      let x: number;
      let y: number;
      if (origin) {
        x = event.clientX - origin.left;
        y = event.clientY - origin.top;
      } else {
        const rect = selectionCanvas.getBoundingClientRect();
        x = event.clientX - rect.left;
        y = event.clientY - rect.top;
      }

      if (out) {
        out.x = x;
        out.y = y;
        return out;
      }

      return { x, y };
    };

    const getResizeHit = (viewportX: number, viewportY: number): ResizeHit | null => {
      const renderer = rendererRef.current;
      if (!renderer) return null;

      const viewport = renderer.scroll.getViewportState();
      const { rowCount, colCount } = renderer.scroll.getCounts();
      if (rowCount === 0 || colCount === 0) return null;

      const frozenWidthClamped = Math.min(viewport.frozenWidth, viewport.width);
      const frozenHeightClamped = Math.min(viewport.frozenHeight, viewport.height);

      const absScrollX = viewport.frozenWidth + viewport.scrollX;
      const absScrollY = viewport.frozenHeight + viewport.scrollY;

      const colAxis = renderer.scroll.cols;
      const rowAxis = renderer.scroll.rows;

      const headersAreControlled = headersControlledRef.current;
      const effectiveHeaderRows = headersAreControlled ? headerRowsRef.current : viewport.frozenRows > 0 ? 1 : 0;
      const effectiveHeaderCols = headersAreControlled ? headerColsRef.current : viewport.frozenCols > 0 ? 1 : 0;

      const headerRowsFrozen = Math.min(effectiveHeaderRows, viewport.frozenRows);
      const headerColsFrozen = Math.min(effectiveHeaderCols, viewport.frozenCols);

      const headerHeight = rowAxis.totalSize(headerRowsFrozen);
      const headerWidth = colAxis.totalSize(headerColsFrozen);

      const inHeaderRow = headerRowsFrozen > 0 && viewportY >= 0 && viewportY <= Math.min(headerHeight, viewport.height);
      const inRowHeaderCol = headerColsFrozen > 0 && viewportX >= 0 && viewportX <= Math.min(headerWidth, viewport.width);

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
    };

    const rangesEqual = (a: CellRange, b: CellRange) =>
      a.startRow === b.startRow && a.endRow === b.endRow && a.startCol === b.startCol && a.endCol === b.endCol;

    const applyFillHandleDrag = (picked: { row: number; col: number }) => {
      const renderer = rendererRef.current;
      if (!renderer) return;
      const state = fillHandleStateRef.current;
      if (!state) return;

      const { rowCount, colCount } = renderer.scroll.getCounts();
      if (rowCount === 0 || colCount === 0) return;
      const headerRows = headerRowsRef.current;
      const headerCols = headerColsRef.current;
      const dataStartRow = headerRows >= rowCount ? 0 : headerRows;
      const dataStartCol = headerCols >= colCount ? 0 : headerCols;

      // Clamp pointerCell into the data region so fill handle drags can't extend selection into headers.
      const pointerRow = clamp(picked.row, dataStartRow, rowCount - 1);
      const pointerCol = clamp(picked.col, dataStartCol, colCount - 1);
      if (pointerRow === lastFillHandlePointerRow && pointerCol === lastFillHandlePointerCol) return;
      lastFillHandlePointerRow = pointerRow;
      lastFillHandlePointerCol = pointerCol;

      fillHandlePointerCellScratch.row = pointerRow;
      fillHandlePointerCellScratch.col = pointerCol;
      const preview = computeFillPreview(state.source, fillHandlePointerCellScratch, fillDragPreviewScratch);
      const union = preview ? fillDragPreviewScratch.unionRange : state.source;
      const targetRange = preview ? fillDragPreviewScratch.targetRange : null;

      const unionStartRow = union.startRow;
      const unionEndRow = union.endRow;
      const unionStartCol = union.startCol;
      const unionEndCol = union.endCol;

      if (
        state.target.startRow === unionStartRow &&
        state.target.endRow === unionEndRow &&
        state.target.startCol === unionStartCol &&
        state.target.endCol === unionEndCol
      ) {
        return;
      }

      const prevHadPreview =
        state.target.startRow !== state.source.startRow ||
        state.target.endRow !== state.source.endRow ||
        state.target.startCol !== state.source.startCol ||
        state.target.endCol !== state.source.endCol;
      const prevPreviewStartRow = state.previewTarget.startRow;
      const prevPreviewEndRow = state.previewTarget.endRow;
      const prevPreviewStartCol = state.previewTarget.startCol;
      const prevPreviewEndCol = state.previewTarget.endCol;

      state.target.startRow = unionStartRow;
      state.target.endRow = unionEndRow;
      state.target.startCol = unionStartCol;
      state.target.endCol = unionEndCol;

      const nextHadPreview = targetRange != null;
      if (targetRange) {
        state.previewTarget.startRow = targetRange.startRow;
        state.previewTarget.endRow = targetRange.endRow;
        state.previewTarget.startCol = targetRange.startCol;
        state.previewTarget.endCol = targetRange.endCol;
      }

      renderer.setFillPreviewRange(state.target);

      const onFillHandleChange = onFillHandleChangeRef.current;
      if (onFillHandleChange) {
        onFillHandleChange({ source: state.source, target: nextHadPreview ? { ...state.target } : state.source });
      }

      const onFillPreviewChange = onFillPreviewChangeRef.current;
      if (onFillPreviewChange) {
        if (!nextHadPreview) {
          if (prevHadPreview) onFillPreviewChange(null);
        } else if (
          !prevHadPreview ||
          state.previewTarget.startRow !== prevPreviewStartRow ||
          state.previewTarget.endRow !== prevPreviewEndRow ||
          state.previewTarget.startCol !== prevPreviewStartCol ||
          state.previewTarget.endCol !== prevPreviewEndCol
        ) {
          onFillPreviewChange({ ...state.previewTarget });
        }
      }
    };

    const applyDragRange = (picked: { row: number; col: number }) => {
      const renderer = rendererRef.current;
      if (!renderer) return;
      const anchor = selectionAnchorRef.current;
      if (!anchor) return;
      if (picked.row === lastDragPickedRow && picked.col === lastDragPickedCol) return;
      lastDragPickedRow = picked.row;
      lastDragPickedCol = picked.col;

      const startRow = Math.min(anchor.row, picked.row);
      const endRow = Math.max(anchor.row, picked.row) + 1;
      const startCol = Math.min(anchor.col, picked.col);
      const endCol = Math.max(anchor.col, picked.col) + 1;

      if (interactionModeRef.current === "rangeSelection") {
        const range: CellRange = { startRow, endRow, startCol, endCol };
        const prevRange = transientRangeRef.current;
        if (
          prevRange &&
          prevRange.startRow === startRow &&
          prevRange.endRow === endRow &&
          prevRange.startCol === startCol &&
          prevRange.endCol === endCol
        ) {
          return;
        }

        transientRangeRef.current = range;
        renderer.setRangeSelection(range);
        announceSelection(renderer.getSelection(), range);
        onRangeSelectionChangeRef.current?.(range);
        return;
      }

      const range = selectionDragRangeScratch;
      range.startRow = startRow;
      range.endRow = endRow;
      range.startCol = startCol;
      range.endCol = endCol;
      if (!renderer.setActiveSelectionRange(range)) return;

      const nextSelection = renderer.getSelection();
      const nextRange = renderer.getSelectionRange();
      announceSelection(nextSelection, nextRange);
      onSelectionRangeChangeRef.current?.(nextRange);
    };

    const scheduleAutoScroll = () => {
      if (autoScrollFrameRef.current !== null) return;

      const tick = () => {
        autoScrollFrameRef.current = null;
        const renderer = rendererRef.current;
        if (!renderer) return;
        if (selectionPointerIdRef.current === null) return;

        const point = lastPointerViewportRef.current;
        if (!point) return;

        const viewport = renderer.scroll.getViewportState();
        if (viewport.width <= 0 || viewport.height <= 0) {
          return;
        }
        const edge = 28;
        const maxSpeed = 24;

        const leftThreshold = viewport.frozenWidth + edge;
        const topThreshold = viewport.frozenHeight + edge;

        const leftFactor = clamp01((leftThreshold - point.x) / edge);
        const rightFactor = clamp01((point.x - (viewport.width - edge)) / edge);
        const topFactor = clamp01((topThreshold - point.y) / edge);
        const bottomFactor = clamp01((point.y - (viewport.height - edge)) / edge);

        const dx = viewport.maxScrollX > 0 ? (rightFactor - leftFactor) * maxSpeed : 0;
        const dy = viewport.maxScrollY > 0 ? (bottomFactor - topFactor) * maxSpeed : 0;

        if (dx === 0 && dy === 0) {
          return;
        }

        const beforeScroll = renderer.scroll.getScroll();
        renderer.scrollBy(dx, dy);
        const afterScroll = renderer.scroll.getScroll();
        syncScrollbars();

        if (beforeScroll.x === afterScroll.x && beforeScroll.y === afterScroll.y) {
          return;
        }

        const clampedX = Math.max(0, Math.min(viewport.width, point.x));
        const clampedY = Math.max(0, Math.min(viewport.height, point.y));
        const picked = renderer.pickCellAt(clampedX, clampedY, pickCellScratch);
        if (picked) {
          if (dragModeRef.current === "fillHandle") {
            applyFillHandleDrag(picked);
          } else {
            applyDragRange(picked);
          }
        }

        autoScrollFrameRef.current = requestAnimationFrame(tick);
      };

      autoScrollFrameRef.current = requestAnimationFrame(tick);
    };

    const onPointerDown = (event: PointerEvent) => {
      const renderer = rendererRef.current;
      if (!renderer) return;

      const touchMode = touchModeRef.current ?? "auto";
      if (event.pointerType === "touch" && touchMode !== "select" && interactionModeRef.current !== "rangeSelection") {
        cacheViewportOrigin();
        // Allow resizing (and double-tap auto-fit) on touch devices even when touch interactions
        // are primarily configured for pan/zoom. Treat a touch that starts on a resize handle as
        // a resize gesture instead of a pan.
        if (
          enableResizeRef.current &&
          touchPointers.size === 0 &&
          resizePointerIdRef.current === null &&
          selectionPointerIdRef.current === null
        ) {
          const point = getViewportPoint(event);
          const hit = getResizeHit(point.x, point.y);
          if (hit) {
            event.preventDefault();
            resizePointerIdRef.current = event.pointerId;
            selectionCanvas.setPointerCapture?.(event.pointerId);

            if (hit.kind === "col") {
              resizeDragRef.current = {
                kind: "col",
                index: hit.index,
                startClient: event.clientX,
                startSize: renderer.getColWidth(hit.index)
              };
              selectionCanvas.style.cursor = "col-resize";
            } else {
              resizeDragRef.current = {
                kind: "row",
                index: hit.index,
                startClient: event.clientY,
                startSize: renderer.getRowHeight(hit.index)
              };
              selectionCanvas.style.cursor = "row-resize";
            }

            return;
          }
        }

        event.preventDefault();
        keyboardAnchorRef.current = null;
        stopAutoScroll();

        selectionPointerIdRef.current = null;
        selectionAnchorRef.current = null;
        lastPointerViewportRef.current = null;
        dragModeRef.current = null;
        fillHandleStateRef.current = null;
        renderer.setFillPreviewRange(null);
        onFillPreviewChangeRef.current?.(null);

        touchPointers.set(event.pointerId, { clientX: event.clientX, clientY: event.clientY });
        try {
          selectionCanvas.setPointerCapture?.(event.pointerId);
        } catch {
          // Ignore capture failures.
        }

        if (touchPointers.size === 1) {
          touchTapDisabled = touchMode === "pan";
          touchPan = {
            pointerId: event.pointerId,
            startClientX: event.clientX,
            startClientY: event.clientY,
            lastClientX: event.clientX,
            lastClientY: event.clientY,
            moved: false
          };
          touchPinch = null;
        } else if (touchPointers.size === 2) {
          const points = Array.from(touchPointers.values());
          const a = points[0]!;
          const b = points[1]!;
          const distance = Math.hypot(b.clientX - a.clientX, b.clientY - a.clientY) || 1;
          touchPinch = { startDistance: distance, startZoom: zoomGestureRef.current ?? zoomRef.current };
          touchPan = null;
          touchTapDisabled = true;
        }
        return;
      }

      // Excel/Sheets behavior: right-clicking inside an existing selection keeps the
      // selection intact; right-clicking outside moves the active cell to the clicked
      // cell.
      //
      // Note: On macOS, Ctrl+click is commonly treated as a right click and fires the
      // `contextmenu` event. Ensure we treat it as a context-click (not additive selection).
      //
      // We intentionally don't require `pointerType` to be present here: tests may dispatch
      // a `MouseEvent("pointerdown")` without the PointerEvent fields.
      const pointerType = (event as unknown as { pointerType?: string }).pointerType;
      const isMousePointer = pointerType === undefined || pointerType === "" || pointerType === "mouse";
      const isContextClick =
        isMousePointer && (event.button === 2 || (isMacPlatform && event.button === 0 && event.ctrlKey && !event.metaKey));
      if (isContextClick) {
        const point = getViewportPoint(event);
        const picked = renderer.pickCellAt(point.x, point.y);
        if (!picked) return;

        // Do not hijack header row/col context menus (if headers are configured).
        const headerRows = headerRowsRef.current;
        const headerCols = headerColsRef.current;
        const isHeaderCell = (headerRows > 0 && picked.row < headerRows) || (headerCols > 0 && picked.col < headerCols);
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

        const selectionChanged =
          (prevSelection?.row ?? null) !== (nextSelection?.row ?? null) ||
          (prevSelection?.col ?? null) !== (nextSelection?.col ?? null);
        const rangeChanged =
          (prevRange?.startRow ?? null) !== (nextRange?.startRow ?? null) ||
          (prevRange?.endRow ?? null) !== (nextRange?.endRow ?? null) ||
          (prevRange?.startCol ?? null) !== (nextRange?.startCol ?? null) ||
          (prevRange?.endCol ?? null) !== (nextRange?.endCol ?? null);

        if (selectionChanged || rangeChanged) {
          announceSelection(nextSelection, nextRange);

          if (selectionChanged) {
            onSelectionChangeRef.current?.(nextSelection);
          }

          if (rangeChanged) {
            onSelectionRangeChangeRef.current?.(nextRange);
          }
        }

        // Best-effort: keep focus on the grid so keyboard navigation continues.
        try {
          containerRef.current?.focus({ preventScroll: true });
        } catch {
          containerRef.current?.focus();
        }

        return;
      }

      event.preventDefault();
      keyboardAnchorRef.current = null;
      cacheViewportOrigin();
      const point = getViewportPoint(event, dragViewportPointScratch);
      lastPointerViewportRef.current = point;
      dragModeRef.current = null;
      fillHandleStateRef.current = null;
      renderer.setFillPreviewRange(null);
      onFillPreviewChangeRef.current?.(null);

      if (enableResizeRef.current) {
        const hit = getResizeHit(point.x, point.y);
        if (hit) {
          resizePointerIdRef.current = event.pointerId;
          selectionCanvas.setPointerCapture?.(event.pointerId);

          if (hit.kind === "col") {
            resizeDragRef.current = {
              kind: "col",
              index: hit.index,
              startClient: event.clientX,
              startSize: renderer.getColWidth(hit.index)
            };
            selectionCanvas.style.cursor = "col-resize";
          } else {
            resizeDragRef.current = {
              kind: "row",
              index: hit.index,
              startClient: event.clientY,
              startSize: renderer.getRowHeight(hit.index)
            };
            selectionCanvas.style.cursor = "row-resize";
          }

          return;
        }
      }

      // Clicking anywhere other than a resize handle breaks the double-click sequence.
      lastResizeClickRef.current = null;

      if (interactionModeRef.current === "default") {
        const source = renderer.getSelectionRange();
        if (source && hitTestSelectionHandle(renderer, point.x, point.y)) {
          selectionPointerIdRef.current = event.pointerId;
          dragModeRef.current = "fillHandle";
          const mode: FillMode = event.altKey ? "formulas" : event.metaKey || event.ctrlKey ? "copy" : "series";
          const target: CellRange = { ...source };
          const previewTarget: CellRange = { ...source };
          fillHandleStateRef.current = { source, target, mode, previewTarget };

          // Seed the fill-handle pointer-cell cache with the selection corner so a "still" drag
          // (high-frequency pointermoves within the same cell) doesn't recompute previews.
          const { rowCount, colCount } = renderer.scroll.getCounts();
          const headerRows = headerRowsRef.current;
          const headerCols = headerColsRef.current;
          const dataStartRow = headerRows >= rowCount ? 0 : headerRows;
          const dataStartCol = headerCols >= colCount ? 0 : headerCols;
          lastFillHandlePointerRow = clamp(source.endRow - 1, dataStartRow, Math.max(0, rowCount - 1));
          lastFillHandlePointerCol = clamp(source.endCol - 1, dataStartCol, Math.max(0, colCount - 1));
          selectionCanvas.setPointerCapture?.(event.pointerId);

          containerRef.current?.focus({ preventScroll: true });
          transientRangeRef.current = null;
          renderer.setRangeSelection(null);

          renderer.setFillPreviewRange(source);
          onFillHandleChangeRef.current?.({ source, target: source });
          scheduleAutoScroll();
          return;
        }
      }

      const picked = renderer.pickCellAt(point.x, point.y);
      if (!picked) {
        clearViewportOrigin();
        return;
      }

      if (interactionModeRef.current === "rangeSelection") {
        selectionPointerIdRef.current = event.pointerId;
        selectionCanvas.setPointerCapture?.(event.pointerId);

        selectionAnchorRef.current = picked;
        lastDragPickedRow = picked.row;
        lastDragPickedCol = picked.col;
        const range: CellRange = {
          startRow: picked.row,
          endRow: picked.row + 1,
          startCol: picked.col,
          endCol: picked.col + 1
        };

        transientRangeRef.current = range;
        renderer.setRangeSelection(range);
        announceSelection(renderer.getSelection(), range);
        onRangeSelectionStartRef.current?.(range);
        scheduleAutoScroll();
        return;
      }

      containerRef.current?.focus({ preventScroll: true });
      transientRangeRef.current = null;
      renderer.setRangeSelection(null);

      const prevSelection = renderer.getSelection();
      const prevRange = renderer.getSelectionRange();

      const isAdditive = event.metaKey || event.ctrlKey;
      const isExtend = event.shiftKey;

      const { rowCount, colCount } = renderer.scroll.getCounts();
      const viewport = renderer.scroll.getViewportState();
      const headerRows = headerRowsRef.current;
      const headerCols = headerColsRef.current;
      const dataStartRow = headerRows >= rowCount ? 0 : headerRows;
      const dataStartCol = headerCols >= colCount ? 0 : headerCols;

      const applyHeaderRange = (range: CellRange, activeCell: { row: number; col: number }) => {
        if (isAdditive) {
          const existing = renderer.getSelectionRanges();
          const nextRanges = [...existing, range];
          renderer.setSelectionRanges(nextRanges, {
            activeIndex: nextRanges.length - 1,
            activeCell
          });
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
        headerRows > 0 &&
        headerCols > 0 &&
        picked.row < headerRows &&
        picked.col < headerCols;
      const isColumnHeader = headerRows > 0 && picked.row < headerRows && picked.col >= headerCols;
      const isRowHeader = headerCols > 0 && picked.col < headerCols && picked.row >= headerRows;

      if (isCornerHeader || isColumnHeader || isRowHeader) {
        const prevSelection = renderer.getSelection();
        const prevRange = renderer.getSelectionRange();

        if (isCornerHeader) {
          const range: CellRange = {
            startRow: dataStartRow,
            endRow: rowCount,
            startCol: dataStartCol,
            endCol: colCount
          };
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

        announceSelection(nextSelection, nextRange);

        if (
          (prevSelection?.row ?? null) !== (nextSelection?.row ?? null) ||
          (prevSelection?.col ?? null) !== (nextSelection?.col ?? null)
        ) {
          onSelectionChangeRef.current?.(nextSelection);
        }

        if (
          (prevRange?.startRow ?? null) !== (nextRange?.startRow ?? null) ||
          (prevRange?.endRow ?? null) !== (nextRange?.endRow ?? null) ||
          (prevRange?.startCol ?? null) !== (nextRange?.startCol ?? null) ||
          (prevRange?.endCol ?? null) !== (nextRange?.endCol ?? null)
        ) {
          onSelectionRangeChangeRef.current?.(nextRange);
        }

        clearViewportOrigin();
        return;
      }

      selectionPointerIdRef.current = event.pointerId;
      selectionCanvas.setPointerCapture?.(event.pointerId);
      lastDragPickedRow = picked.row;
      lastDragPickedCol = picked.col;

      if (isAdditive) {
        selectionAnchorRef.current = picked;
        renderer.addSelectionRange({
          startRow: picked.row,
          endRow: picked.row + 1,
          startCol: picked.col,
          endCol: picked.col + 1
        });
      } else if (isExtend && prevSelection) {
        selectionAnchorRef.current = prevSelection;
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
        selectionAnchorRef.current = picked;
        renderer.setSelection(picked);
      }

      const nextSelection = renderer.getSelection();
      const nextRange = renderer.getSelectionRange();

      announceSelection(nextSelection, nextRange);

      if (
        (prevSelection?.row ?? null) !== (nextSelection?.row ?? null) ||
        (prevSelection?.col ?? null) !== (nextSelection?.col ?? null)
      ) {
        onSelectionChangeRef.current?.(nextSelection);
      }

      if (
        (prevRange?.startRow ?? null) !== (nextRange?.startRow ?? null) ||
        (prevRange?.endRow ?? null) !== (nextRange?.endRow ?? null) ||
        (prevRange?.startCol ?? null) !== (nextRange?.startCol ?? null) ||
        (prevRange?.endCol ?? null) !== (nextRange?.endCol ?? null)
      ) {
        onSelectionRangeChangeRef.current?.(nextRange);
      }

      scheduleAutoScroll();
    };

    const onPointerMove = (event: PointerEvent) => {
      const renderer = rendererRef.current;
      if (!renderer) return;

      if (resizePointerIdRef.current !== null) {
        if (event.pointerId !== resizePointerIdRef.current) return;
        const drag = resizeDragRef.current;
        if (!drag) return;

        event.preventDefault();

        if (drag.kind === "col") {
          const delta = event.clientX - drag.startClient;
          const minWidth = MIN_COL_WIDTH * zoomRef.current;
          renderer.setColWidth(drag.index, Math.max(minWidth, drag.startSize + delta));
        } else {
          const delta = event.clientY - drag.startClient;
          const minHeight = MIN_ROW_HEIGHT * zoomRef.current;
          renderer.setRowHeight(drag.index, Math.max(minHeight, drag.startSize + delta));
        }

        syncScrollbars();
        return;
      }

      const touchMode = touchModeRef.current ?? "auto";
      if (event.pointerType === "touch" && touchMode !== "select" && interactionModeRef.current !== "rangeSelection") {
        const touchPoint = touchPointers.get(event.pointerId);
        if (!touchPoint) return;

        event.preventDefault();
        touchPoint.clientX = event.clientX;
        touchPoint.clientY = event.clientY;

        if (touchPinch && touchPointers.size >= 2) {
          let a: { clientX: number; clientY: number } | null = null;
          let b: { clientX: number; clientY: number } | null = null;
          for (const point of touchPointers.values()) {
            if (!a) {
              a = point;
            } else {
              b = point;
              break;
            }
          }
          if (!a || !b) return;
          const distance = Math.hypot(b.clientX - a.clientX, b.clientY - a.clientY) || 1;
          const centerClientX = (a.clientX + b.clientX) / 2;
          const centerClientY = (a.clientY + b.clientY) / 2;
          const origin = selectionCanvasViewportOriginRef.current ?? cacheViewportOrigin();
          const anchorX = centerClientX - origin.left;
          const anchorY = centerClientY - origin.top;
          setZoomInternal(touchPinch.startZoom * (distance / touchPinch.startDistance), { anchorX, anchorY });
          return;
        }

        if (touchPan && touchPointers.size === 1 && event.pointerId === touchPan.pointerId) {
          const dx = event.clientX - touchPan.lastClientX;
          const dy = event.clientY - touchPan.lastClientY;
          touchPan.lastClientX = event.clientX;
          touchPan.lastClientY = event.clientY;

          if (!touchPan.moved) {
            const totalDx = event.clientX - touchPan.startClientX;
            const totalDy = event.clientY - touchPan.startClientY;
            if (Math.hypot(totalDx, totalDy) < TOUCH_PAN_THRESHOLD_PX) return;
            touchPan.moved = true;
          }

          renderer.scrollBy(-dx, -dy);
          syncScrollbars();
        }

        return;
      }

      if (selectionPointerIdRef.current === null) return;
      if (event.pointerId !== selectionPointerIdRef.current) return;

      event.preventDefault();

      const point = getViewportPoint(event, dragViewportPointScratch);
      lastPointerViewportRef.current = point;

      const picked = renderer.pickCellAt(point.x, point.y, pickCellScratch);
      if (!picked) return;

      if (dragModeRef.current === "fillHandle") {
        applyFillHandleDrag(picked);
      } else {
        applyDragRange(picked);
      }
      scheduleAutoScroll();
    };

    const endDrag = (event: PointerEvent) => {
      if (resizePointerIdRef.current !== null && event.pointerId === resizePointerIdRef.current) {
        const drag = resizeDragRef.current;

        resizePointerIdRef.current = null;
        resizeDragRef.current = null;
        clearViewportOrigin();
        selectionCanvas.style.cursor = "default";
        try {
          selectionCanvas.releasePointerCapture?.(event.pointerId);
        } catch {
          // Some environments throw if the pointer isn't captured; ignore.
        }

        const resizeRenderer = rendererRef.current;
        if (resizeRenderer && drag) {
          const endSize =
            drag.kind === "col" ? resizeRenderer.getColWidth(drag.index) : resizeRenderer.getRowHeight(drag.index);
          const defaultSize =
            drag.kind === "col" ? resizeRenderer.scroll.cols.defaultSize : resizeRenderer.scroll.rows.defaultSize;

          const sizeChanged = endSize !== drag.startSize;

          if (sizeChanged) {
            onAxisSizeChangeRef.current?.({
              kind: drag.kind,
              index: drag.index,
              size: endSize,
              previousSize: drag.startSize,
              defaultSize,
              zoom: zoomRef.current,
              source: "resize"
            });
            lastResizeClickRef.current = null;
            syncScrollbars();
            return;
          }

          // Treat a non-moving resize interaction as a "click" on the handle. If we see two clicks
          // within the double-click threshold, auto-fit (Excel behavior).
          if (event.type === "pointerup" && interactionModeRef.current !== "rangeSelection") {
            const ts = nowMs();
            const last = lastResizeClickRef.current;
            if (last && last.hit.kind === drag.kind && last.hit.index === drag.index && ts - last.time <= DOUBLE_CLICK_MS) {
              lastResizeClickRef.current = null;

              const prevSize = endSize;
              const nextSize =
                drag.kind === "col"
                  ? resizeRenderer.autoFitCol(drag.index, { maxWidth: AUTO_FIT_MAX_COL_WIDTH })
                  : resizeRenderer.autoFitRow(drag.index, { maxHeight: AUTO_FIT_MAX_ROW_HEIGHT });
              syncScrollbars();

              if (nextSize !== prevSize) {
                onAxisSizeChangeRef.current?.({
                  kind: drag.kind,
                  index: drag.index,
                  size: nextSize,
                  previousSize: prevSize,
                  defaultSize,
                  zoom: zoomRef.current,
                  source: "autoFit"
                });
              }
              return;
            }

            lastResizeClickRef.current = { time: ts, hit: { kind: drag.kind, index: drag.index } };
          } else {
            lastResizeClickRef.current = null;
          }

          syncScrollbars();
          return;
        }

        syncScrollbars();
        return;
      }

      const touchMode = touchModeRef.current ?? "auto";
      if (event.pointerType === "touch" && touchMode !== "select" && interactionModeRef.current !== "rangeSelection") {
        if (!touchPointers.has(event.pointerId)) return;
        event.preventDefault();

        const renderer = rendererRef.current;
        const wasPinching = touchPinch !== null;

        touchPointers.delete(event.pointerId);
        if (touchPointers.size < 2) {
          clearViewportOrigin();
        }

        try {
          selectionCanvas.releasePointerCapture?.(event.pointerId);
        } catch {
          // Ignore.
        }

        if (touchPointers.size < 2) {
          touchPinch = null;
        }

        if (touchPan && event.pointerId === touchPan.pointerId) {
          const didMove = touchPan.moved;
          touchPan = null;

          if (!didMove && !touchTapDisabled && renderer) {
            containerRef.current?.focus({ preventScroll: true });
            transientRangeRef.current = null;
            renderer.setRangeSelection(null);

            const prevSelection = renderer.getSelection();
            const prevRange = renderer.getSelectionRange();

            const point = getViewportPoint(event);

            const resizeHit = enableResizeRef.current ? getResizeHit(point.x, point.y) : null;
            if (resizeHit) {
              const ts = nowMs();
              const last = lastResizeClickRef.current;

              if (last && last.hit.kind === resizeHit.kind && last.hit.index === resizeHit.index && ts - last.time <= DOUBLE_CLICK_MS) {
                lastResizeClickRef.current = null;

                const prevSize =
                  resizeHit.kind === "col"
                    ? renderer.getColWidth(resizeHit.index)
                    : renderer.getRowHeight(resizeHit.index);
                const defaultSize =
                  resizeHit.kind === "col" ? renderer.scroll.cols.defaultSize : renderer.scroll.rows.defaultSize;

                const nextSize =
                  resizeHit.kind === "col"
                    ? renderer.autoFitCol(resizeHit.index, { maxWidth: AUTO_FIT_MAX_COL_WIDTH })
                    : renderer.autoFitRow(resizeHit.index, { maxHeight: AUTO_FIT_MAX_ROW_HEIGHT });
                syncScrollbars();

                if (nextSize !== prevSize) {
                  onAxisSizeChangeRef.current?.({
                    kind: resizeHit.kind,
                    index: resizeHit.index,
                    size: nextSize,
                    previousSize: prevSize,
                    defaultSize,
                    zoom: zoomRef.current,
                    source: "autoFit"
                  });
                }
              } else {
                lastResizeClickRef.current = { time: ts, hit: resizeHit };
              }
            } else {
              // Touch taps outside resize handles break the double-tap sequence.
              lastResizeClickRef.current = null;

              const picked = renderer.pickCellAt(point.x, point.y);
              if (picked) {
                const { rowCount, colCount } = renderer.scroll.getCounts();
                const viewport = renderer.scroll.getViewportState();
                const headerRows = headerRowsRef.current;
                const headerCols = headerColsRef.current;
                const dataStartRow = headerRows >= rowCount ? 0 : headerRows;
                const dataStartCol = headerCols >= colCount ? 0 : headerCols;

                const isCornerHeader =
                  headerRows > 0 && headerCols > 0 && picked.row < headerRows && picked.col < headerCols;
                const isColumnHeader = headerRows > 0 && picked.row < headerRows && picked.col >= headerCols;
                const isRowHeader = headerCols > 0 && picked.col < headerCols && picked.row >= headerRows;

                if (isCornerHeader || isColumnHeader || isRowHeader) {
                  const currentSelection = renderer.getSelection();

                if (isCornerHeader) {
                  const range: CellRange = {
                    startRow: dataStartRow,
                    endRow: rowCount,
                    startCol: dataStartCol,
                    endCol: colCount
                  };

                  const activeCell =
                    currentSelection ??
                    ({
                      row: Math.max(dataStartRow, viewport.main.rows.start),
                      col: Math.max(dataStartCol, viewport.main.cols.start)
                    } as const);

                  renderer.setSelectionRange(range, { activeCell });
                } else if (isColumnHeader) {
                  const range: CellRange = {
                    startRow: dataStartRow,
                    endRow: rowCount,
                    startCol: picked.col,
                    endCol: Math.min(colCount, picked.col + 1)
                  };

                  const baseRow = currentSelection ? currentSelection.row : Math.max(dataStartRow, viewport.main.rows.start);
                  renderer.setSelectionRange(range, { activeCell: { row: baseRow, col: picked.col } });
                } else {
                  const range: CellRange = {
                    startRow: picked.row,
                    endRow: Math.min(rowCount, picked.row + 1),
                    startCol: dataStartCol,
                    endCol: colCount
                  };

                  const baseCol = currentSelection ? currentSelection.col : Math.max(dataStartCol, viewport.main.cols.start);
                  renderer.setSelectionRange(range, { activeCell: { row: picked.row, col: baseCol } });
                }
              } else {
                renderer.setSelection(picked);
              }
              const nextSelection = renderer.getSelection();
              const nextRange = renderer.getSelectionRange();
              announceSelection(nextSelection, nextRange);

              if (
                (prevSelection?.row ?? null) !== (nextSelection?.row ?? null) ||
                (prevSelection?.col ?? null) !== (nextSelection?.col ?? null)
              ) {
                onSelectionChangeRef.current?.(nextSelection);
              }

              if (
                (prevRange?.startRow ?? null) !== (nextRange?.startRow ?? null) ||
                (prevRange?.endRow ?? null) !== (nextRange?.endRow ?? null) ||
                (prevRange?.startCol ?? null) !== (nextRange?.startCol ?? null) ||
                (prevRange?.endCol ?? null) !== (nextRange?.endCol ?? null)
              ) {
                onSelectionRangeChangeRef.current?.(nextRange);
              }
            }
            }
          }
        }

        if (touchPointers.size === 0) {
          touchPan = null;
          touchPinch = null;
          touchTapDisabled = false;
          clearViewportOrigin();
        }

        if (wasPinching && touchPointers.size === 1 && !touchPan) {
          const [pointerId, pos] = touchPointers.entries().next().value as [number, { clientX: number; clientY: number }];
          touchPan = {
            pointerId,
            startClientX: pos.clientX,
            startClientY: pos.clientY,
            lastClientX: pos.clientX,
            lastClientY: pos.clientY,
            moved: true
          };
        }

        return;
      }

      if (selectionPointerIdRef.current === null) return;
      if (event.pointerId !== selectionPointerIdRef.current) return;

      const renderer = rendererRef.current;

      selectionPointerIdRef.current = null;
      selectionAnchorRef.current = null;
      lastPointerViewportRef.current = null;
      clearViewportOrigin();
      stopAutoScroll();

      const dragMode = dragModeRef.current;
      dragModeRef.current = null;

      if (renderer && dragMode === "fillHandle") {
        const state = fillHandleStateRef.current;
        fillHandleStateRef.current = null;
        renderer.setFillPreviewRange(null);
        onFillPreviewChangeRef.current?.(null);

        const shouldCommit = event.type === "pointerup";

        if (state && shouldCommit && !rangesEqual(state.source, state.target)) {
          const { source, target } = state;
          const commitResult = onFillCommitRef.current
            ? onFillCommitRef.current({ sourceRange: source, targetRange: state.previewTarget, mode: state.mode })
            : onFillHandleCommitRef.current?.({ source, target });
          void Promise.resolve(commitResult).catch(() => {
            // Consumers own commit error handling; swallow to avoid unhandled rejections.
          });

          const prevSelection = renderer.getSelection();
          const prevRange = renderer.getSelectionRange();

          const ranges = renderer.getSelectionRanges();
          const activeIndex = renderer.getActiveSelectionIndex();
          const updatedRanges = ranges.length === 0 ? [target] : [...ranges];
          updatedRanges[Math.min(activeIndex, updatedRanges.length - 1)] = target;
          renderer.setSelectionRanges(updatedRanges, { activeIndex });

          const nextSelection = renderer.getSelection();
          const nextRange = renderer.getSelectionRange();
          announceSelection(nextSelection, nextRange);

          if (
            (prevSelection?.row ?? null) !== (nextSelection?.row ?? null) ||
            (prevSelection?.col ?? null) !== (nextSelection?.col ?? null)
          ) {
            onSelectionChangeRef.current?.(nextSelection);
          }

          if (
            (prevRange?.startRow ?? null) !== (nextRange?.startRow ?? null) ||
            (prevRange?.endRow ?? null) !== (nextRange?.endRow ?? null) ||
            (prevRange?.startCol ?? null) !== (nextRange?.startCol ?? null) ||
            (prevRange?.endCol ?? null) !== (nextRange?.endCol ?? null)
          ) {
            onSelectionRangeChangeRef.current?.(nextRange);
          }
        }
      } else {
        fillHandleStateRef.current = null;
        renderer?.setFillPreviewRange(null);
        onFillPreviewChangeRef.current?.(null);
      }

      if (interactionModeRef.current === "rangeSelection") {
        const range = transientRangeRef.current;
        if (range) onRangeSelectionEndRef.current?.(range);
      }

      selectionCanvas.style.cursor = "default";

      try {
        selectionCanvas.releasePointerCapture?.(event.pointerId);
      } catch {
        // Some environments throw if the pointer isn't captured; ignore.
      }
    };

    const cancelFillHandleDrag = () => {
      if (dragModeRef.current !== "fillHandle") return;
      const renderer = rendererRef.current;

      const pointerId = selectionPointerIdRef.current;
      selectionPointerIdRef.current = null;
      selectionAnchorRef.current = null;
      lastPointerViewportRef.current = null;
      clearViewportOrigin();
      stopAutoScroll();

      dragModeRef.current = null;
      fillHandleStateRef.current = null;
      renderer?.setFillPreviewRange(null);
      onFillPreviewChangeRef.current?.(null);

      selectionCanvas.style.cursor = "default";

      if (pointerId !== null) {
        try {
          selectionCanvas.releasePointerCapture?.(pointerId);
        } catch {
          // Ignore capture release failures.
        }
      }
    };

    cancelFillHandleDragRef.current = cancelFillHandleDrag;

    const onPointerHover = (event: PointerEvent) => {
      if (resizePointerIdRef.current !== null || selectionPointerIdRef.current !== null) return;
      const renderer = rendererRef.current;
      if (!renderer) return;
      // Avoid layout reads during high-frequency hover events by preferring `offsetX/offsetY`
      // (already in viewport coords for the canvas layers) when safe.
      const useOffsets =
        (event.target === selectionCanvas || event.target instanceof HTMLCanvasElement) &&
        Number.isFinite(event.offsetX) &&
        Number.isFinite(event.offsetY);
      const point = hoverViewportPointScratch;
      if (useOffsets) {
        point.x = event.offsetX;
        point.y = event.offsetY;
      } else {
        getViewportPoint(event, point);
      }

      if (enableResizeRef.current) {
        const hit = getResizeHit(point.x, point.y);
        if (hit) {
          selectionCanvas.style.cursor = hit.kind === "col" ? "col-resize" : "row-resize";
          return;
        }
      }

      if (interactionModeRef.current === "default") {
        if (hitTestSelectionHandle(renderer, point.x, point.y)) {
          selectionCanvas.style.cursor = "crosshair";
          return;
        }
      }

      selectionCanvas.style.cursor = "default";
    };

    const onPointerLeave = () => {
      if (resizePointerIdRef.current !== null || selectionPointerIdRef.current !== null) return;
      selectionCanvas.style.cursor = "default";
    };

    const onDoubleClick = (event: MouseEvent) => {
      if (interactionModeRef.current === "rangeSelection") return;
      const renderer = rendererRef.current;
      if (!renderer) return;
      const useOffsets =
        (event.target === selectionCanvas || event.target instanceof HTMLCanvasElement) &&
        Number.isFinite(event.offsetX) &&
        Number.isFinite(event.offsetY);
      const point = hoverViewportPointScratch;
      if (useOffsets) {
        point.x = event.offsetX;
        point.y = event.offsetY;
      } else {
        getViewportPoint(event, point);
      }

      // Auto-fit is handled via pointer-driven double-click detection (so it works with
      // `pointerdown.preventDefault()` and on touch devices). Keep dblclick for in-cell edit only.
      if (enableResizeRef.current && getResizeHit(point.x, point.y)) return;

      const picked = renderer.pickCellAt(point.x, point.y);
      if (!picked) return;
      onRequestCellEditRef.current?.({ row: picked.row, col: picked.col });
    };

    selectionCanvas.addEventListener("pointerdown", onPointerDown);
    selectionCanvas.addEventListener("pointermove", onPointerMove);
    selectionCanvas.addEventListener("pointermove", onPointerHover);
    selectionCanvas.addEventListener("pointerleave", onPointerLeave);
    selectionCanvas.addEventListener("pointerup", endDrag);
    selectionCanvas.addEventListener("pointercancel", endDrag);
    selectionCanvas.addEventListener("dblclick", onDoubleClick);

    return () => {
      selectionCanvas.removeEventListener("pointerdown", onPointerDown);
      selectionCanvas.removeEventListener("pointermove", onPointerMove);
      selectionCanvas.removeEventListener("pointermove", onPointerHover);
      selectionCanvas.removeEventListener("pointerleave", onPointerLeave);
      selectionCanvas.removeEventListener("pointerup", endDrag);
      selectionCanvas.removeEventListener("pointercancel", endDrag);
      selectionCanvas.removeEventListener("dblclick", onDoubleClick);
      cancelFillHandleDragRef.current = null;
      stopAutoScroll();
    };
  }, []);

  useEffect(() => {
    rendererRef.current?.setRemotePresences(props.remotePresences ?? null);
  }, [props.remotePresences]);

  useEffect(() => {
    if (interactionMode !== "rangeSelection") {
      rendererRef.current?.setRangeSelection(null);
      transientRangeRef.current = null;
    }

    rendererRef.current?.setFillHandleEnabled(fillHandleEnabled);
  }, [interactionMode, fillHandleEnabled]);

  useEffect(() => {
    const vThumb = vThumbRef.current;
    const vTrack = vTrackRef.current;
    if (!vThumb || !vTrack) return;

    const onPointerDown = (event: PointerEvent) => {
      const renderer = rendererRef.current;
      if (!renderer) return;

      event.preventDefault();

      const pointerId = event.pointerId;
      const viewport = renderer.scroll.getViewportState();
      const maxScrollY = viewport.maxScrollY;
      const trackRect = vTrack.getBoundingClientRect();
      const minThumbSize = 24 * zoomRef.current;

      const thumb = computeScrollbarThumb({
        scrollPos: renderer.scroll.getScroll().y,
        viewportSize: Math.max(0, viewport.height - viewport.frozenHeight),
        contentSize: Math.max(0, viewport.totalHeight - viewport.frozenHeight),
        trackSize: trackRect.height,
        minThumbSize
      });

      const thumbTravel = Math.max(0, trackRect.height - thumb.size);
      if (thumbTravel === 0 || maxScrollY === 0) return;

      const startClientY = event.clientY;
      const startScrollY = renderer.scroll.getScroll().y;

      const onMove = (moveEvent: PointerEvent) => {
        if (moveEvent.pointerId !== pointerId) return;
        moveEvent.preventDefault();
        const delta = moveEvent.clientY - startClientY;
        const nextScroll = startScrollY + (delta / thumbTravel) * maxScrollY;
        renderer.setScroll(renderer.scroll.getScroll().x, nextScroll);
        syncScrollbars();
      };

      const cleanup = () => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        window.removeEventListener("pointercancel", onCancel);
      };

      const onUp = (upEvent: PointerEvent) => {
        if (upEvent.pointerId !== pointerId) return;
        cleanup();
      };

      const onCancel = (cancelEvent: PointerEvent) => {
        if (cancelEvent.pointerId !== pointerId) return;
        cleanup();
      };

      window.addEventListener("pointermove", onMove, { passive: false });
      window.addEventListener("pointerup", onUp, { passive: false });
      window.addEventListener("pointercancel", onCancel, { passive: false });
    };

    vThumb.addEventListener("pointerdown", onPointerDown, { passive: false });
    return () => vThumb.removeEventListener("pointerdown", onPointerDown);
  }, []);

  useEffect(() => {
    const vTrack = vTrackRef.current;
    if (!vTrack) return;

    const onTrackPointerDown = (event: PointerEvent) => {
      if (event.target !== vTrack) return;
      const renderer = rendererRef.current;
      if (!renderer) return;

      // Track clicks should scroll but must not start a selection drag.
      event.preventDefault();
      event.stopPropagation();

      const viewport = renderer.scroll.getViewportState();
      const maxScrollY = viewport.maxScrollY;
      const trackRect = vTrack.getBoundingClientRect();
      const minThumbSize = 24 * zoomRef.current;

      const thumb = computeScrollbarThumb({
        scrollPos: renderer.scroll.getScroll().y,
        viewportSize: Math.max(0, viewport.height - viewport.frozenHeight),
        contentSize: Math.max(0, viewport.totalHeight - viewport.frozenHeight),
        trackSize: trackRect.height,
        minThumbSize
      });

      const thumbTravel = Math.max(0, trackRect.height - thumb.size);
      if (thumbTravel === 0 || maxScrollY === 0) return;

      const pointerPos = event.clientY - trackRect.top;
      const targetOffset = pointerPos - thumb.size / 2;
      const clamped = clamp(targetOffset, 0, thumbTravel);
      const nextScroll = (clamped / thumbTravel) * maxScrollY;

      renderer.setScroll(renderer.scroll.getScroll().x, nextScroll);
      syncScrollbars();
    };

    vTrack.addEventListener("pointerdown", onTrackPointerDown, { passive: false });
    return () => vTrack.removeEventListener("pointerdown", onTrackPointerDown);
  }, []);

  useEffect(() => {
    const hThumb = hThumbRef.current;
    const hTrack = hTrackRef.current;
    if (!hThumb || !hTrack) return;

    const onPointerDown = (event: PointerEvent) => {
      const renderer = rendererRef.current;
      if (!renderer) return;

      event.preventDefault();

      const pointerId = event.pointerId;
      const viewport = renderer.scroll.getViewportState();
      const maxScrollX = viewport.maxScrollX;
      const trackRect = hTrack.getBoundingClientRect();
      const minThumbSize = 24 * zoomRef.current;

      const thumb = computeScrollbarThumb({
        scrollPos: renderer.scroll.getScroll().x,
        viewportSize: Math.max(0, viewport.width - viewport.frozenWidth),
        contentSize: Math.max(0, viewport.totalWidth - viewport.frozenWidth),
        trackSize: trackRect.width,
        minThumbSize
      });

      const thumbTravel = Math.max(0, trackRect.width - thumb.size);
      if (thumbTravel === 0 || maxScrollX === 0) return;

      const startClientX = event.clientX;
      const startScrollX = renderer.scroll.getScroll().x;

      const onMove = (moveEvent: PointerEvent) => {
        if (moveEvent.pointerId !== pointerId) return;
        moveEvent.preventDefault();
        const delta = moveEvent.clientX - startClientX;
        const nextScroll = startScrollX + (delta / thumbTravel) * maxScrollX;
        renderer.setScroll(nextScroll, renderer.scroll.getScroll().y);
        syncScrollbars();
      };

      const cleanup = () => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
        window.removeEventListener("pointercancel", onCancel);
      };

      const onUp = (upEvent: PointerEvent) => {
        if (upEvent.pointerId !== pointerId) return;
        cleanup();
      };

      const onCancel = (cancelEvent: PointerEvent) => {
        if (cancelEvent.pointerId !== pointerId) return;
        cleanup();
      };

      window.addEventListener("pointermove", onMove, { passive: false });
      window.addEventListener("pointerup", onUp, { passive: false });
      window.addEventListener("pointercancel", onCancel, { passive: false });
    };

    hThumb.addEventListener("pointerdown", onPointerDown, { passive: false });
    return () => hThumb.removeEventListener("pointerdown", onPointerDown);
  }, []);

  useEffect(() => {
    const hTrack = hTrackRef.current;
    if (!hTrack) return;

    const onTrackPointerDown = (event: PointerEvent) => {
      if (event.target !== hTrack) return;
      const renderer = rendererRef.current;
      if (!renderer) return;

      // Track clicks should scroll but must not start a selection drag.
      event.preventDefault();
      event.stopPropagation();

      const viewport = renderer.scroll.getViewportState();
      const maxScrollX = viewport.maxScrollX;
      const trackRect = hTrack.getBoundingClientRect();
      const minThumbSize = 24 * zoomRef.current;

      const thumb = computeScrollbarThumb({
        scrollPos: renderer.scroll.getScroll().x,
        viewportSize: Math.max(0, viewport.width - viewport.frozenWidth),
        contentSize: Math.max(0, viewport.totalWidth - viewport.frozenWidth),
        trackSize: trackRect.width,
        minThumbSize
      });

      const thumbTravel = Math.max(0, trackRect.width - thumb.size);
      if (thumbTravel === 0 || maxScrollX === 0) return;

      const pointerPos = event.clientX - trackRect.left;
      const targetOffset = pointerPos - thumb.size / 2;
      const clamped = clamp(targetOffset, 0, thumbTravel);
      const nextScroll = (clamped / thumbTravel) * maxScrollX;

      renderer.setScroll(nextScroll, renderer.scroll.getScroll().y);
      syncScrollbars();
    };

    hTrack.addEventListener("pointerdown", onTrackPointerDown, { passive: false });
    return () => hTrack.removeEventListener("pointerdown", onTrackPointerDown);
  }, []);

  const requestCellEdit = (options?: { initialKey?: string }) => {
    const renderer = rendererRef.current;
    if (!renderer) return;
    const selection = renderer.getSelection();
    if (!selection) return;
    onRequestCellEditRef.current?.({ row: selection.row, col: selection.col, initialKey: options?.initialKey });
  };

  const onKeyDown = (event: React.KeyboardEvent<HTMLDivElement>) => {
    const renderer = rendererRef.current;
    if (!renderer) return;

    if (event.key === "Escape" && dragModeRef.current === "fillHandle") {
      event.preventDefault();
      cancelFillHandleDragRef.current?.();
      return;
    }

    if (interactionModeRef.current === "rangeSelection") return;

    const selection = renderer.getSelection();
    const { rowCount, colCount } = renderer.scroll.getCounts();
    if (rowCount === 0 || colCount === 0) return;

    const active =
      selection ??
      {
        row: clampIndex(headerRowsRef.current, 0, rowCount - 1),
        col: clampIndex(headerColsRef.current, 0, colCount - 1)
      };
    const ctrlOrMeta = event.ctrlKey || event.metaKey;
    const dataStartRow = headerRowsRef.current >= rowCount ? 0 : headerRowsRef.current;
    const dataStartCol = headerColsRef.current >= colCount ? 0 : headerColsRef.current;

    const getMergedRangeAtCell = (row: number, col: number): CellRange | null => {
      const provider = providerRef.current;
      const direct = provider.getMergedRangeAt?.(row, col) ?? null;
      if (direct) {
        if (direct.endRow - direct.startRow <= 1 && direct.endCol - direct.startCol <= 1) return null;
        return direct;
      }

      if (provider.getMergedRangesInRange) {
        const candidates = provider.getMergedRangesInRange({ startRow: row, endRow: row + 1, startCol: col, endCol: col + 1 });
        for (const candidate of candidates) {
          if (row < candidate.startRow || row >= candidate.endRow) continue;
          if (col < candidate.startCol || col >= candidate.endCol) continue;
          if (candidate.endRow - candidate.startRow <= 1 && candidate.endCol - candidate.startCol <= 1) continue;
          return candidate;
        }
      }

      return null;
    };

    const applySelectionRange = (range: CellRange) => {
      keyboardAnchorRef.current = null;

      const prevSelection = renderer.getSelection();
      const prevRange = renderer.getSelectionRange();

      renderer.setSelectionRange(range, { activeCell: prevSelection ?? active });

      const nextSelection = renderer.getSelection();
      const nextRange = renderer.getSelectionRange();

      announceSelection(nextSelection, nextRange);

      if (nextSelection) {
        renderer.scrollToCell(nextSelection.row, nextSelection.col, { align: "auto", padding: 8 });
        syncScrollbars();
      }
      if (
        (prevSelection?.row ?? null) !== (nextSelection?.row ?? null) ||
        (prevSelection?.col ?? null) !== (nextSelection?.col ?? null)
      ) {
        onSelectionChangeRef.current?.(nextSelection);
      }

      if (
        (prevRange?.startRow ?? null) !== (nextRange?.startRow ?? null) ||
        (prevRange?.endRow ?? null) !== (nextRange?.endRow ?? null) ||
        (prevRange?.startCol ?? null) !== (nextRange?.startCol ?? null) ||
        (prevRange?.endCol ?? null) !== (nextRange?.endCol ?? null)
      ) {
        onSelectionRangeChangeRef.current?.(nextRange);
      }
    };

    if (event.key === "F2") {
      event.preventDefault();
      if (!selection) renderer.setSelection(active);
      requestCellEdit();
      return;
    }

    // Shift+Space selects the entire row, like Excel.
    if (!ctrlOrMeta && !event.altKey && event.shiftKey && (event.code === "Space" || event.key === " ")) {
      event.preventDefault();

      const startCol = headerColsRef.current >= colCount ? 0 : headerColsRef.current;
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

      const startRow = headerRowsRef.current >= rowCount ? 0 : headerRowsRef.current;
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

      const startRow = headerRowsRef.current >= rowCount ? 0 : headerRowsRef.current;
      const startCol = headerColsRef.current >= colCount ? 0 : headerColsRef.current;
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
      requestCellEdit({ initialKey: event.key });
      return;
    }

    const prevSelection = renderer.getSelection();
    const prevRange = renderer.getSelectionRange();
    const mergedAtActive = getMergedRangeAtCell(active.row, active.col);

    const rangeArea = (range: CellRange) =>
      Math.max(0, range.endRow - range.startRow) * Math.max(0, range.endCol - range.startCol);

    // Excel-like behavior: Tab/Enter moves the active cell *within* the current selection range
    // (wrapping) instead of collapsing selection. For selections that are *only* a merged cell range,
    // we treat them like a single cell and fall back to normal navigation.
    if ((event.key === "Tab" || event.key === "Enter") && prevRange && rangeArea(prevRange) > 1) {
      const current = prevSelection ?? { row: prevRange.startRow, col: prevRange.startCol };
      const activeRow = clamp(current.row, prevRange.startRow, prevRange.endRow - 1);
      const activeCol = clamp(current.col, prevRange.startCol, prevRange.endCol - 1);

      const mergedAtCell = getMergedRangeAtCell(activeRow, activeCol);
      const isSingleMergedSelection =
        mergedAtCell != null &&
        mergedAtCell.startRow === prevRange.startRow &&
        mergedAtCell.endRow === prevRange.endRow &&
        mergedAtCell.startCol === prevRange.startCol &&
        mergedAtCell.endCol === prevRange.endCol;

      if (!isSingleMergedSelection) {
        event.preventDefault();
        keyboardAnchorRef.current = null;

        const backward = event.shiftKey;

        // When selection ranges include merged cells, treat merged ranges as a *single* tab stop
        // (the anchor cell) and skip over interior merged cells.
        const getMergedRangeAt = (row: number, col: number) => getMergedRangeAtCell(row, col);

        const nextCellInRange = (): { row: number; col: number } => {
          let nextRow = activeRow;
          let nextCol = activeCol;
          const maxIterations = Math.min(10_000, Math.max(1, rangeArea(prevRange)) + 10);

          for (let iter = 0; iter < maxIterations; iter++) {
            if (event.key === "Tab") {
              nextCol += backward ? -1 : 1;
              if (nextCol >= prevRange.endCol) {
                nextCol = prevRange.startCol;
                nextRow += 1;
                if (nextRow >= prevRange.endRow) nextRow = prevRange.startRow;
              } else if (nextCol < prevRange.startCol) {
                nextCol = prevRange.endCol - 1;
                nextRow -= 1;
                if (nextRow < prevRange.startRow) nextRow = prevRange.endRow - 1;
              }
            } else {
              nextRow += backward ? -1 : 1;
              if (nextRow >= prevRange.endRow) {
                nextRow = prevRange.startRow;
                nextCol += 1;
                if (nextCol >= prevRange.endCol) nextCol = prevRange.startCol;
              } else if (nextRow < prevRange.startRow) {
                nextRow = prevRange.endRow - 1;
                nextCol -= 1;
                if (nextCol < prevRange.startCol) nextCol = prevRange.endCol - 1;
              }
            }

            while (true) {
              const merged = getMergedRangeAt(nextRow, nextCol);
              if (!merged) break;
              if (merged.startRow === nextRow && merged.startCol === nextCol) break;

              // Forward movement should skip over merged cell interiors to avoid getting stuck on
              // the merged anchor. Backward movement can safely land on an interior cell because
              // the renderer resolves it to the merged anchor cell.
              if (backward) break;

              if (event.key === "Tab") {
                nextCol = merged.endCol;
                if (nextCol >= prevRange.endCol) {
                  nextCol = prevRange.startCol;
                  nextRow += 1;
                  if (nextRow >= prevRange.endRow) nextRow = prevRange.startRow;
                }
              } else {
                nextRow = merged.endRow;
                if (nextRow >= prevRange.endRow) {
                  nextRow = prevRange.startRow;
                  nextCol += 1;
                  if (nextCol >= prevRange.endCol) nextCol = prevRange.startCol;
                }
              }
            }

            return { row: nextRow, col: nextCol };
          }

          return { row: activeRow, col: activeCol };
        };

        const next = nextCellInRange();

        const ranges = renderer.getSelectionRanges();
        const activeIndex = renderer.getActiveSelectionIndex();
        renderer.setSelectionRanges(ranges, { activeIndex, activeCell: next });

        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();

        announceSelection(nextSelection, nextRange);

        if (nextSelection) {
          renderer.scrollToCell(nextSelection.row, nextSelection.col, { align: "auto", padding: 8 });
          syncScrollbars();
        }

        if (
          (prevSelection?.row ?? null) !== (nextSelection?.row ?? null) ||
          (prevSelection?.col ?? null) !== (nextSelection?.col ?? null)
        ) {
          onSelectionChangeRef.current?.(nextSelection);
        }

        return;
      }
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
        nextRow = ctrlOrMeta ? rowCount - 1 : mergedAtActive ? mergedAtActive.endRow : active.row + 1;
        break;
      case "ArrowLeft":
        nextCol = ctrlOrMeta ? dataStartCol : active.col - 1;
        break;
      case "ArrowRight":
        nextCol = ctrlOrMeta ? colCount - 1 : mergedAtActive ? mergedAtActive.endCol : active.col + 1;
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
        nextRow = event.shiftKey
          ? mergedAtActive
            ? mergedAtActive.startRow - 1
            : active.row - 1
          : mergedAtActive
            ? mergedAtActive.endRow
            : active.row + 1;
        break;
      case "Tab":
        nextCol = event.shiftKey
          ? mergedAtActive
            ? mergedAtActive.startCol - 1
            : active.col - 1
          : mergedAtActive
            ? mergedAtActive.endCol
            : active.col + 1;
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
      const anchor = keyboardAnchorRef.current ?? prevSelection ?? active;
      if (!keyboardAnchorRef.current) keyboardAnchorRef.current = anchor;

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
      keyboardAnchorRef.current = null;
      renderer.setSelection({ row: nextRow, col: nextCol });
    }

    const nextSelection = renderer.getSelection();
    const nextRange = renderer.getSelectionRange();

    announceSelection(nextSelection, nextRange);

    if (nextSelection) {
      renderer.scrollToCell(nextSelection.row, nextSelection.col, { align: "auto", padding: 8 });
      syncScrollbars();
    }

    if (
      (prevSelection?.row ?? null) !== (nextSelection?.row ?? null) ||
      (prevSelection?.col ?? null) !== (nextSelection?.col ?? null)
    ) {
      onSelectionChangeRef.current?.(nextSelection);
    }

    if (
      (prevRange?.startRow ?? null) !== (nextRange?.startRow ?? null) ||
      (prevRange?.endRow ?? null) !== (nextRange?.endRow ?? null) ||
      (prevRange?.startCol ?? null) !== (nextRange?.startCol ?? null) ||
      (prevRange?.endCol ?? null) !== (nextRange?.endCol ?? null)
    ) {
      onSelectionRangeChangeRef.current?.(nextRange);
    }
  };

  const scrollbarInset = 2 * zoom;
  const scrollbarThickness = 10 * zoom;
  const scrollbarGap = 4 * zoom;
  const scrollbarCorner = scrollbarInset + scrollbarThickness + scrollbarGap;
  const scrollbarRadius = 6 * zoom;
  const scrollbarThumbInset = 1 * zoom;

  const containerStyle: React.CSSProperties = useMemo(
    () => ({
      position: "relative",
      overflow: "hidden",
      width: "100%",
      height: "100%",
      touchAction: "none",
      background: resolvedTheme.gridBg,
      ...props.style
    }),
    [props.style, resolvedTheme.gridBg]
  );

  const canvasStyle: React.CSSProperties = {
    position: "absolute",
    left: 0,
    top: 0,
    width: "100%",
    height: "100%",
    display: "block",
    touchAction: "none"
  };

  const ariaLabel = props.ariaLabelledBy ? undefined : (props.ariaLabel ?? "Spreadsheet grid");

  return (
    <div
      ref={containerRef}
      style={containerStyle}
      data-testid="canvas-grid"
      tabIndex={0}
      role="grid"
      aria-rowcount={props.rowCount}
      aria-colcount={props.colCount}
      aria-multiselectable="true"
      aria-activedescendant={a11yActiveCell ? activeCellId : undefined}
      aria-label={ariaLabel}
      aria-labelledby={props.ariaLabelledBy}
      aria-describedby={statusId}
      onKeyDown={onKeyDown}
    >
      <div
        id={statusId}
        data-testid="canvas-grid-a11y-status"
        role="status"
        aria-live="polite"
        aria-atomic="true"
        style={SR_ONLY_STYLE}
      >
        {a11yStatusText}
      </div>
      {a11yActiveCell ? (
        <div
          id={activeCellId}
          data-testid="canvas-grid-a11y-active-cell"
          role="gridcell"
          aria-rowindex={a11yActiveCell.row + 1}
          aria-colindex={a11yActiveCell.col + 1}
          aria-selected="true"
          style={SR_ONLY_STYLE}
        >
          {a11yActiveCell.label}
        </div>
      ) : null}
      <canvas
        ref={gridCanvasRef}
        style={{ ...canvasStyle, pointerEvents: "none" }}
        data-testid="canvas-grid-background"
        aria-hidden="true"
      />
      <canvas
        ref={contentCanvasRef}
        style={{ ...canvasStyle, pointerEvents: "none" }}
        data-testid="canvas-grid-content"
        aria-hidden="true"
      />
      <canvas
        ref={selectionCanvasRef}
        style={{ ...canvasStyle, pointerEvents: "auto" }}
        data-testid="canvas-grid-selection"
        aria-hidden="true"
      />

      <div
        ref={vTrackRef}
        aria-hidden="true"
        style={{
          position: "absolute",
          right: scrollbarInset,
          top: scrollbarInset,
          bottom: scrollbarCorner,
          width: scrollbarThickness,
          background: resolvedTheme.scrollbarTrack,
          borderRadius: scrollbarRadius
        }}
      >
        <div
          ref={vThumbRef}
          style={{
            position: "absolute",
            top: 0,
            left: scrollbarThumbInset,
            right: scrollbarThumbInset,
            height: 40,
            background: resolvedTheme.scrollbarThumb,
            borderRadius: scrollbarRadius,
            cursor: "pointer"
          }}
        />
      </div>

      <div
        ref={hTrackRef}
        aria-hidden="true"
        style={{
          position: "absolute",
          left: scrollbarInset,
          right: scrollbarCorner,
          bottom: scrollbarInset,
          height: scrollbarThickness,
          background: resolvedTheme.scrollbarTrack,
          borderRadius: scrollbarRadius
        }}
      >
        <div
          ref={hThumbRef}
          style={{
            position: "absolute",
            top: scrollbarThumbInset,
            bottom: scrollbarThumbInset,
            left: 0,
            width: 40,
            background: resolvedTheme.scrollbarThumb,
            borderRadius: scrollbarRadius,
            cursor: "pointer"
          }}
        />
      </div>
    </div>
  );
}
