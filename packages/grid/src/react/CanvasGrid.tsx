import React, { useCallback, useEffect, useId, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState } from "react";
import type { CellProvider, CellRange } from "../model/CellProvider";
import type { GridPresence } from "../presence/types";
import { CanvasGridRenderer, formatCellDisplayText, type GridPerfStats } from "../rendering/CanvasGridRenderer";
import type { GridTheme } from "../theme/GridTheme";
import { resolveGridTheme } from "../theme/GridTheme";
import { resolveGridThemeFromCssVars } from "../theme/resolveThemeFromCssVars";
import { computeScrollbarThumb } from "../virtualization/scrollbarMath";
import type { GridViewportState } from "../virtualization/VirtualScrollManager";

export type ScrollToCellAlign = "auto" | "start" | "center" | "end";

export interface GridApi {
  scrollTo(x: number, y: number): void;
  scrollBy(deltaX: number, deltaY: number): void;
  getScroll(): { x: number; y: number };
  setFrozen(frozenRows: number, frozenCols: number): void;
  setRowHeight(row: number, height: number): void;
  setColWidth(col: number, width: number): void;
  resetRowHeight(row: number): void;
  resetColWidth(col: number): void;
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
  getViewportState(): GridViewportState | null;
  /**
   * Set a transient range selection overlay.
   *
   * This does not affect the primary grid selection; it's intended for
   * formula-bar range picking UX.
   */
  setRangeSelection(range: CellRange | null): void;
  setRemotePresences(presences: GridPresence[] | null): void;
  renderImmediately(): void;
}

export type GridInteractionMode = "default" | "rangeSelection";

export interface CanvasGridProps {
  provider: CellProvider;
  rowCount: number;
  colCount: number;
  frozenRows?: number;
  frozenCols?: number;
  theme?: Partial<GridTheme>;
  defaultRowHeight?: number;
  defaultColWidth?: number;
  enableResize?: boolean;
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
  onRequestCellEdit?: (request: { row: number; col: number; initialKey?: string }) => void;
  style?: React.CSSProperties;
  ariaLabel?: string;
  ariaLabelledBy?: string;
}

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
    row0 >= 0 && col0 >= 0
      ? toA1Address(row0, col0)
      : `row ${selection.row + 1}, column ${selection.col + 1}`;

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

  const rendererRef = useRef<CanvasGridRenderer | null>(null);
  const onSelectionChangeRef = useRef(props.onSelectionChange);
  const onSelectionRangeChangeRef = useRef(props.onSelectionRangeChange);
  const onRangeSelectionStartRef = useRef(props.onRangeSelectionStart);
  const onRangeSelectionChangeRef = useRef(props.onRangeSelectionChange);
  const onRangeSelectionEndRef = useRef(props.onRangeSelectionEnd);
  const onRequestCellEditRef = useRef(props.onRequestCellEdit);

  onSelectionChangeRef.current = props.onSelectionChange;
  onSelectionRangeChangeRef.current = props.onSelectionRangeChange;
  onRangeSelectionStartRef.current = props.onRangeSelectionStart;
  onRangeSelectionChangeRef.current = props.onRangeSelectionChange;
  onRangeSelectionEndRef.current = props.onRangeSelectionEnd;
  onRequestCellEditRef.current = props.onRequestCellEdit;

  const selectionAnchorRef = useRef<{ row: number; col: number } | null>(null);
  const keyboardAnchorRef = useRef<{ row: number; col: number } | null>(null);
  const selectionPointerIdRef = useRef<number | null>(null);
  const transientRangeRef = useRef<CellRange | null>(null);
  const lastPointerViewportRef = useRef<{ x: number; y: number } | null>(null);
  const autoScrollFrameRef = useRef<number | null>(null);
  const resizePointerIdRef = useRef<number | null>(null);
  const resizeDragRef = useRef<ResizeDragState | null>(null);

  const frozenRows = props.frozenRows ?? 0;
  const frozenCols = props.frozenCols ?? 0;
  const frozenRowsRef = useRef(frozenRows);
  const frozenColsRef = useRef(frozenCols);
  frozenRowsRef.current = frozenRows;
  frozenColsRef.current = frozenCols;

  const headerRows = frozenRows > 0 ? 1 : 0;
  const headerCols = frozenCols > 0 ? 1 : 0;
  const headerRowsRef = useRef(headerRows);
  const headerColsRef = useRef(headerCols);
  headerRowsRef.current = headerRows;
  headerColsRef.current = headerCols;

  const providerRef = useRef(props.provider);
  providerRef.current = props.provider;
  const prefetchOverscanRows = props.prefetchOverscanRows ?? 10;
  const prefetchOverscanCols = props.prefetchOverscanCols ?? 5;
  const interactionMode = props.interactionMode ?? "default";
  const interactionModeRef = useRef<GridInteractionMode>(interactionMode);
  interactionModeRef.current = interactionMode;

  const enableResizeRef = useRef(props.enableResize ?? false);
  enableResizeRef.current = props.enableResize ?? false;

  const statusId = useId();
  const [cssTheme, setCssTheme] = useState<Partial<GridTheme>>({});
  const resolvedTheme = useMemo(() => resolveGridTheme(cssTheme, props.theme), [cssTheme, props.theme]);
  const [a11yStatusText, setA11yStatusText] = useState<string>(() =>
    describeCell(null, null, providerRef.current, headerRowsRef.current, headerColsRef.current)
  );

  const announceSelection = useCallback((selection: { row: number; col: number } | null, range: CellRange | null) => {
    const text = describeCell(selection, range, providerRef.current, headerRowsRef.current, headerColsRef.current);
    setA11yStatusText((prev) => (prev === text ? prev : text));
  }, []);

  const rendererFactory = useMemo(
    () =>
      () =>
        new CanvasGridRenderer({
          provider: props.provider,
          rowCount: props.rowCount,
          colCount: props.colCount,
          defaultRowHeight: props.defaultRowHeight,
          defaultColWidth: props.defaultColWidth,
          prefetchOverscanRows,
          prefetchOverscanCols
        }),
    [
      props.provider,
      props.rowCount,
      props.colCount,
      props.defaultRowHeight,
      props.defaultColWidth,
      prefetchOverscanRows,
      prefetchOverscanCols
    ]
  );

  const syncScrollbars = () => {
    const renderer = rendererRef.current;
    const vTrack = vTrackRef.current;
    const vThumb = vThumbRef.current;
    const hTrack = hTrackRef.current;
    const hThumb = hThumbRef.current;

    if (!renderer || !vTrack || !vThumb || !hTrack || !hThumb) return;

    const viewport = renderer.scroll.getViewportState();
    const scroll = renderer.scroll.getScroll();

    const vTrackSize = vTrack.getBoundingClientRect().height;
    const hTrackSize = hTrack.getBoundingClientRect().width;

    const frozenHeight = viewport.frozenHeight;
    const frozenWidth = viewport.frozenWidth;

    const vThumbMetrics = computeScrollbarThumb({
      scrollPos: scroll.y,
      viewportSize: Math.max(0, viewport.height - frozenHeight),
      contentSize: Math.max(0, viewport.totalHeight - frozenHeight),
      trackSize: vTrackSize
    });

    vThumb.style.height = `${vThumbMetrics.size}px`;
    vThumb.style.transform = `translateY(${vThumbMetrics.offset}px)`;

    const hThumbMetrics = computeScrollbarThumb({
      scrollPos: scroll.x,
      viewportSize: Math.max(0, viewport.width - frozenWidth),
      contentSize: Math.max(0, viewport.totalWidth - frozenWidth),
      trackSize: hTrackSize
    });

    hThumb.style.width = `${hThumbMetrics.size}px`;
    hThumb.style.transform = `translateX(${hThumbMetrics.offset}px)`;
  };

 
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
      getRowHeight: (row) => rendererRef.current?.getRowHeight(row) ?? (props.defaultRowHeight ?? 21),
      getColWidth: (col) => rendererRef.current?.getColWidth(col) ?? (props.defaultColWidth ?? 100),
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
      getViewportState: () => rendererRef.current?.getViewportState() ?? null,
      setRangeSelection: (range) => rendererRef.current?.setRangeSelection(range),
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
    rendererRef.current = renderer;

    renderer.attach({ grid: gridCanvas, content: contentCanvas, selection: selectionCanvas });
    renderer.setFrozen(frozenRows, frozenCols);

    const resize = () => {
      const rect = container.getBoundingClientRect();
      const dpr = window.devicePixelRatio || 1;
      renderer.resize(rect.width, rect.height, dpr);
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

    return () => {
      for (const observer of observers) observer.disconnect();
      detachDark();
      detachContrast();
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
      if (!rendererRef.current) return;
      let deltaX = event.deltaX;
      let deltaY = event.deltaY;

      if (event.deltaMode === 1) {
        // DOM_DELTA_LINE: browsers use a "line" abstraction; normalize to CSS pixels.
        const line = 16;
        deltaX *= line;
        deltaY *= line;
      } else if (event.deltaMode === 2) {
        // DOM_DELTA_PAGE.
        const viewport = rendererRef.current.scroll.getViewportState();
        deltaX *= viewport.width;
        deltaY *= viewport.height;
      }

      // Common UX: shift+wheel scrolls horizontally.
      if (event.shiftKey && deltaX === 0) {
        deltaX = deltaY;
        deltaY = 0;
      }

      if (deltaX === 0 && deltaY === 0) return;

      event.preventDefault();
      rendererRef.current.scrollBy(deltaX, deltaY);
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

    const stopAutoScroll = () => {
      if (autoScrollFrameRef.current === null) return;
      cancelAnimationFrame(autoScrollFrameRef.current);
      autoScrollFrameRef.current = null;
    };

    const getViewportPoint = (event: { clientX: number; clientY: number }) => {
      const rect = selectionCanvas.getBoundingClientRect();
      return { x: event.clientX - rect.left, y: event.clientY - rect.top };
    };

    const getResizeHit = (viewportX: number, viewportY: number): ResizeHit | null => {
      const renderer = rendererRef.current;
      if (!renderer) return null;

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

    const applyDragRange = (picked: { row: number; col: number }) => {
      const renderer = rendererRef.current;
      if (!renderer) return;
      const anchor = selectionAnchorRef.current;
      if (!anchor) return;

      const range: CellRange = {
        startRow: Math.min(anchor.row, picked.row),
        endRow: Math.max(anchor.row, picked.row) + 1,
        startCol: Math.min(anchor.col, picked.col),
        endCol: Math.max(anchor.col, picked.col) + 1
      };

      if (interactionModeRef.current === "rangeSelection") {
        const prevRange = transientRangeRef.current;
        if (
          prevRange &&
          prevRange.startRow === range.startRow &&
          prevRange.endRow === range.endRow &&
          prevRange.startCol === range.startCol &&
          prevRange.endCol === range.endCol
        ) {
          return;
        }

        transientRangeRef.current = range;
        renderer.setRangeSelection(range);
        announceSelection(renderer.getSelection(), range);
        onRangeSelectionChangeRef.current?.(range);
        return;
      }

      const prevRange = renderer.getSelectionRange();
      if (
        prevRange &&
        prevRange.startRow === range.startRow &&
        prevRange.endRow === range.endRow &&
        prevRange.startCol === range.startCol &&
        prevRange.endCol === range.endCol
      ) {
        return;
      }

      const ranges = renderer.getSelectionRanges();
      const activeIndex = renderer.getActiveSelectionIndex();
      const updatedRanges = ranges.length === 0 ? [range] : ranges;
      updatedRanges[Math.min(activeIndex, updatedRanges.length - 1)] = range;
      renderer.setSelectionRanges(updatedRanges, { activeIndex });

      const nextSelection = renderer.getSelection();
      const nextRange = renderer.getSelectionRange();
      announceSelection(nextSelection, nextRange);
      onSelectionRangeChangeRef.current?.(nextRange ?? range);
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
        const picked = renderer.pickCellAt(clampedX, clampedY);
        if (picked) applyDragRange(picked);

        autoScrollFrameRef.current = requestAnimationFrame(tick);
      };

      autoScrollFrameRef.current = requestAnimationFrame(tick);
    };

    const onPointerDown = (event: PointerEvent) => {
      const renderer = rendererRef.current;
      if (!renderer) return;

      event.preventDefault();
      keyboardAnchorRef.current = null;
      const point = getViewportPoint(event);
      lastPointerViewportRef.current = point;

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

      const picked = renderer.pickCellAt(point.x, point.y);
      if (!picked) return;

      if (interactionModeRef.current === "rangeSelection") {
        selectionPointerIdRef.current = event.pointerId;
        selectionCanvas.setPointerCapture?.(event.pointerId);

        selectionAnchorRef.current = picked;
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

        return;
      }

      selectionPointerIdRef.current = event.pointerId;
      selectionCanvas.setPointerCapture?.(event.pointerId);

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
          renderer.setColWidth(drag.index, Math.max(MIN_COL_WIDTH, drag.startSize + delta));
        } else {
          const delta = event.clientY - drag.startClient;
          renderer.setRowHeight(drag.index, Math.max(MIN_ROW_HEIGHT, drag.startSize + delta));
        }

        syncScrollbars();
        return;
      }

      if (selectionPointerIdRef.current === null) return;
      if (event.pointerId !== selectionPointerIdRef.current) return;

      event.preventDefault();

      const point = getViewportPoint(event);
      lastPointerViewportRef.current = point;

      const picked = renderer.pickCellAt(point.x, point.y);
      if (!picked) return;

      applyDragRange(picked);
      scheduleAutoScroll();
    };

    const endDrag = (event: PointerEvent) => {
      if (resizePointerIdRef.current !== null && event.pointerId === resizePointerIdRef.current) {
        resizePointerIdRef.current = null;
        resizeDragRef.current = null;
        selectionCanvas.style.cursor = "default";
        try {
          selectionCanvas.releasePointerCapture?.(event.pointerId);
        } catch {
          // Some environments throw if the pointer isn't captured; ignore.
        }
      }

      if (selectionPointerIdRef.current === null) return;
      if (event.pointerId !== selectionPointerIdRef.current) return;

      selectionPointerIdRef.current = null;
      selectionAnchorRef.current = null;
      lastPointerViewportRef.current = null;
      stopAutoScroll();

      if (interactionModeRef.current === "rangeSelection") {
        const range = transientRangeRef.current;
        if (range) onRangeSelectionEndRef.current?.(range);
      }

      try {
        selectionCanvas.releasePointerCapture?.(event.pointerId);
      } catch {
        // Some environments throw if the pointer isn't captured; ignore.
      }
    };

    const onPointerHover = (event: PointerEvent) => {
      if (!enableResizeRef.current) return;
      if (resizePointerIdRef.current !== null || selectionPointerIdRef.current !== null) return;
      const point = getViewportPoint(event);
      const hit = getResizeHit(point.x, point.y);
      selectionCanvas.style.cursor = hit?.kind === "col" ? "col-resize" : hit?.kind === "row" ? "row-resize" : "default";
    };

    const onPointerLeave = () => {
      if (resizePointerIdRef.current !== null || selectionPointerIdRef.current !== null) return;
      selectionCanvas.style.cursor = "default";
    };

    const onDoubleClick = (event: MouseEvent) => {
      if (interactionModeRef.current === "rangeSelection") return;
      const renderer = rendererRef.current;
      if (!renderer) return;
      const point = getViewportPoint(event);

      if (enableResizeRef.current) {
        const hit = getResizeHit(point.x, point.y);
        if (hit) {
          event.preventDefault();
          if (hit.kind === "col") renderer.autoFitCol(hit.index, { maxWidth: 500 });
          else renderer.autoFitRow(hit.index, { maxHeight: 500 });
          syncScrollbars();
          return;
        }
      }

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
  }, [interactionMode]);

  useEffect(() => {
    const vThumb = vThumbRef.current;
    const vTrack = vTrackRef.current;
    if (!vThumb || !vTrack) return;

    const onPointerDown = (event: PointerEvent) => {
      const renderer = rendererRef.current;
      if (!renderer) return;

      event.preventDefault();

      const viewport = renderer.scroll.getViewportState();
      const maxScrollY = viewport.maxScrollY;
      const trackRect = vTrack.getBoundingClientRect();

      const thumb = computeScrollbarThumb({
        scrollPos: renderer.scroll.getScroll().y,
        viewportSize: Math.max(0, viewport.height - viewport.frozenHeight),
        contentSize: Math.max(0, viewport.totalHeight - viewport.frozenHeight),
        trackSize: trackRect.height
      });

      const thumbTravel = Math.max(0, trackRect.height - thumb.size);
      if (thumbTravel === 0 || maxScrollY === 0) return;

      const startClientY = event.clientY;
      const startScrollY = renderer.scroll.getScroll().y;

      const onMove = (moveEvent: PointerEvent) => {
        moveEvent.preventDefault();
        const delta = moveEvent.clientY - startClientY;
        const nextScroll = startScrollY + (delta / thumbTravel) * maxScrollY;
        renderer.setScroll(renderer.scroll.getScroll().x, nextScroll);
        syncScrollbars();
      };

      const onUp = () => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
      };

      window.addEventListener("pointermove", onMove, { passive: false });
      window.addEventListener("pointerup", onUp, { passive: false });
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

      const viewport = renderer.scroll.getViewportState();
      const maxScrollX = viewport.maxScrollX;
      const trackRect = hTrack.getBoundingClientRect();

      const thumb = computeScrollbarThumb({
        scrollPos: renderer.scroll.getScroll().x,
        viewportSize: Math.max(0, viewport.width - viewport.frozenWidth),
        contentSize: Math.max(0, viewport.totalWidth - viewport.frozenWidth),
        trackSize: trackRect.width
      });

      const thumbTravel = Math.max(0, trackRect.width - thumb.size);
      if (thumbTravel === 0 || maxScrollX === 0) return;

      const startClientX = event.clientX;
      const startScrollX = renderer.scroll.getScroll().x;

      const onMove = (moveEvent: PointerEvent) => {
        moveEvent.preventDefault();
        const delta = moveEvent.clientX - startClientX;
        const nextScroll = startScrollX + (delta / thumbTravel) * maxScrollX;
        renderer.setScroll(nextScroll, renderer.scroll.getScroll().y);
        syncScrollbars();
      };

      const onUp = () => {
        window.removeEventListener("pointermove", onMove);
        window.removeEventListener("pointerup", onUp);
      };

      window.addEventListener("pointermove", onMove, { passive: false });
      window.addEventListener("pointerup", onUp, { passive: false });
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
    const mergedAtActive = providerRef.current.getMergedRangeAt?.(active.row, active.col) ?? null;

    const rangeArea = (range: CellRange) =>
      Math.max(0, range.endRow - range.startRow) * Math.max(0, range.endCol - range.startCol);

    // Excel-like behavior: Tab/Enter moves the active cell *within* the current selection range
    // (wrapping) instead of collapsing selection. For selections that are *only* a merged cell range,
    // we treat them like a single cell and fall back to normal navigation.
    if ((event.key === "Tab" || event.key === "Enter") && prevRange && rangeArea(prevRange) > 1) {
      const current = prevSelection ?? { row: prevRange.startRow, col: prevRange.startCol };
      const activeRow = clamp(current.row, prevRange.startRow, prevRange.endRow - 1);
      const activeCol = clamp(current.col, prevRange.startCol, prevRange.endCol - 1);

      const mergedAtCell = providerRef.current.getMergedRangeAt?.(activeRow, activeCol) ?? null;
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
        const stepRowForward = mergedAtCell ? mergedAtCell.endRow : activeRow + 1;
        const stepRowBackward = mergedAtCell ? mergedAtCell.startRow - 1 : activeRow - 1;
        const stepColForward = mergedAtCell ? mergedAtCell.endCol : activeCol + 1;
        const stepColBackward = mergedAtCell ? mergedAtCell.startCol - 1 : activeCol - 1;

        let nextRow = activeRow;
        let nextCol = activeCol;

        if (event.key === "Tab") {
          if (!backward) {
            if (stepColForward < prevRange.endCol) {
              nextCol = stepColForward;
            } else if (activeRow + 1 < prevRange.endRow) {
              nextRow = activeRow + 1;
              nextCol = prevRange.startCol;
            } else {
              nextRow = prevRange.startRow;
              nextCol = prevRange.startCol;
            }
          } else {
            if (stepColBackward >= prevRange.startCol) {
              nextCol = stepColBackward;
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
            if (stepRowForward < prevRange.endRow) {
              nextRow = stepRowForward;
            } else if (stepColForward < prevRange.endCol) {
              nextRow = prevRange.startRow;
              nextCol = stepColForward;
            } else {
              nextRow = prevRange.startRow;
              nextCol = prevRange.startCol;
            }
          } else {
            if (stepRowBackward >= prevRange.startRow) {
              nextRow = stepRowBackward;
            } else if (stepColBackward >= prevRange.startCol) {
              nextRow = prevRange.endRow - 1;
              nextCol = stepColBackward;
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
        syncScrollbars();

        const nextSelection = renderer.getSelection();
        const nextRange = renderer.getSelectionRange();

        announceSelection(nextSelection, nextRange);

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

    renderer.scrollToCell(nextRow, nextCol, { align: "auto", padding: 8 });
    syncScrollbars();

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
  };

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
    display: "block"
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
        style={{
          position: "absolute",
          width: 1,
          height: 1,
          padding: 0,
          margin: -1,
          overflow: "hidden",
          clip: "rect(0, 0, 0, 0)",
          whiteSpace: "nowrap",
          border: 0
        }}
      >
        {a11yStatusText}
      </div>
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
          right: 2,
          top: 2,
          bottom: 16,
          width: 10,
          background: resolvedTheme.scrollbarTrack,
          borderRadius: 6
        }}
      >
        <div
          ref={vThumbRef}
          style={{
            position: "absolute",
            top: 0,
            left: 1,
            right: 1,
            height: 40,
            background: resolvedTheme.scrollbarThumb,
            borderRadius: 6,
            cursor: "pointer"
          }}
        />
      </div>

      <div
        ref={hTrackRef}
        aria-hidden="true"
        style={{
          position: "absolute",
          left: 2,
          right: 16,
          bottom: 2,
          height: 10,
          background: resolvedTheme.scrollbarTrack,
          borderRadius: 6
        }}
      >
        <div
          ref={hThumbRef}
          style={{
            position: "absolute",
            top: 1,
            bottom: 1,
            left: 0,
            width: 40,
            background: resolvedTheme.scrollbarThumb,
            borderRadius: 6,
            cursor: "pointer"
          }}
        />
      </div>
    </div>
  );
}
