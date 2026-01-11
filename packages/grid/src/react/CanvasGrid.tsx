import React, { useEffect, useId, useImperativeHandle, useLayoutEffect, useMemo, useRef, useState } from "react";
import type { CellProvider, CellRange } from "../model/CellProvider";
import type { GridPresence } from "../presence/types";
import { CanvasGridRenderer, formatCellDisplayText, type GridPerfStats } from "../rendering/CanvasGridRenderer";
import type { GridTheme } from "../theme/GridTheme";
import { resolveGridTheme } from "../theme/GridTheme";
import { readGridThemeFromCssVars } from "../theme/resolveThemeFromCssVars";
import { computeScrollbarThumb } from "../virtualization/scrollbarMath";

export interface GridApi {
  scrollTo(x: number, y: number): void;
  scrollBy(deltaX: number, deltaY: number): void;
  getScroll(): { x: number; y: number };
  setFrozen(frozenRows: number, frozenCols: number): void;
  setSelection(row: number, col: number): void;
  setSelectionRange(range: CellRange | null): void;
  getSelectionRange(): CellRange | null;
  clearSelection(): void;
  getSelection(): { row: number; col: number } | null;
  getPerfStats(): Readonly<GridPerfStats> | null;
  setPerfStatsEnabled(enabled: boolean): void;
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
  style?: React.CSSProperties;
  ariaLabel?: string;
  ariaLabelledBy?: string;
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
  frozenRows: number,
  frozenCols: number
): string {
  if (!selection) return "No cell selected.";

  const row0 = selection.row - frozenRows;
  const col0 = selection.col - frozenCols;
  const address =
    row0 >= 0 && col0 >= 0
      ? toA1Address(row0, col0)
      : `row ${selection.row + 1}, column ${selection.col + 1}`;

  const cell = provider.getCell(selection.row, selection.col);
  const valueText = formatCellDisplayText(cell?.value ?? null);
  const valueDescription = valueText.trim() === "" ? "blank" : valueText;

  let selectionDescription = "none";
  if (range) {
    const startRow0 = range.startRow - frozenRows;
    const startCol0 = range.startCol - frozenCols;
    const endRow0 = range.endRow - frozenRows - 1;
    const endCol0 = range.endCol - frozenCols - 1;
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

  onSelectionChangeRef.current = props.onSelectionChange;
  onSelectionRangeChangeRef.current = props.onSelectionRangeChange;
  onRangeSelectionStartRef.current = props.onRangeSelectionStart;
  onRangeSelectionChangeRef.current = props.onRangeSelectionChange;
  onRangeSelectionEndRef.current = props.onRangeSelectionEnd;

  const selectionAnchorRef = useRef<{ row: number; col: number } | null>(null);
  const selectionPointerIdRef = useRef<number | null>(null);
  const transientRangeRef = useRef<CellRange | null>(null);

  const frozenRows = props.frozenRows ?? 0;
  const frozenCols = props.frozenCols ?? 0;
  const frozenRowsRef = useRef(frozenRows);
  const frozenColsRef = useRef(frozenCols);
  frozenRowsRef.current = frozenRows;
  frozenColsRef.current = frozenCols;

  const providerRef = useRef(props.provider);
  providerRef.current = props.provider;
  const prefetchOverscanRows = props.prefetchOverscanRows ?? 10;
  const prefetchOverscanCols = props.prefetchOverscanCols ?? 5;
  const interactionMode = props.interactionMode ?? "default";
  const interactionModeRef = useRef<GridInteractionMode>(interactionMode);
  interactionModeRef.current = interactionMode;

  const statusId = useId();
  const [cssTheme, setCssTheme] = useState<Partial<GridTheme>>({});
  const resolvedTheme = useMemo(() => resolveGridTheme(cssTheme, props.theme), [cssTheme, props.theme]);
  const [a11yStatusText, setA11yStatusText] = useState<string>(() =>
    describeCell(null, null, providerRef.current, frozenRowsRef.current, frozenColsRef.current)
  );

  const announceSelection = (selection: { row: number; col: number } | null, range: CellRange | null) => {
    const text = describeCell(selection, range, providerRef.current, frozenRowsRef.current, frozenColsRef.current);
    setA11yStatusText((prev) => (prev === text ? prev : text));
  };

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
      setSelection: (row, col) => {
        const renderer = rendererRef.current;
        if (!renderer) return;

        const prevSelection = renderer.getSelection();
        const prevRange = renderer.getSelectionRange();
        renderer.setSelection({ row, col });

        if (!prevSelection || prevSelection.row !== row || prevSelection.col !== col) {
          announceSelection({ row, col }, { startRow: row, endRow: row + 1, startCol: col, endCol: col + 1 });
          onSelectionChangeRef.current?.({ row, col });
        }

        const nextRange: CellRange = { startRow: row, endRow: row + 1, startCol: col, endCol: col + 1 };
        if (
          !prevRange ||
          prevRange.startRow !== nextRange.startRow ||
          prevRange.endRow !== nextRange.endRow ||
          prevRange.startCol !== nextRange.startCol ||
          prevRange.endCol !== nextRange.endCol
        ) {
          announceSelection({ row, col }, nextRange);
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
      clearSelection: () => {
        const renderer = rendererRef.current;
        const prevSelection = renderer?.getSelection() ?? null;
        const prevRange = renderer?.getSelectionRange() ?? null;
        renderer?.setSelectionRange(null);
        renderer?.setRangeSelection(null);
        announceSelection(null, null);
        if (prevSelection) onSelectionChangeRef.current?.(null);
        if (prevRange) onSelectionRangeChangeRef.current?.(null);
      },
      getSelection: () => rendererRef.current?.getSelection() ?? null,
      getPerfStats: () => rendererRef.current?.getPerfStats() ?? null,
      setPerfStatsEnabled: (enabled) => rendererRef.current?.setPerfStatsEnabled(enabled),
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

    const nextCssTheme = readGridThemeFromCssVars(getComputedStyle(container));
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
      const next = readGridThemeFromCssVars(getComputedStyle(container));
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
    const container = containerRef.current;
    if (!container) return;
    const renderer = rendererRef.current;
    if (!renderer) return;

    const onWheel = (event: WheelEvent) => {
      if (!rendererRef.current) return;
      event.preventDefault();
      rendererRef.current.scrollBy(event.deltaX, event.deltaY);
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

    const getPickedCell = (event: PointerEvent) => {
      const renderer = rendererRef.current;
      if (!renderer) return null;
      const rect = selectionCanvas.getBoundingClientRect();
      const x = event.clientX - rect.left;
      const y = event.clientY - rect.top;
      return renderer.pickCellAt(x, y);
    };

    const onPointerDown = (event: PointerEvent) => {
      const renderer = rendererRef.current;
      if (!renderer) return;

      event.preventDefault();
      containerRef.current?.focus({ preventScroll: true });
      const picked = getPickedCell(event);
      if (!picked) return;

      selectionAnchorRef.current = picked;
      selectionPointerIdRef.current = event.pointerId;
      selectionCanvas.setPointerCapture?.(event.pointerId);

      if (interactionModeRef.current === "rangeSelection") {
        const range: CellRange = {
          startRow: picked.row,
          endRow: picked.row + 1,
          startCol: picked.col,
          endCol: picked.col + 1
        };

        transientRangeRef.current = range;
        renderer.setRangeSelection(range);
        onRangeSelectionStartRef.current?.(range);
        return;
      }

      transientRangeRef.current = null;
        renderer.setRangeSelection(null);

        const prevSelection = renderer.getSelection();
        const prevRange = renderer.getSelectionRange();

        renderer.setSelection(picked);

        if (!prevSelection || prevSelection.row !== picked.row || prevSelection.col !== picked.col) {
          announceSelection(picked, { startRow: picked.row, endRow: picked.row + 1, startCol: picked.col, endCol: picked.col + 1 });
          onSelectionChangeRef.current?.(picked);
        }

        const nextRange: CellRange = {
          startRow: picked.row,
          endRow: picked.row + 1,
        startCol: picked.col,
        endCol: picked.col + 1
      };
        if (
          !prevRange ||
          prevRange.startRow !== nextRange.startRow ||
          prevRange.endRow !== nextRange.endRow ||
          prevRange.startCol !== nextRange.startCol ||
          prevRange.endCol !== nextRange.endCol
        ) {
          announceSelection(picked, nextRange);
          onSelectionRangeChangeRef.current?.(nextRange);
        }
      };

    const onPointerMove = (event: PointerEvent) => {
      const renderer = rendererRef.current;
      if (!renderer) return;
      if (selectionPointerIdRef.current === null) return;
      if (event.pointerId !== selectionPointerIdRef.current) return;

      event.preventDefault();

      const anchor = selectionAnchorRef.current;
      if (!anchor) return;

      const picked = getPickedCell(event);
      if (!picked) return;

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

      renderer.setSelectionRange(range);
      announceSelection(renderer.getSelection(), renderer.getSelectionRange());
      onSelectionRangeChangeRef.current?.(range);
    };

    const endDrag = (event: PointerEvent) => {
      if (selectionPointerIdRef.current === null) return;
      if (event.pointerId !== selectionPointerIdRef.current) return;

      selectionPointerIdRef.current = null;
      selectionAnchorRef.current = null;

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

    selectionCanvas.addEventListener("pointerdown", onPointerDown);
    selectionCanvas.addEventListener("pointermove", onPointerMove);
    selectionCanvas.addEventListener("pointerup", endDrag);
    selectionCanvas.addEventListener("pointercancel", endDrag);

    return () => {
      selectionCanvas.removeEventListener("pointerdown", onPointerDown);
      selectionCanvas.removeEventListener("pointermove", onPointerMove);
      selectionCanvas.removeEventListener("pointerup", endDrag);
      selectionCanvas.removeEventListener("pointercancel", endDrag);
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
      aria-label={ariaLabel}
      aria-labelledby={props.ariaLabelledBy}
      aria-describedby={statusId}
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
