import React, { useEffect, useImperativeHandle, useLayoutEffect, useMemo, useRef } from "react";
import type { CellProvider } from "../model/CellProvider";
import type { GridPresence } from "../presence/types";
import { CanvasGridRenderer } from "../rendering/CanvasGridRenderer";
import { computeScrollbarThumb } from "../virtualization/scrollbarMath";

export interface GridApi {
  scrollTo(x: number, y: number): void;
  scrollBy(deltaX: number, deltaY: number): void;
  getScroll(): { x: number; y: number };
  setFrozen(frozenRows: number, frozenCols: number): void;
  setSelection(row: number, col: number): void;
  clearSelection(): void;
  getSelection(): { row: number; col: number } | null;
  setRemotePresences(presences: GridPresence[] | null): void;
  renderImmediately(): void;
}

export interface CanvasGridProps {
  provider: CellProvider;
  rowCount: number;
  colCount: number;
  frozenRows?: number;
  frozenCols?: number;
  defaultRowHeight?: number;
  defaultColWidth?: number;
  remotePresences?: GridPresence[] | null;
  apiRef?: React.Ref<GridApi>;
  onSelectionChange?: (cell: { row: number; col: number } | null) => void;
  style?: React.CSSProperties;
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
  onSelectionChangeRef.current = props.onSelectionChange;

  const frozenRows = props.frozenRows ?? 0;
  const frozenCols = props.frozenCols ?? 0;

  const rendererFactory = useMemo(
    () =>
      () =>
        new CanvasGridRenderer({
          provider: props.provider,
          rowCount: props.rowCount,
          colCount: props.colCount,
          defaultRowHeight: props.defaultRowHeight,
          defaultColWidth: props.defaultColWidth
        }),
    [props.provider, props.rowCount, props.colCount, props.defaultRowHeight, props.defaultColWidth]
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
      setSelection: (row, col) => rendererRef.current?.setSelection({ row, col }),
      clearSelection: () => {
        const renderer = rendererRef.current;
        const prevSelection = renderer?.getSelection() ?? null;
        renderer?.setSelection(null);
        if (prevSelection) onSelectionChangeRef.current?.(null);
      },
      getSelection: () => rendererRef.current?.getSelection() ?? null,
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

    const onPointerDown = (event: PointerEvent) => {
      if (!rendererRef.current) return;
      const rect = selectionCanvas.getBoundingClientRect();
      const x = event.clientX - rect.left;
      const y = event.clientY - rect.top;
      const picked = rendererRef.current.pickCellAt(x, y);
      if (!picked) return;
      const prevSelection = rendererRef.current.getSelection();
      rendererRef.current.setSelection(picked);
      if (!prevSelection || prevSelection.row !== picked.row || prevSelection.col !== picked.col) {
        onSelectionChangeRef.current?.(picked);
      }
    };

    selectionCanvas.addEventListener("pointerdown", onPointerDown);
    return () => selectionCanvas.removeEventListener("pointerdown", onPointerDown);
  }, []);

  useEffect(() => {
    rendererRef.current?.setRemotePresences(props.remotePresences ?? null);
  }, [props.remotePresences]);

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
      background: "#ffffff",
      ...props.style
    }),
    [props.style]
  );

  const canvasStyle: React.CSSProperties = {
    position: "absolute",
    left: 0,
    top: 0,
    width: "100%",
    height: "100%",
    display: "block"
  };

  return (
    <div ref={containerRef} style={containerStyle}>
      <canvas ref={gridCanvasRef} style={{ ...canvasStyle, pointerEvents: "none" }} />
      <canvas ref={contentCanvasRef} style={{ ...canvasStyle, pointerEvents: "none" }} />
      <canvas ref={selectionCanvasRef} style={{ ...canvasStyle, pointerEvents: "auto" }} />

      <div
        ref={vTrackRef}
        style={{
          position: "absolute",
          right: 2,
          top: 2,
          bottom: 16,
          width: 10,
          background: "rgba(0,0,0,0.04)",
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
            background: "rgba(0,0,0,0.25)",
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
          background: "rgba(0,0,0,0.04)",
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
            background: "rgba(0,0,0,0.25)",
            borderRadius: 6,
            cursor: "pointer"
          }}
        />
      </div>
    </div>
  );
}
