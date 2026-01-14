import type { CellCoord, Range, SelectionState } from "./types";

import { resolveCssVar } from "../theme/cssVars.js";

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface GridMetrics {
  getCellRect(cell: CellCoord, out?: Rect): Rect | null;
  /**
   * Visible row indices (0-based) in the current viewport.
   */
  visibleRows: readonly number[];
  /**
   * Visible column indices (0-based) in the current viewport.
   */
  visibleCols: readonly number[];
}

export interface SelectionRenderOptions {
  /**
   * Optional clip rect (in CSS pixels) applied while drawing selection fills/borders.
   * The canvas is still fully cleared before applying the clip.
   */
  clipRect?: Rect;
}

export interface SelectionRenderStyle {
  fillColor: string;
  borderColor: string;
  activeBorderColor: string;
  fillHandleColor: string;
  borderWidth: number;
  activeBorderWidth: number;
  fillHandleSize: number;
}

export interface SelectionRendererOptions {
  /**
   * Element used as the CSS variable root when resolving theme tokens.
   *
   * When omitted, `resolveCssVar` falls back to `document.documentElement`.
   *
   * This is important for per-grid theming where `--formula-grid-*` vars can be
   * overridden on a specific grid root element without affecting the app shell.
   */
  cssVarRoot?: HTMLElement | null;
  /**
   * Partial overrides applied on top of the default theme-derived style.
   *
   * Note: This only applies when `style` is not provided.
   */
  styleOverrides?: Partial<SelectionRenderStyle>;
}

export type SelectionRangeRenderInfo = {
  range: Range;
  rect: Rect;
  edges: { top: boolean; right: boolean; bottom: boolean; left: boolean };
};

export type SelectionRenderDebugInfo = {
  ranges: SelectionRangeRenderInfo[];
  activeCellRect: Rect | null;
  fillHandleRect: Rect | null;
};

function defaultStyleFromTheme(cssVarRoot?: HTMLElement | null): SelectionRenderStyle {
  const resolveVar = (varName: string, fallback: string): string => {
    if (cssVarRoot) {
      return resolveCssVar(varName, { root: cssVarRoot, fallback });
    }
    return resolveCssVar(varName, { fallback });
  };

  const resolveToken = (primary: string, fallback: () => string): string => {
    const value = resolveVar(primary, "");
    if (value) return value;
    return fallback();
  };

  return {
    fillColor: resolveToken("--formula-grid-selection-fill", () => resolveVar("--selection-fill", "transparent")),
    borderColor: resolveToken("--formula-grid-selection-border", () => resolveVar("--selection-border", "transparent")),
    activeBorderColor: resolveToken("--formula-grid-selection-border", () => resolveVar("--selection-border", "transparent")),
    fillHandleColor: resolveToken("--formula-grid-selection-handle", () =>
      resolveToken("--formula-grid-selection-border", () =>
        resolveVar("--selection-border", "transparent"),
      ),
    ),
    borderWidth: 2,
    activeBorderWidth: 3,
    fillHandleSize: 8,
  };
}

export class SelectionRenderer {
  private readonly cssVarRoot: HTMLElement | null;
  private readonly styleOverrides: Partial<SelectionRenderStyle> | null;
  private readonly cellScratch: CellCoord = { row: 0, col: 0 };
  private readonly rectScratchA: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private readonly rectScratchB: Rect = { x: 0, y: 0, width: 0, height: 0 };
  private readonly rectScratchEdge: Rect = { x: 0, y: 0, width: 0, height: 0 };

  constructor(style: SelectionRenderStyle | null = null, options: SelectionRendererOptions = {}) {
    this.style = style;
    this.cssVarRoot = options.cssVarRoot ?? null;
    this.styleOverrides = options.styleOverrides ?? null;
  }

  private style: SelectionRenderStyle | null;
  private lastDebug: SelectionRenderDebugInfo | null = null;
  private lastFillHandleRect: Rect | null = null;
  private lastFillHandleSelection: SelectionState | null = null;

  getLastDebug(): SelectionRenderDebugInfo | null {
    return this.lastDebug;
  }

  getLastFillHandleRect(selection?: SelectionState): Rect | null | undefined {
    if (selection && this.lastFillHandleSelection !== selection) return undefined;
    return this.lastFillHandleRect;
  }

  getFillHandleRect(selection: SelectionState, metrics: GridMetrics, options: SelectionRenderOptions = {}): Rect | null {
    let style = this.style ?? defaultStyleFromTheme(this.cssVarRoot);
    if (!this.style && this.styleOverrides) style = { ...style, ...this.styleOverrides };
    const rect = this.computeFillHandleRect(selection, metrics, style, options);
    this.lastFillHandleRect = rect;
    this.lastFillHandleSelection = selection;
    return rect;
  }

  render(
    ctx: CanvasRenderingContext2D,
    selection: SelectionState,
    metrics: GridMetrics,
    options: SelectionRenderOptions = {}
  ): void {
    let style = this.style ?? defaultStyleFromTheme(this.cssVarRoot);
    if (!this.style && this.styleOverrides) style = { ...style, ...this.styleOverrides };

    // `clearRect` is affected by the current transform. Reset to identity to
    // clear the full backing store regardless of DPR scaling.
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
    ctx.restore();

    const visibleRanges = this.computeVisibleRanges(selection.ranges, metrics, {
      clipRect: options.clipRect,
      borderWidth: style.borderWidth,
    });

    this.lastDebug = {
      ranges: visibleRanges,
      activeCellRect: metrics.getCellRect(selection.active),
      fillHandleRect: null,
    };

    // We draw in CSS pixels; the caller should already have adjusted for DPR.
    if (options.clipRect) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(options.clipRect.x, options.clipRect.y, options.clipRect.width, options.clipRect.height);
      ctx.clip();
    }

    this.renderFill(ctx, visibleRanges, style);
    this.renderBorders(ctx, visibleRanges, style);
    this.renderActiveCell(ctx, selection.active, metrics, style);
    const fillHandleRect = this.renderFillHandle(ctx, selection, metrics, style, options);
    if (this.lastDebug) this.lastDebug.fillHandleRect = fillHandleRect;

    if (options.clipRect) {
      ctx.restore();
    }
  }

  private renderFill(ctx: CanvasRenderingContext2D, ranges: SelectionRangeRenderInfo[], style: SelectionRenderStyle) {
    ctx.save();
    ctx.fillStyle = style.fillColor;
    for (const range of ranges) {
      const rect = range.rect;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    }
    ctx.restore();
  }

  private renderBorders(ctx: CanvasRenderingContext2D, ranges: SelectionRangeRenderInfo[], style: SelectionRenderStyle) {
    ctx.save();
    ctx.strokeStyle = style.borderColor;
    ctx.lineWidth = style.borderWidth;
    for (const range of ranges) {
      this.strokeVisibleRange(ctx, range);
    }
    ctx.restore();
  }

  private renderActiveCell(
    ctx: CanvasRenderingContext2D,
    cell: CellCoord,
    metrics: GridMetrics,
    style: SelectionRenderStyle,
  ) {
    const rect = metrics.getCellRect(cell, this.rectScratchEdge);
    if (!rect) return;
    if (rect.width <= 0 || rect.height <= 0) return;

    ctx.save();
    ctx.strokeStyle = style.activeBorderColor;
    ctx.lineWidth = style.activeBorderWidth;
    ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width - 1, rect.height - 1);

    ctx.restore();
  }

  private renderFillHandle(
    ctx: CanvasRenderingContext2D,
    selection: SelectionState,
    metrics: GridMetrics,
    style: SelectionRenderStyle,
    options: SelectionRenderOptions,
  ): Rect | null {
    const rect = this.computeFillHandleRect(selection, metrics, style, options);
    this.lastFillHandleRect = rect;
    this.lastFillHandleSelection = selection;
    if (!rect) return null;

    ctx.save();
    ctx.fillStyle = style.fillHandleColor;
    ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
    ctx.restore();
    return rect;
  }

  private computeFillHandleRect(
    selection: SelectionState,
    metrics: GridMetrics,
    style: SelectionRenderStyle,
    options: SelectionRenderOptions
  ): Rect | null {
    if (selection.type === "row" || selection.type === "column" || selection.type === "all") return null;

    const range = selection.ranges[selection.activeRangeIndex] ?? selection.ranges[0];
    if (!range) return null;

    const info = this.rangeToVisibleRange(range, metrics, {
      clipRect: options.clipRect,
      borderWidth: style.borderWidth,
    });
    if (!info) return null;
    if (!info.edges.bottom || !info.edges.right) return null;

    const size = style.fillHandleSize;
    if (size <= 0) return null;
    return {
      x: info.rect.x + info.rect.width - size / 2,
      y: info.rect.y + info.rect.height - size / 2,
      width: size,
      height: size,
    };
  }

  private computeVisibleRanges(
    ranges: Range[],
    metrics: GridMetrics,
    options: { clipRect?: Rect; borderWidth: number }
  ): SelectionRangeRenderInfo[] {
    const out: SelectionRangeRenderInfo[] = [];
    for (const range of ranges) {
      const info = this.rangeToVisibleRange(range, metrics, options);
      if (!info) continue;
      out.push(info);
    }
    return out;
  }

  private rangeToVisibleRange(
    range: Range,
    metrics: GridMetrics,
    options: { clipRect?: Rect; borderWidth: number }
  ): SelectionRangeRenderInfo | null {
    const visibleStartRow = firstVisibleIndex(metrics.visibleRows, range.startRow, range.endRow);
    const visibleEndRow = lastVisibleIndex(metrics.visibleRows, range.startRow, range.endRow);
    const visibleStartCol = firstVisibleIndex(metrics.visibleCols, range.startCol, range.endCol);
    const visibleEndCol = lastVisibleIndex(metrics.visibleCols, range.startCol, range.endCol);
    if (visibleStartRow == null || visibleEndRow == null || visibleStartCol == null || visibleEndCol == null) {
      return null;
    }

    const cellScratch = this.cellScratch;
    const rectScratchA = this.rectScratchA;
    const rectScratchB = this.rectScratchB;
    cellScratch.row = visibleStartRow;
    cellScratch.col = visibleStartCol;
    const start = metrics.getCellRect(cellScratch, rectScratchA);
    cellScratch.row = visibleEndRow;
    cellScratch.col = visibleEndCol;
    const end = metrics.getCellRect(cellScratch, rectScratchB);
    if (!start || !end) return null;

    const x = start.x;
    const y = start.y;
    const width = end.x + end.width - start.x;
    const height = end.y + end.height - start.y;
    if (width <= 0 || height <= 0) return null;

    const edges = (() => {
      // Historically we relied on `getCellRect` returning `null` for offscreen cells
      // to decide whether a selection edge should be drawn. With virtual scrolling,
      // `getCellRect` may return valid coordinates for offscreen cells, so we need
      // to determine visibility against the viewport clip instead.
      const clip = options.clipRect;
      if (!clip) {
        return {
          top: visibleStartRow === range.startRow,
          bottom: visibleEndRow === range.endRow,
          left: visibleStartCol === range.startCol,
          right: visibleEndCol === range.endCol,
        };
      }

      const half = Math.max(0, options.borderWidth) / 2;
      const clipRight = clip.x + clip.width;
      const clipBottom = clip.y + clip.height;

      const edgeScratch = this.rectScratchEdge;
      cellScratch.row = range.startRow;
      cellScratch.col = visibleStartCol;
      const topCell = metrics.getCellRect(cellScratch, edgeScratch);
      const yTop = topCell ? topCell.y + 0.5 : null;

      cellScratch.row = range.endRow;
      cellScratch.col = visibleStartCol;
      const bottomCell = metrics.getCellRect(cellScratch, edgeScratch);
      const yBottom = bottomCell ? bottomCell.y + bottomCell.height - 0.5 : null;

      cellScratch.row = visibleStartRow;
      cellScratch.col = range.startCol;
      const leftCell = metrics.getCellRect(cellScratch, edgeScratch);
      const xLeft = leftCell ? leftCell.x + 0.5 : null;

      cellScratch.row = visibleStartRow;
      cellScratch.col = range.endCol;
      const rightCell = metrics.getCellRect(cellScratch, edgeScratch);
      const xRight = rightCell ? rightCell.x + rightCell.width - 0.5 : null;

      return {
        top: yTop != null && yTop + half >= clip.y && yTop - half <= clipBottom,
        bottom: yBottom != null && yBottom - half <= clipBottom && yBottom + half >= clip.y,
        left: xLeft != null && xLeft + half >= clip.x && xLeft - half <= clipRight,
        right: xRight != null && xRight - half <= clipRight && xRight + half >= clip.x,
      };
    })();

    return { range, rect: { x, y, width, height }, edges };
  }

  private strokeVisibleRange(ctx: CanvasRenderingContext2D, range: SelectionRangeRenderInfo): void {
    const { rect, edges } = range;

    // Sub-pixel alignment for crisp borders.
    const xLeft = rect.x + 0.5;
    const xRight = rect.x + rect.width - 0.5;
    const yTop = rect.y + 0.5;
    const yBottom = rect.y + rect.height - 0.5;

    // Fast path: fully visible range => draw a single rectangle.
    if (edges.top && edges.right && edges.bottom && edges.left) {
      ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width - 1, rect.height - 1);
      return;
    }

    // Draw only edges whose boundaries are visible in the viewport. This avoids
    // drawing a misleading "clamped" border when the selection extends offscreen.
    ctx.beginPath();
    if (edges.top) {
      ctx.moveTo(xLeft, yTop);
      ctx.lineTo(xRight, yTop);
    }
    if (edges.bottom) {
      ctx.moveTo(xLeft, yBottom);
      ctx.lineTo(xRight, yBottom);
    }
    if (edges.left) {
      ctx.moveTo(xLeft, yTop);
      ctx.lineTo(xLeft, yBottom);
    }
    if (edges.right) {
      ctx.moveTo(xRight, yTop);
      ctx.lineTo(xRight, yBottom);
    }
    ctx.stroke();
  }
}

function firstVisibleIndex(values: readonly number[], start: number, end: number): number | null {
  for (const value of values) {
    if (value < start) continue;
    if (value > end) break;
    return value;
  }
  return null;
}

function lastVisibleIndex(values: readonly number[], start: number, end: number): number | null {
  for (let i = values.length - 1; i >= 0; i -= 1) {
    const value = values[i]!;
    if (value > end) continue;
    if (value < start) break;
    return value;
  }
  return null;
}
