import { VariableSizeAxis } from "./VariableSizeAxis.ts";

export interface GridViewportState {
  width: number;
  height: number;

  scrollX: number;
  scrollY: number;

  maxScrollX: number;
  maxScrollY: number;

  frozenRows: number;
  frozenCols: number;
  frozenWidth: number;
  frozenHeight: number;

  totalWidth: number;
  totalHeight: number;

  main: {
    rows: { start: number; end: number; offset: number };
    cols: { start: number; end: number; offset: number };
  };
}

export class VirtualScrollManager {
  readonly rows: VariableSizeAxis;
  readonly cols: VariableSizeAxis;

  private rowCount: number;
  private colCount: number;

  private viewportWidth = 0;
  private viewportHeight = 0;

  private frozenRows = 0;
  private frozenCols = 0;

  private scrollX = 0;
  private scrollY = 0;

  private cachedViewportState: GridViewportState | undefined;
  private cachedViewportWidth = -1;
  private cachedViewportHeight = -1;
  private cachedScrollX = -1;
  private cachedScrollY = -1;
  private cachedFrozenRows = -1;
  private cachedFrozenCols = -1;
  private cachedRowCount = -1;
  private cachedColCount = -1;
  private cachedRowsVersion = -1;
  private cachedColsVersion = -1;

  constructor(options: {
    rowCount: number;
    colCount: number;
    defaultRowHeight?: number;
    defaultColWidth?: number;
  }) {
    const defaultRowHeight = options.defaultRowHeight ?? 21;
    const defaultColWidth = options.defaultColWidth ?? 100;

    if (!Number.isSafeInteger(options.rowCount) || options.rowCount < 0) {
      throw new Error(`rowCount must be a non-negative safe integer, got ${options.rowCount}`);
    }
    if (!Number.isSafeInteger(options.colCount) || options.colCount < 0) {
      throw new Error(`colCount must be a non-negative safe integer, got ${options.colCount}`);
    }

    this.rows = new VariableSizeAxis(defaultRowHeight);
    this.cols = new VariableSizeAxis(defaultColWidth);
    this.rowCount = options.rowCount;
    this.colCount = options.colCount;
  }

  getCounts(): { rowCount: number; colCount: number } {
    return { rowCount: this.rowCount, colCount: this.colCount };
  }

  setCounts(rowCount: number, colCount: number): void {
    if (!Number.isSafeInteger(rowCount) || rowCount < 0) {
      throw new Error(`rowCount must be a non-negative safe integer, got ${rowCount}`);
    }
    if (!Number.isSafeInteger(colCount) || colCount < 0) {
      throw new Error(`colCount must be a non-negative safe integer, got ${colCount}`);
    }
    this.rowCount = rowCount;
    this.colCount = colCount;
    this.setScroll(this.scrollX, this.scrollY);
  }

  setViewportSize(width: number, height: number): void {
    this.viewportWidth = Math.max(0, width);
    this.viewportHeight = Math.max(0, height);
    this.setScroll(this.scrollX, this.scrollY);
  }

  setFrozen(frozenRows: number, frozenCols: number): void {
    if (!Number.isSafeInteger(frozenRows) || frozenRows < 0) {
      throw new Error(`frozenRows must be a non-negative safe integer, got ${frozenRows}`);
    }
    if (!Number.isSafeInteger(frozenCols) || frozenCols < 0) {
      throw new Error(`frozenCols must be a non-negative safe integer, got ${frozenCols}`);
    }

    this.frozenRows = Math.min(frozenRows, this.rowCount);
    this.frozenCols = Math.min(frozenCols, this.colCount);
    this.setScroll(this.scrollX, this.scrollY);
  }

  getScroll(): { x: number; y: number } {
    return { x: this.scrollX, y: this.scrollY };
  }

  setScroll(scrollX: number, scrollY: number): void {
    const { maxScrollX, maxScrollY } = this.getMaxScroll();
    this.scrollX = Math.min(Math.max(0, scrollX), maxScrollX);
    this.scrollY = Math.min(Math.max(0, scrollY), maxScrollY);
  }

  scrollBy(deltaX: number, deltaY: number): void {
    this.setScroll(this.scrollX + deltaX, this.scrollY + deltaY);
  }

  getMaxScroll(): { maxScrollX: number; maxScrollY: number } {
    const frozenWidth = this.cols.totalSize(this.frozenCols);
    const frozenHeight = this.rows.totalSize(this.frozenRows);

    const totalWidth = this.cols.totalSize(this.colCount);
    const totalHeight = this.rows.totalSize(this.rowCount);

    const scrollableWidth = Math.max(0, totalWidth - frozenWidth);
    const scrollableHeight = Math.max(0, totalHeight - frozenHeight);

    const viewportScrollableWidth = Math.max(0, this.viewportWidth - frozenWidth);
    const viewportScrollableHeight = Math.max(0, this.viewportHeight - frozenHeight);

    return {
      maxScrollX: Math.max(0, scrollableWidth - viewportScrollableWidth),
      maxScrollY: Math.max(0, scrollableHeight - viewportScrollableHeight)
    };
  }

  getViewportState(): GridViewportState {
    const rowsVersion = this.rows.getVersion();
    const colsVersion = this.cols.getVersion();

    if (
      this.cachedViewportState &&
      this.viewportWidth === this.cachedViewportWidth &&
      this.viewportHeight === this.cachedViewportHeight &&
      this.scrollX === this.cachedScrollX &&
      this.scrollY === this.cachedScrollY &&
      this.frozenRows === this.cachedFrozenRows &&
      this.frozenCols === this.cachedFrozenCols &&
      this.rowCount === this.cachedRowCount &&
      this.colCount === this.cachedColCount &&
      rowsVersion === this.cachedRowsVersion &&
      colsVersion === this.cachedColsVersion
    ) {
      return this.cachedViewportState;
    }

    const frozenWidth = this.cols.totalSize(this.frozenCols);
    const frozenHeight = this.rows.totalSize(this.frozenRows);

    const totalWidth = this.cols.totalSize(this.colCount);
    const totalHeight = this.rows.totalSize(this.rowCount);

    const viewportScrollableWidth = Math.max(0, this.viewportWidth - frozenWidth);
    const viewportScrollableHeight = Math.max(0, this.viewportHeight - frozenHeight);

    const scrollableWidth = Math.max(0, totalWidth - frozenWidth);
    const scrollableHeight = Math.max(0, totalHeight - frozenHeight);

    const maxScrollX = Math.max(0, scrollableWidth - viewportScrollableWidth);
    const maxScrollY = Math.max(0, scrollableHeight - viewportScrollableHeight);

    const absScrollX = frozenWidth + this.scrollX;
    const absScrollY = frozenHeight + this.scrollY;

    const colsRange =
      this.colCount === this.frozenCols
        ? { start: this.colCount, end: this.colCount, offset: 0 }
        : this.cols.visibleRange(absScrollX, viewportScrollableWidth, {
            min: this.frozenCols,
            maxExclusive: this.colCount
          });

    const rowsRange =
      this.rowCount === this.frozenRows
        ? { start: this.rowCount, end: this.rowCount, offset: 0 }
        : this.rows.visibleRange(absScrollY, viewportScrollableHeight, {
            min: this.frozenRows,
            maxExclusive: this.rowCount
          });

    const viewport: GridViewportState = {
      width: this.viewportWidth,
      height: this.viewportHeight,
      scrollX: this.scrollX,
      scrollY: this.scrollY,
      maxScrollX,
      maxScrollY,
      frozenRows: this.frozenRows,
      frozenCols: this.frozenCols,
      frozenWidth,
      frozenHeight,
      totalWidth,
      totalHeight,
      main: {
        rows: rowsRange,
        cols: colsRange
      }
    };

    this.cachedViewportState = viewport;
    this.cachedViewportWidth = this.viewportWidth;
    this.cachedViewportHeight = this.viewportHeight;
    this.cachedScrollX = this.scrollX;
    this.cachedScrollY = this.scrollY;
    this.cachedFrozenRows = this.frozenRows;
    this.cachedFrozenCols = this.frozenCols;
    this.cachedRowCount = this.rowCount;
    this.cachedColCount = this.colCount;
    this.cachedRowsVersion = rowsVersion;
    this.cachedColsVersion = colsVersion;

    return viewport;
  }
}
