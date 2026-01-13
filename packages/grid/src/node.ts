// Node-friendly entrypoint for benchmarks/tests.
//
// `@formula/grid`'s primary entrypoint (`src/index.ts`) re-exports React components (TSX),
// which Node cannot execute under built-in "strip types" TS support (type stripping does not
// transform JSX). The desktop perf suite runs under Node, so it imports from this module.

export type {
  CellProvider,
  CellProviderUpdate,
  CellRange,
  MergedCellRange,
  CellData,
  CellRichText,
  CellRichTextRun,
  CellStyle,
  CellBorderLineStyle,
  CellBorderSpec,
  CellBorders,
  CellDiagonalBorders
} from "./model/CellProvider.ts";

export type {
  CanvasGridImageResolver,
  CanvasGridImageSource,
  CanvasGridRendererOptions,
  GridPerfStats,
  ScrollToCellAlign
} from "./rendering/CanvasGridRenderer.ts";
export { CanvasGridRenderer } from "./rendering/CanvasGridRenderer.ts";
export { DirtyRegionTracker } from "./rendering/DirtyRegionTracker.ts";
export type { Rect } from "./rendering/DirtyRegionTracker.ts";
export { LruCache } from "./utils/LruCache.ts";

export type { GridTheme } from "./theme/GridTheme.ts";
export { DEFAULT_GRID_THEME, resolveGridTheme } from "./theme/GridTheme.ts";

export { VirtualScrollManager } from "./virtualization/VirtualScrollManager.ts";
export type { GridViewportState } from "./virtualization/VirtualScrollManager.ts";

export {
  SR_ONLY_STYLE,
  applySrOnlyStyle,
  describeActiveCellLabel,
  describeCell,
  formatCellDisplayText,
  toA1Address,
  toColumnName
} from "./a11y/a11y.ts";
