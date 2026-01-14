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
export { MockCellProvider } from "./model/MockCellProvider.ts";

export type {
  CanvasGridImageResolver,
  CanvasGridImageSource,
  CanvasGridRendererOptions,
  GridViewportChangeEvent,
  GridViewportChangeListener,
  GridViewportChangeReason,
  GridViewportSubscriptionOptions,
  GridPerfStats,
  ScrollToCellAlign
} from "./rendering/CanvasGridRenderer.ts";
export { CanvasGridRenderer } from "./rendering/CanvasGridRenderer.ts";
export { DirtyRegionTracker } from "./rendering/DirtyRegionTracker.ts";
export type { Rect } from "./rendering/DirtyRegionTracker.ts";
export { DEFAULT_GRID_FONT_FAMILY, DEFAULT_GRID_MONOSPACE_FONT_FAMILY } from "./rendering/defaultFontFamilies.ts";
export { LruCache } from "./utils/LruCache.ts";
export { MAX_GRID_ZOOM, MIN_GRID_ZOOM, clampZoom } from "./utils/zoomMath.ts";

export type { GridTheme } from "./theme/GridTheme.ts";
export { DEFAULT_GRID_THEME, resolveGridTheme } from "./theme/GridTheme.ts";
export {
  GRID_THEME_CSS_VAR_NAMES,
  readGridThemeFromCssVars,
  resolveCssVarValue,
  resolveGridThemeFromCssVars
} from "./theme/resolveThemeFromCssVars.ts";

export { VirtualScrollManager } from "./virtualization/VirtualScrollManager.ts";
export type { GridViewportState } from "./virtualization/VirtualScrollManager.ts";
export { VariableSizeAxis } from "./virtualization/VariableSizeAxis.ts";
export type { AxisVisibleRange } from "./virtualization/VariableSizeAxis.ts";
export { alignScrollToDevicePixels } from "./virtualization/alignScrollToDevicePixels.ts";
export type { ScrollbarThumb } from "./virtualization/scrollbarMath.ts";
export { computeScrollbarThumb } from "./virtualization/scrollbarMath.ts";

export { wheelDeltaToPixels } from "./react/wheelDeltaToPixels.ts";
export type { WheelDeltaToPixelsOptions } from "./react/wheelDeltaToPixels.ts";

export type { FillDragAxis, FillDragCommit, FillDragPreview, FillMode, RectLike } from "./interaction/fillHandle.ts";
export { computeFillPreview, hitTestSelectionHandle } from "./interaction/fillHandle.ts";

export {
  SR_ONLY_STYLE,
  applySrOnlyStyle,
  describeActiveCellLabel,
  describeCell,
  describeCellForA11y,
  formatCellDisplayText,
  toA1Address,
  toColumnName
} from "./a11y/a11y.ts";
