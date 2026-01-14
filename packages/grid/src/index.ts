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
} from "./model/CellProvider";
export { MockCellProvider } from "./model/MockCellProvider";

export type { GridApi, CanvasGridProps, GridAxisSizeChange, GridInteractionMode, ScrollToCellAlign, FillCommitEvent } from "./react/CanvasGrid";
export { CanvasGrid } from "./react/CanvasGrid";

export { wheelDeltaToPixels } from "./react/wheelDeltaToPixels";
export type { WheelDeltaToPixelsOptions } from "./react/wheelDeltaToPixels";

export type { FillMode } from "./interaction/fillHandle";

export type { GridPresence, GridPresenceCursor, GridPresenceRange } from "./presence/types";

export type {
  CanvasGridImageResolver,
  CanvasGridImageSource,
  CanvasGridRendererOptions,
  GridPerfStats,
  GridViewportChangeEvent,
  GridViewportChangeListener,
  GridViewportChangeReason,
  GridViewportSubscriptionOptions
} from "./rendering/CanvasGridRenderer";
export type { GridTheme } from "./theme/GridTheme";
export { DEFAULT_GRID_THEME, resolveGridTheme } from "./theme/GridTheme";
export { DEFAULT_GRID_FONT_FAMILY, DEFAULT_GRID_MONOSPACE_FONT_FAMILY } from "./rendering/defaultFontFamilies";
export {
  GRID_THEME_CSS_VAR_NAMES,
  readGridThemeFromCssVars,
  resolveCssVarValue,
  resolveGridThemeFromCssVars
} from "./theme/resolveThemeFromCssVars";

export { CanvasGridRenderer } from "./rendering/CanvasGridRenderer";
export { DirtyRegionTracker } from "./rendering/DirtyRegionTracker";
export type { Rect } from "./rendering/DirtyRegionTracker";
export { LruCache } from "./utils/LruCache";
export { MAX_GRID_ZOOM, MIN_GRID_ZOOM, clampZoom } from "./utils/zoomMath";

export { VariableSizeAxis } from "./virtualization/VariableSizeAxis";
export type { AxisVisibleRange } from "./virtualization/VariableSizeAxis";
export type { ScrollbarThumb } from "./virtualization/scrollbarMath";
export { computeScrollbarThumb } from "./virtualization/scrollbarMath";
export { alignScrollToDevicePixels } from "./virtualization/alignScrollToDevicePixels";
export { VirtualScrollManager } from "./virtualization/VirtualScrollManager";
export type { GridViewportState } from "./virtualization/VirtualScrollManager";
export { GridPlaceholder } from "./GridPlaceholder";
export type { GridPlaceholderProps } from "./GridPlaceholder";

export {
  SR_ONLY_STYLE,
  applySrOnlyStyle,
  describeActiveCellLabel,
  describeCell,
  describeCellForA11y,
  formatCellDisplayText,
  toA1Address,
  toColumnName
} from "./a11y/a11y";
