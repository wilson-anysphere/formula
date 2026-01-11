import type { GridTheme } from "./GridTheme";

export const GRID_THEME_CSS_VAR_NAMES: Record<keyof GridTheme, string> = {
  gridBg: "--formula-grid-bg",
  gridLine: "--formula-grid-line",
  headerBg: "--formula-grid-header-bg",
  headerText: "--formula-grid-header-text",
  cellText: "--formula-grid-cell-text",
  errorText: "--formula-grid-error-text",
  selectionFill: "--formula-grid-selection-fill",
  selectionBorder: "--formula-grid-selection-border",
  selectionHandle: "--formula-grid-selection-handle",
  scrollbarTrack: "--formula-grid-scrollbar-track",
  scrollbarThumb: "--formula-grid-scrollbar-thumb",
  freezeLine: "--formula-grid-freeze-line",
  commentIndicator: "--formula-grid-comment-indicator",
  commentIndicatorResolved: "--formula-grid-comment-indicator-resolved",
  remotePresenceDefault: "--formula-grid-remote-presence-default"
};

export function readGridThemeFromCssVars(style: { getPropertyValue: (name: string) => string }): Partial<GridTheme> {
  const partial: Partial<GridTheme> = {};

  for (const [key, cssVar] of Object.entries(GRID_THEME_CSS_VAR_NAMES) as Array<[keyof GridTheme, string]>) {
    const value = style.getPropertyValue(cssVar);
    if (!value) continue;
    const trimmed = value.trim();
    if (!trimmed) continue;
    partial[key] = trimmed;
  }

  return partial;
}

