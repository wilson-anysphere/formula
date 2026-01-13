import type { WorkbookThemePalette } from "@formula/workbook-backend";

export type { WorkbookThemePalette };

export type ChartTheme = {
  /**
   * CSS colors to use for series fills/strokes.
   *
   * Excel defaults to the theme's accent1..accent6 cycling.
   */
  seriesColors: string[];
};

export const FALLBACK_CHART_THEME: ChartTheme = {
  seriesColors: [
    "var(--chart-series-1)",
    "var(--chart-series-2)",
    "var(--chart-series-3)",
    "var(--chart-series-4)",
  ],
};

export function chartThemeFromWorkbookPalette(
  palette?: WorkbookThemePalette | null,
): ChartTheme {
  if (!palette) return FALLBACK_CHART_THEME;

  const wrapHighContrastFallback = (index: number, color: string): string =>
    `var(--chart-series-hc-${index}, ${color})`;

  const seriesColors = [
    palette.accent1 ? wrapHighContrastFallback(1, palette.accent1) : null,
    palette.accent2 ? wrapHighContrastFallback(2, palette.accent2) : null,
    palette.accent3 ? wrapHighContrastFallback(3, palette.accent3) : null,
    palette.accent4 ? wrapHighContrastFallback(4, palette.accent4) : null,
    palette.accent5 ? wrapHighContrastFallback(5, palette.accent5) : null,
    palette.accent6 ? wrapHighContrastFallback(6, palette.accent6) : null,
  ].filter((value): value is string => Boolean(value));

  if (seriesColors.length === 0) return FALLBACK_CHART_THEME;
  return { seriesColors };
}
