export type WorkbookThemePalette = {
  dk1: string;
  lt1: string;
  dk2: string;
  lt2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  hlink: string;
  followedHlink: string;
};

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

  const seriesColors = [
    palette.accent1,
    palette.accent2,
    palette.accent3,
    palette.accent4,
    palette.accent5,
    palette.accent6,
  ].filter(Boolean);

  if (seriesColors.length === 0) return FALLBACK_CHART_THEME;
  return { seriesColors };
}

