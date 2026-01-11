import type { ChartTheme } from "./types";

export const DEFAULT_CHART_THEME: ChartTheme = {
  fonts: {
    title: { family: "sans-serif", sizePx: 14, weight: 600 },
    axis: { family: "sans-serif", sizePx: 11 },
    legend: { family: "sans-serif", sizePx: 11 },
  },
  palette: [
    "var(--chart-series-1)",
    "var(--chart-series-2)",
    "var(--chart-series-3)",
    "var(--chart-series-4)",
  ],
};

export function resolveChartTheme(theme?: Partial<ChartTheme> | null): ChartTheme {
  if (!theme) return DEFAULT_CHART_THEME;
  return {
    fonts: {
      title: { ...DEFAULT_CHART_THEME.fonts.title, ...(theme.fonts?.title ?? {}) },
      axis: { ...DEFAULT_CHART_THEME.fonts.axis, ...(theme.fonts?.axis ?? {}) },
      legend: { ...DEFAULT_CHART_THEME.fonts.legend, ...(theme.fonts?.legend ?? {}) },
    },
    palette: theme.palette ?? DEFAULT_CHART_THEME.palette,
  };
}
