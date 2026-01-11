import {
  defaultChartTheme as defaultChartThemeImpl,
  renderChartToCanvas as renderChartToCanvasImpl,
  renderChartToSvg as renderChartToSvgImpl,
  resolveChartData as resolveChartDataImpl,
} from "./renderChart.js";

export type ChartKind = "bar" | "line" | "pie" | "scatter" | "unknown";

export type ChartAxisModel = {
  tickCount?: number;
  majorGridlines?: boolean;
};

export type ChartLegendModel = {
  show?: boolean;
  position?: "right" | "bottom";
};

export type ChartSeriesModel = {
  name?: string | null;
  categories?: { ref?: string; strCache?: Array<string | null> };
  values?: { ref?: string; numCache?: Array<number | null> };
  xValues?: { ref?: string; numCache?: Array<number | null> };
  yValues?: { ref?: string; numCache?: Array<number | null> };
};

export type ChartModel = {
  chartType: { kind: ChartKind; name?: string };
  title?: string | null;
  legend?: ChartLegendModel;
  axes?: {
    category?: ChartAxisModel;
    value?: ChartAxisModel;
    x?: ChartAxisModel;
    y?: ChartAxisModel;
  };
  options?: {
    markers?: boolean;
  };
  series: ChartSeriesModel[];
};

export type ResolvedChartSeriesData = {
  name: string | null;
  categories: string[];
  values: number[];
  xValues: number[];
  yValues: number[];
};

export type ResolvedChartData = {
  series: ResolvedChartSeriesData[];
};

export type ChartTheme = {
  background: string;
  border: string;
  axis: string;
  gridline: string;
  title: string;
  label: string;
  fontFamily: string;
  fontSize: number;
  seriesColors: string[];
};

export type SizePx = { width: number; height: number };

export const defaultChartTheme: ChartTheme = defaultChartThemeImpl as unknown as ChartTheme;

export function resolveChartData(model: ChartModel, liveData?: Partial<ResolvedChartData>): ResolvedChartData {
  return resolveChartDataImpl(model, liveData) as unknown as ResolvedChartData;
}

export function renderChartToSvg(model: ChartModel, data: ResolvedChartData, theme: ChartTheme, sizePx: SizePx): string {
  return renderChartToSvgImpl(model, data, theme, sizePx);
}

export function renderChartToCanvas(
  ctx: CanvasRenderingContext2D,
  model: ChartModel,
  data: ResolvedChartData,
  theme: ChartTheme,
  sizePx: SizePx
): void {
  renderChartToCanvasImpl(ctx, model, data, theme, sizePx);
}

