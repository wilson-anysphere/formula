import type {
  AxisLayout,
  BandScale,
  ChartAxisModel,
  ChartLayout,
  ChartModel,
  ChartSeriesModel,
  ChartTheme,
  LegendEntryLayout,
  LegendLayout,
  LinearScale,
  Rect,
  Scale,
  TickLayout,
  ViewportRect,
} from "./types";
import { rectBottom, rectRight, round } from "./geometry";
import { resolveChartTheme } from "./theme";
import { extractSeriesNumbers, extractSeriesStrings } from "./data";
import { estimateLineHeight, estimateTextWidth, layoutTextBlock, wrapTextToWidth } from "./text";
import { formatTickValue, generateLinearTicks } from "./ticks";

const OUTER_PADDING_PX = 8;
const TITLE_PADDING_Y_PX = 4;
const TITLE_GAP_PX = 4;

const LEGEND_GAP_PX = 8;
const LEGEND_PADDING_PX = 8;
const LEGEND_MARKER_SIZE_PX = 10;
const LEGEND_MARKER_GAP_PX = 6;
const LEGEND_ENTRY_GAP_PX = 4;

const AXIS_TICK_MARK_PX = 4;
const AXIS_LABEL_GAP_PX = 4;
const PLOT_INSET_TOP_PX = 8;
const PLOT_INSET_RIGHT_PX = 8;

export function computeChartLayout(
  model: ChartModel,
  theme: Partial<ChartTheme> | null | undefined,
  viewport: ViewportRect
): ChartLayout;
export function computeChartLayout(args: {
  model: ChartModel;
  theme?: Partial<ChartTheme> | null;
  viewport: ViewportRect;
}): ChartLayout;
export function computeChartLayout(
  modelOrArgs: ChartModel | { model: ChartModel; theme?: Partial<ChartTheme> | null; viewport: ViewportRect },
  themeMaybe?: Partial<ChartTheme> | null,
  viewportMaybe?: ViewportRect
): ChartLayout {
  const args =
    typeof modelOrArgs === "object" && modelOrArgs !== null && "model" in modelOrArgs && "viewport" in modelOrArgs
      ? modelOrArgs
      : { model: modelOrArgs as ChartModel, theme: themeMaybe, viewport: viewportMaybe as ViewportRect };

  if (!args.viewport) {
    throw new Error("computeChartLayout: viewport is required");
  }

  const theme = resolveChartTheme(args.theme);
  const model = args.model;

  const chartAreaRect: Rect = {
    x: round(args.viewport.x),
    y: round(args.viewport.y),
    width: Math.max(0, round(args.viewport.width)),
    height: Math.max(0, round(args.viewport.height)),
  };

  const titleTextRaw = (model.title ?? "").trim();
  let titleRect: Rect | null = null;
  let titleText: ChartLayout["titleText"] = null;

  let contentTop = chartAreaRect.y + OUTER_PADDING_PX;
  const contentLeft = chartAreaRect.x + OUTER_PADDING_PX;
  const contentRight = rectRight(chartAreaRect) - OUTER_PADDING_PX;
  const contentBottom = rectBottom(chartAreaRect) - OUTER_PADDING_PX;

  if (titleTextRaw) {
    const maxWidth = Math.max(0, contentRight - contentLeft);
    const lines = wrapTextToWidth(titleTextRaw, theme.fonts.title, maxWidth);
    const lineHeight = estimateLineHeight(theme.fonts.title);
    const titleHeight = lines.length * lineHeight + TITLE_PADDING_Y_PX * 2;

    titleRect = {
      x: contentLeft,
      y: contentTop,
      width: maxWidth,
      height: round(titleHeight),
    };

    titleText = layoutTextBlock({
      rect: titleRect,
      font: theme.fonts.title,
      align: "center",
      lines,
      paddingY: TITLE_PADDING_Y_PX,
    });

    contentTop = rectBottom(titleRect) + TITLE_GAP_PX;
  }

  const contentRect: Rect = {
    x: contentLeft,
    y: contentTop,
    width: Math.max(0, contentRight - contentLeft),
    height: Math.max(0, contentBottom - contentTop),
  };

  const legendPref = normalizeLegendPosition(model.legend?.position);
  const legendOverlay = model.legend?.overlay ?? false;

  let legendRect: Rect | null = null;
  let legend: LegendLayout | null = null;
  let plotContainerRect: Rect = contentRect;

  if (legendPref === "right") {
    const entriesText = model.series.map((s, i) => (s.name ?? "").trim() || `Series ${i + 1}`);
    const legendFont = theme.fonts.legend;
    const legendLineHeight = estimateLineHeight(legendFont);
    let maxLabelWidth = 0;
    for (const label of entriesText) {
      const w = estimateTextWidth(label, legendFont);
      if (w > maxLabelWidth) maxLabelWidth = w;
    }

    // A simple width heuristic that grows with label length but caps to 40% of the
    // available content width.
    const requiredWidth =
      LEGEND_PADDING_PX * 2 + LEGEND_MARKER_SIZE_PX + LEGEND_MARKER_GAP_PX + maxLabelWidth;
    const maxWidth = contentRect.width * 0.4;
    const width = Math.max(0, Math.min(requiredWidth, maxWidth));

    const entriesHeight =
      entriesText.length * legendLineHeight +
      Math.max(0, entriesText.length - 1) * LEGEND_ENTRY_GAP_PX;
    const desiredHeight = entriesHeight + LEGEND_PADDING_PX * 2;
    const height = Math.min(contentRect.height, desiredHeight);
    const legendY = contentRect.y + Math.max(0, (contentRect.height - height) / 2);

    legendRect = {
      x: round(rectRight(contentRect) - width),
      y: round(legendY),
      width: round(width),
      height: round(height),
    };

    const entries: LegendEntryLayout[] = [];
    let y = legendRect.y + LEGEND_PADDING_PX;
    for (let i = 0; i < entriesText.length; i += 1) {
      const label = entriesText[i];
      const color = theme.palette[i % theme.palette.length] ?? theme.palette[0] ?? "currentColor";
      const markerRect: Rect = {
        x: legendRect.x + LEGEND_PADDING_PX,
        y: round(y + (legendLineHeight - LEGEND_MARKER_SIZE_PX) / 2),
        width: LEGEND_MARKER_SIZE_PX,
        height: LEGEND_MARKER_SIZE_PX,
      };
      const labelRect: Rect = {
        x: markerRect.x + markerRect.width + LEGEND_MARKER_GAP_PX,
        y: round(y),
        width: Math.max(
          0,
          rectRight(legendRect) - LEGEND_PADDING_PX - (markerRect.x + markerRect.width + LEGEND_MARKER_GAP_PX)
        ),
        height: round(legendLineHeight),
      };

      entries.push({ seriesIndex: i, label, color, markerRect, labelRect });
      y += legendLineHeight + LEGEND_ENTRY_GAP_PX;
    }

    legend = { rect: legendRect, font: legendFont, entries };

    if (!legendOverlay) {
      plotContainerRect = {
        x: contentRect.x,
        y: contentRect.y,
        width: Math.max(0, contentRect.width - width - LEGEND_GAP_PX),
        height: contentRect.height,
      };
    }
  }

  if (model.chartType.kind === "pie") {
    const inset = PLOT_INSET_TOP_PX;
    const plotAreaRect: Rect = {
      x: round(plotContainerRect.x + inset),
      y: round(plotContainerRect.y + inset),
      width: Math.max(1, round(plotContainerRect.width - inset * 2)),
      height: Math.max(1, round(plotContainerRect.height - inset * 2)),
    };

    return {
      chartAreaRect,
      plotAreaRect,
      titleRect,
      titleText,
      legendRect,
      legend,
      axes: {},
      scales: {},
    };
  }

  const { xAxis, yAxis } = resolvePrimaryAxes(model);

  const axisFont = theme.fonts.axis;
  const axisLineHeight = estimateLineHeight(axisFont);

  const yValues = collectNumericValues(model, yAxis, "y");
  const yAxisInfo = computeValueAxisInfo({
    values: yValues,
    axis: yAxis,
    includeZero: model.chartType.kind !== "scatter",
  });
  const yTickLabels = yAxisInfo.ticks.map((v) => formatTickValue(v, yAxisInfo.formatCode));
  let maxYTickLabelWidth = 0;
  for (const lbl of yTickLabels) {
    const w = estimateTextWidth(lbl, axisFont);
    if (w > maxYTickLabelWidth) maxYTickLabelWidth = w;
  }

  const axisMarginLeft = Math.max(0, AXIS_TICK_MARK_PX + AXIS_LABEL_GAP_PX + maxYTickLabelWidth);
  const axisMarginBottom = Math.max(0, AXIS_TICK_MARK_PX + AXIS_LABEL_GAP_PX + axisLineHeight);

  const plotAreaRect: Rect = {
    x: round(plotContainerRect.x + axisMarginLeft),
    y: round(plotContainerRect.y + PLOT_INSET_TOP_PX),
    width: Math.max(1, round(plotContainerRect.width - axisMarginLeft - PLOT_INSET_RIGHT_PX)),
    height: Math.max(1, round(plotContainerRect.height - PLOT_INSET_TOP_PX - axisMarginBottom)),
  };

  const xAxisLayout = buildXAxisLayout({
    model,
    axis: xAxis,
    axisFont,
    plotAreaRect,
  });

  const yAxisLayout = buildYAxisLayout({
    axis: yAxis,
    axisFont,
    plotAreaRect,
    info: yAxisInfo,
  });

  const scales: Record<string, Scale> = {};
  const axes: Record<string, AxisLayout> = {};

  scales[xAxisLayout.id] = xAxisLayout.scale;
  scales[yAxisLayout.id] = yAxisLayout.scale;
  axes[xAxisLayout.id] = xAxisLayout.axis;
  axes[yAxisLayout.id] = yAxisLayout.axis;

  return {
    chartAreaRect,
    plotAreaRect,
    titleRect,
    titleText,
    legendRect,
    legend,
    axes,
    scales,
  };
}

function resolvePrimaryAxes(model: ChartModel): { xAxis: ChartAxisModel; yAxis: ChartAxisModel } {
  const axes = model.axes ?? [];

  const defaultX: ChartAxisModel =
    model.chartType.kind === "scatter"
      ? { kind: "value", position: "bottom" }
      : { kind: "category", position: "bottom" };

  const defaultY: ChartAxisModel = { kind: "value", position: "left" };

  const xAxis =
    axes.find(
      (a) =>
        normalizeAxisPosition(a.position) === "bottom" &&
        (model.chartType.kind === "scatter" ? normalizeAxisKind(a.kind) === "value" : true)
    ) ??
    axes.find((a) => normalizeAxisKind(a.kind) === normalizeAxisKind(defaultX.kind)) ??
    defaultX;

  const yAxis =
    axes.find((a) => normalizeAxisPosition(a.position) === "left" && normalizeAxisKind(a.kind) === "value") ??
    axes.find((a) => normalizeAxisKind(a.kind) === "value") ??
    defaultY;

  return { xAxis, yAxis };
}

function normalizeLegendPosition(pos: unknown): "right" | "none" {
  if (pos === "right" || pos === "r") return "right";
  return "none";
}

function normalizeAxisPosition(pos: unknown): "left" | "right" | "top" | "bottom" | null {
  if (pos === "left" || pos === "l") return "left";
  if (pos === "right" || pos === "r") return "right";
  if (pos === "top" || pos === "t") return "top";
  if (pos === "bottom" || pos === "b") return "bottom";
  return null;
}

function normalizeAxisKind(kind: unknown): "category" | "value" | null {
  if (kind === "category" || kind === "catAx") return "category";
  if (kind === "value" || kind === "valAx") return "value";
  return null;
}

function isReverseOrder(scaling: unknown): boolean {
  if (!scaling || typeof scaling !== "object") return false;
  const s: any = scaling;
  if (s.orientation === "maxMin") return true;
  if (s.reverseOrder === true) return true;
  if (s.reverseOrder === 1 || s.reverseOrder === "1") return true;
  return false;
}

function axisFormatCode(axis: ChartAxisModel): string | null | undefined {
  return axis.numberFormatCode ?? axis.formatCode;
}

function collectNumericValues(model: ChartModel, axis: ChartAxisModel, role: "x" | "y"): number[] {
  const out: number[] = [];
  if (model.chartType.kind === "scatter") {
    for (const s of model.series) {
      const key = role === "x" ? "xValues" : "yValues";
      const vals = extractSeriesNumbers(s, key);
      if (role === "x" && vals.length === 0) {
        // Some producers (including our small fixtures) may serialize x values as strings.
        // Excel treats non-numeric x values as sequential indices; mirror that so the
        // layout engine can still derive a stable numeric domain for the value axis.
        const len = Math.max(cacheLength(s.xValues), cacheLength(s.yValues));
        if (len > 0) {
          for (let i = 0; i < len; i += 1) out.push(i + 1);
        }
        continue;
      }
      for (const v of vals) out.push(v);
    }
    return out;
  }

  // Category charts only have a numeric value axis (y).
  if (role === "y") {
    for (const s of model.series) {
      const vals = extractSeriesNumbers(s, "values");
      for (const v of vals) out.push(v);
    }
  }

  return out;
}

function cacheLength(data: unknown): number {
  if (!data) return 0;
  if (Array.isArray(data)) return data.length;
  if (typeof data === "object" && data !== null && "cache" in data && Array.isArray((data as any).cache)) {
    return (data as any).cache.length;
  }
  return 0;
}

function computeValueAxisInfo(args: {
  values: number[];
  axis: ChartAxisModel;
  includeZero: boolean;
}): { domain: [number, number]; ticks: number[]; formatCode?: string | null; reverseOrder: boolean } {
  const values = args.values.filter(Number.isFinite);
  const scaling = args.axis.scaling ?? {};
  const minExplicit = typeof scaling.min === "number" && Number.isFinite(scaling.min);
  const maxExplicit = typeof scaling.max === "number" && Number.isFinite(scaling.max);

  let min = minExplicit ? (scaling.min as number) : Infinity;
  let max = maxExplicit ? (scaling.max as number) : -Infinity;
  if (!minExplicit || !maxExplicit) {
    for (const v of values) {
      if (!minExplicit && v < min) min = v;
      if (!maxExplicit && v > max) max = v;
    }
  }

  // Preserve explicit bounds even if the data cache is empty.
  if (!Number.isFinite(min) && Number.isFinite(max)) min = max - 1;
  if (!Number.isFinite(max) && Number.isFinite(min)) max = min + 1;
  if (!Number.isFinite(min) && !Number.isFinite(max)) {
    min = 0;
    max = 1;
  }

  if (args.includeZero && !minExplicit && min > 0) min = 0;
  if (args.includeZero && !maxExplicit && max < 0) max = 0;

  const tickResult = generateLinearTicks({ domain: [min, max], minExplicit, maxExplicit });
  return {
    domain: tickResult.domain,
    ticks: tickResult.ticks,
    formatCode: axisFormatCode(args.axis),
    reverseOrder: isReverseOrder(scaling),
  };
}

function scaleLinear(v: number, scale: LinearScale): number {
  const [d0, d1] = scale.domain;
  const [r0, r1] = scale.range;
  if (d0 === d1) return r0;
  const t = (v - d0) / (d1 - d0);
  return r0 + t * (r1 - r0);
}

function buildYAxisLayout(args: {
  axis: ChartAxisModel;
  axisFont: ChartTheme["fonts"]["axis"];
  plotAreaRect: Rect;
  info: { domain: [number, number]; ticks: number[]; formatCode?: string | null; reverseOrder: boolean };
}): { id: string; axis: AxisLayout; scale: LinearScale } {
  const plotTop = args.plotAreaRect.y;
  const plotBottom = rectBottom(args.plotAreaRect);
  const plotLeft = args.plotAreaRect.x;
  const plotRight = rectRight(args.plotAreaRect);

  const range: [number, number] = args.info.reverseOrder ? [plotTop, plotBottom] : [plotBottom, plotTop];
  const scale: LinearScale = { type: "linear", domain: args.info.domain, range: [round(range[0]), round(range[1])] };

  const axisLine = {
    x1: plotLeft,
    y1: plotTop,
    x2: plotLeft,
    y2: plotBottom,
  };

  const labelHeight = estimateLineHeight(args.axisFont);
  const ticks: TickLayout[] = [];
  const gridlines: AxisLayout["gridlines"] = [];
  const wantsGridlines = Boolean(args.axis.majorGridlines);

  for (const tickValue of args.info.ticks) {
    const y = round(scaleLinear(tickValue, scale));
    const label = formatTickValue(tickValue, args.info.formatCode);
    const labelWidth = estimateTextWidth(label, args.axisFont);
    const labelRect: Rect = {
      x: round(plotLeft - AXIS_TICK_MARK_PX - AXIS_LABEL_GAP_PX - labelWidth),
      y: round(y - labelHeight / 2),
      width: round(labelWidth),
      height: round(labelHeight),
    };
    ticks.push({
      value: tickValue,
      label,
      position: { x: plotLeft, y },
      labelRect,
    });

    if (wantsGridlines) {
      gridlines.push({ x1: plotLeft, y1: y, x2: plotRight, y2: y });
    }
  }

  return {
    id: "y",
    scale,
    axis: {
      id: "y",
      orientation: "y",
      kind: "value",
      axisLine,
      ticks,
      gridlines,
    },
  };
}

function buildXAxisLayout(args: {
  model: ChartModel;
  axis: ChartAxisModel;
  axisFont: ChartTheme["fonts"]["axis"];
  plotAreaRect: Rect;
}): { id: string; axis: AxisLayout; scale: Scale } {
  const plotLeft = args.plotAreaRect.x;
  const plotRight = rectRight(args.plotAreaRect);
  const plotBottom = rectBottom(args.plotAreaRect);
  const axisY = plotBottom;

  const axisLine = { x1: plotLeft, y1: axisY, x2: plotRight, y2: axisY };

  if (args.model.chartType.kind === "scatter") {
    const xValues = collectNumericValues(args.model, args.axis, "x");
    const info = computeValueAxisInfo({
      values: xValues,
      axis: args.axis,
      includeZero: false,
    });

    const range: [number, number] = info.reverseOrder ? [plotRight, plotLeft] : [plotLeft, plotRight];
    const scale: LinearScale = { type: "linear", domain: info.domain, range: [round(range[0]), round(range[1])] };

    const labelHeight = estimateLineHeight(args.axisFont);
    const ticks: TickLayout[] = [];

    for (const tickValue of info.ticks) {
      const x = round(scaleLinear(tickValue, scale));
      const label = formatTickValue(tickValue, info.formatCode);
      const labelWidth = estimateTextWidth(label, args.axisFont);
      const labelRect: Rect = {
        x: round(x - labelWidth / 2),
        y: round(axisY + AXIS_TICK_MARK_PX + AXIS_LABEL_GAP_PX),
        width: round(labelWidth),
        height: round(labelHeight),
      };
      ticks.push({
        value: tickValue,
        label,
        position: { x, y: axisY },
        labelRect,
      });
    }

    return {
      id: "x",
      scale,
      axis: {
        id: "x",
        orientation: "x",
        kind: "value",
        axisLine,
        ticks,
        gridlines: [],
      },
    };
  }

  const categories = resolveCategories(args.model.series);
  const reverse = isReverseOrder(args.axis.scaling);
  const domain = reverse ? [...categories].reverse() : categories;
  const count = Math.max(1, domain.length);
  const step = args.plotAreaRect.width / count;
  const scale: BandScale = {
    type: "band",
    domain,
    range: [plotLeft, plotRight],
    step: round(step),
    bandwidth: round(step * 0.8),
  };

  const labelHeight = estimateLineHeight(args.axisFont);
  const ticks: TickLayout[] = [];

  for (let i = 0; i < domain.length; i += 1) {
    const label = domain[i];
    const x = round(plotLeft + (i + 0.5) * step);
    const labelWidth = estimateTextWidth(label, args.axisFont);
    const labelRect: Rect = {
      x: round(x - labelWidth / 2),
      y: round(axisY + AXIS_TICK_MARK_PX + AXIS_LABEL_GAP_PX),
      width: round(labelWidth),
      height: round(labelHeight),
    };
    ticks.push({ value: label, label, position: { x, y: axisY }, labelRect });
  }

  return {
    id: "x",
    scale,
    axis: {
      id: "x",
      orientation: "x",
      kind: "category",
      axisLine,
      ticks,
      gridlines: [],
    },
  };
}

function resolveCategories(series: ChartSeriesModel[]): string[] {
  const first = series[0];
  if (!first) return [];
  const cats = extractSeriesStrings(first, "categories").filter((s) => s.trim().length > 0);
  if (cats.length) return cats;

  let valuesLen = 0;
  for (const s of series) {
    const len = extractSeriesNumbers(s, "values").length;
    if (len > valuesLen) valuesLen = len;
  }
  return Array.from({ length: valuesLen }, (_, i) => String(i + 1));
}
