import { computeChartLayout, type ChartLayout, type ChartModel as LayoutChartModel, type ChartTheme as LayoutChartTheme, type LinearScale } from "./layout/index.js";
import { renderSceneToCanvas, renderSceneToSvg, path, type FontSpec as SceneFontSpec, type Node, type Paint, type Scene, type Stroke, type TextAlign, type TextBaseline } from "./scene/index.js";
import { resolveCssVar } from "../theme/cssVars.js";

export type ChartModel = LayoutChartModel & {
  /**
   * Renderer-only options (not part of the OOXML model yet).
   */
  options?: {
    markers?: boolean;
    /**
     * When provided, overrides the default placeholder text rendered when the chart cannot
     * be drawn (e.g. unsupported kind, empty data, or host-side guards like range limits).
     */
    placeholder?: string;
  };
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

export const defaultChartTheme: ChartTheme = {
  background: "var(--chart-bg)",
  border: "var(--chart-border)",
  axis: "var(--chart-axis)",
  gridline: "var(--chart-border)",
  title: "var(--chart-title)",
  label: "var(--chart-label)",
  fontFamily: "sans-serif",
  fontSize: 11,
  seriesColors: [
    "var(--chart-series-1)",
    "var(--chart-series-2)",
    "var(--chart-series-3)",
    "var(--chart-series-4)",
  ],
};

function resolveTheme(theme: ChartTheme | null | undefined): ChartTheme {
  if (!theme) return defaultChartTheme;
  return {
    ...defaultChartTheme,
    ...theme,
    seriesColors: Array.isArray(theme.seriesColors) && theme.seriesColors.length ? theme.seriesColors : defaultChartTheme.seriesColors,
  };
}

function resolveCssColor(value: string): string {
  const trimmed = value.trim();
  const varMatch = /^var\(\s*(--[^,\s)]+)(?:\s*,[^)]+)?\s*\)$/.exec(trimmed);
  if (varMatch) {
    return resolveCssVar(varMatch[1], { fallback: trimmed });
  }
  if (trimmed.startsWith("--")) {
    return resolveCssVar(trimmed, { fallback: trimmed });
  }
  return trimmed;
}

function resolveThemeForCanvas(theme: ChartTheme): ChartTheme {
  return {
    ...theme,
    background: resolveCssColor(theme.background),
    border: resolveCssColor(theme.border),
    axis: resolveCssColor(theme.axis),
    gridline: resolveCssColor(theme.gridline),
    title: resolveCssColor(theme.title),
    label: resolveCssColor(theme.label),
    seriesColors: theme.seriesColors.map(resolveCssColor),
  };
}

function toPaint(color: string): Paint {
  return { color };
}

function toStroke(color: string, width: number, dash?: number[]): Stroke {
  return { paint: { color }, width, ...(dash ? { dash } : {}) };
}

function toSceneFont(fontFamily: string, sizePx: number, weight?: SceneFontSpec["weight"]): SceneFontSpec {
  return { family: fontFamily, sizePx, ...(weight != null ? { weight } : {}) };
}

function scaleLinear(v: number, scale: LinearScale): number {
  const [d0, d1] = scale.domain;
  const [r0, r1] = scale.range;
  if (d0 === d1) return r0;
  const t = (v - d0) / (d1 - d0);
  return r0 + t * (r1 - r0);
}

function isFiniteNumber(n: unknown): n is number {
  return typeof n === "number" && Number.isFinite(n);
}

type CacheLike<T> =
  | Array<T | null>
  | { cache?: Array<T | null> | null }
  | { strCache?: Array<T | null> | null }
  | { numCache?: Array<T | null> | null }
  | null
  | undefined;

function extractCache<T>(data: CacheLike<T>): Array<T | null> {
  if (!data) return [];
  if (Array.isArray(data)) return data;
  if (typeof data === "object") {
    const maybe: any = data;
    if (Array.isArray(maybe.cache)) return maybe.cache;
    if (Array.isArray(maybe.strCache)) return maybe.strCache;
    if (Array.isArray(maybe.numCache)) return maybe.numCache;
  }
  return [];
}

function extractStringCache(data: CacheLike<string | number>): string[] {
  const values = extractCache(data);
  return values.map((v) => (v == null ? "" : String(v)));
}

function extractNumberCache(data: CacheLike<number | string>): number[] {
  const values = extractCache(data);
  return values.map((v) => {
    const n = typeof v === "number" ? v : Number(v);
    return Number.isFinite(n) ? n : Number.NaN;
  });
}

export function resolveChartData(model: ChartModel, liveData?: Partial<ResolvedChartData>): ResolvedChartData {
  const modelSeries = Array.isArray(model?.series) ? model.series : [];
  const liveSeries = Array.isArray(liveData?.series) ? liveData!.series! : [];
  const out: ResolvedChartSeriesData[] = [];

  const count = Math.max(modelSeries.length, liveSeries.length);
  for (let i = 0; i < count; i += 1) {
    const m: any = modelSeries[i] ?? {};
    const l: any = liveSeries[i] ?? {};

    const name = (m.name ?? l.name ?? null) as string | null;

    const modelCats = extractStringCache(m.categories);
    // Rust chart models can represent numeric/date categories separately
    // (`categoriesNum`). The UI model converter maps this to `categories`, but
    // keep the renderer resilient if a raw model sneaks through.
    const modelCatsNum = extractStringCache(m.categoriesNum ?? m.categories_num);
    const modelVals = extractNumberCache(m.values);
    const modelX = extractNumberCache(m.xValues);
    const modelY = extractNumberCache(m.yValues);

    const cats = modelCats.length
      ? modelCats
      : modelCatsNum.length
        ? modelCatsNum
        : Array.isArray(l.categories)
          ? l.categories.map((v: any) => String(v ?? ""))
          : [];
    const vals = modelVals.length ? modelVals : Array.isArray(l.values) ? l.values.map((v: any) => Number(v)) : [];
    const xs = modelX.length ? modelX : Array.isArray(l.xValues) ? l.xValues.map((v: any) => Number(v)) : [];
    const ys = modelY.length ? modelY : Array.isArray(l.yValues) ? l.yValues.map((v: any) => Number(v)) : [];

    out.push({
      name: name == null ? null : String(name),
      categories: cats,
      values: vals,
      xValues: xs,
      yValues: ys,
    });
  }

  return { series: out };
}

function applyResolvedDataToModel(model: ChartModel, data: ResolvedChartData): ChartModel {
  const series = Array.isArray(model.series) ? model.series : [];
  return {
    ...model,
    series: series.map((s: any, idx: number) => {
      const resolved = data.series[idx];
      if (!resolved) return s;
      return {
        ...s,
        ...(resolved.name != null ? { name: resolved.name } : {}),
        ...(resolved.categories.length ? { categories: { cache: resolved.categories } } : {}),
        ...(resolved.values.length ? { values: { cache: resolved.values } } : {}),
        ...(resolved.xValues.length ? { xValues: { cache: resolved.xValues } } : {}),
        ...(resolved.yValues.length ? { yValues: { cache: resolved.yValues } } : {}),
      };
    }),
  };
}

function buildLayoutTheme(theme: ChartTheme): Partial<LayoutChartTheme> {
  return {
    fonts: {
      title: { family: theme.fontFamily, sizePx: theme.fontSize + 3, weight: 600 },
      axis: { family: theme.fontFamily, sizePx: theme.fontSize },
      legend: { family: theme.fontFamily, sizePx: theme.fontSize },
    },
  };
}

function buildTitleNodes(layout: ChartLayout, theme: ChartTheme): Node[] {
  if (!layout.titleText) return [];
  const font = toSceneFont(layout.titleText.font.family, layout.titleText.font.sizePx, layout.titleText.font.weight as SceneFontSpec["weight"]);
  const nodes: Node[] = [];
  for (const line of layout.titleText.lines) {
    let align: TextAlign = "start";
    let x = line.x;
    if (layout.titleText.align === "center") {
      align = "center";
      x = line.x + line.width / 2;
    } else if (layout.titleText.align === "end") {
      align = "end";
      x = line.x + line.width;
    }

    nodes.push({
      kind: "text",
      x,
      y: line.y,
      text: line.text,
      font,
      fill: toPaint(theme.title),
      align,
      baseline: "top",
      maxWidth: line.width,
    });
  }
  return nodes;
}

function buildLegendNodes(model: ChartModel, layout: ChartLayout, theme: ChartTheme): Node[] {
  if (model.chartType.kind === "pie") return [];
  if (!layout.legend) return [];
  const font = toSceneFont(layout.legend.font.family, layout.legend.font.sizePx, layout.legend.font.weight as SceneFontSpec["weight"]);

  const nodes: Node[] = [];
  for (const entry of layout.legend.entries) {
    const color = theme.seriesColors[entry.seriesIndex % theme.seriesColors.length] ?? theme.axis;
    nodes.push({
      kind: "rect",
      x: entry.markerRect.x,
      y: entry.markerRect.y,
      width: entry.markerRect.width,
      height: entry.markerRect.height,
      fill: toPaint(color),
      stroke: { paint: toPaint(theme.border), width: 0.5 },
    });
    nodes.push({
      kind: "text",
      x: entry.labelRect.x,
      y: entry.labelRect.y,
      text: entry.label,
      font,
      fill: toPaint(theme.label),
      align: "start",
      baseline: "top",
      maxWidth: entry.labelRect.width,
    });
  }
  return nodes;
}

function buildGridlineNodes(layout: ChartLayout, theme: ChartTheme): Node[] {
  const nodes: Node[] = [];
  for (const axis of Object.values(layout.axes)) {
    for (const grid of axis.gridlines) {
      nodes.push({
        kind: "line",
        x1: grid.x1,
        y1: grid.y1,
        x2: grid.x2,
        y2: grid.y2,
        stroke: toStroke(theme.gridline, 1, [2, 2]),
      });
    }
  }
  return nodes;
}

function buildAxisAndTickNodes(layout: ChartLayout, theme: ChartTheme): Node[] {
  const nodes: Node[] = [];
  const axisFont = toSceneFont(theme.fontFamily, theme.fontSize);
  const tickMarkLen = 4;

  for (const axis of Object.values(layout.axes)) {
    nodes.push({
      kind: "line",
      x1: axis.axisLine.x1,
      y1: axis.axisLine.y1,
      x2: axis.axisLine.x2,
      y2: axis.axisLine.y2,
      stroke: toStroke(theme.axis, 1),
    });

    for (const tick of axis.ticks) {
      if (axis.orientation === "y") {
        nodes.push({
          kind: "line",
          x1: tick.position.x,
          y1: tick.position.y,
          x2: tick.position.x - tickMarkLen,
          y2: tick.position.y,
          stroke: toStroke(theme.axis, 1),
        });

        nodes.push({
          kind: "text",
          x: tick.labelRect.x + tick.labelRect.width,
          y: tick.labelRect.y,
          text: tick.label,
          font: axisFont,
          fill: toPaint(theme.label),
          align: "end",
          baseline: "top",
          maxWidth: tick.labelRect.width,
        });
      } else {
        nodes.push({
          kind: "line",
          x1: tick.position.x,
          y1: tick.position.y,
          x2: tick.position.x,
          y2: tick.position.y + tickMarkLen,
          stroke: toStroke(theme.axis, 1),
        });

        nodes.push({
          kind: "text",
          x: tick.labelRect.x + tick.labelRect.width / 2,
          y: tick.labelRect.y,
          text: tick.label,
          font: axisFont,
          fill: toPaint(theme.label),
          align: "center",
          baseline: "top",
          maxWidth: tick.labelRect.width,
        });
      }
    }
  }

  return nodes;
}

function buildBarNodes(data: ResolvedChartData, layout: ChartLayout, theme: ChartTheme): Node[] {
  const xScale = layout.scales.x;
  const yScale = layout.scales.y;
  if (xScale?.type !== "band" || yScale?.type !== "linear") return [];

  const seriesCount = Math.max(1, data.series.length);
  const catCount = xScale.domain.length;
  if (catCount === 0) return [];

  const barW = xScale.bandwidth / seriesCount;
  const groupOffset = (xScale.step - xScale.bandwidth) / 2;

  const zeroY = scaleLinear(0, yScale);
  const nodes: Node[] = [];

  for (let ci = 0; ci < catCount; ci += 1) {
    const groupX = xScale.range[0] + ci * xScale.step + groupOffset;
    for (let si = 0; si < seriesCount; si += 1) {
      const v = data.series[si]?.values?.[ci];
      if (!isFiniteNumber(v)) continue;
      const y = scaleLinear(v, yScale);
      const top = Math.min(zeroY, y);
      const h = Math.abs(zeroY - y);
      nodes.push({
        kind: "rect",
        x: groupX + si * barW,
        y: top,
        width: barW,
        height: h,
        fill: toPaint(theme.seriesColors[si % theme.seriesColors.length] ?? theme.axis),
      });
    }
  }

  return nodes;
}

function buildLineNodes(model: ChartModel, data: ResolvedChartData, layout: ChartLayout, theme: ChartTheme): Node[] {
  const xScale = layout.scales.x;
  const yScale = layout.scales.y;
  if (xScale?.type !== "band" || yScale?.type !== "linear") return [];
  const catCount = xScale.domain.length;
  const seriesCount = data.series.length;

  const groupOffset = (xScale.step - xScale.bandwidth) / 2;
  const nodes: Node[] = [];

  for (let si = 0; si < seriesCount; si += 1) {
    const points: Array<{ x: number; y: number }> = [];
    for (let ci = 0; ci < catCount; ci += 1) {
      const v = data.series[si]?.values?.[ci];
      if (!isFiniteNumber(v)) continue;
      const cx = xScale.range[0] + ci * xScale.step + groupOffset + xScale.bandwidth / 2;
      points.push({ x: cx, y: scaleLinear(v, yScale) });
    }

    const color = theme.seriesColors[si % theme.seriesColors.length] ?? theme.axis;
    nodes.push({
      kind: "polyline",
      points,
      fill: { color: "transparent" },
      stroke: toStroke(color, 2),
    });

    if (model.options?.markers) {
      for (const p of points) {
        nodes.push({
          kind: "circle",
          cx: p.x,
          cy: p.y,
          r: 3,
          fill: toPaint(color),
        });
      }
    }
  }

  return nodes;
}

function buildScatterNodes(data: ResolvedChartData, layout: ChartLayout, theme: ChartTheme): Node[] {
  const xScale = layout.scales.x;
  const yScale = layout.scales.y;
  if (xScale?.type !== "linear" || yScale?.type !== "linear") return [];

  const nodes: Node[] = [];
  for (let si = 0; si < data.series.length; si += 1) {
    const color = theme.seriesColors[si % theme.seriesColors.length] ?? theme.axis;
    const xs = data.series[si]?.xValues ?? [];
    const ys = data.series[si]?.yValues ?? [];
    for (let i = 0; i < Math.min(xs.length, ys.length); i += 1) {
      const x = xs[i];
      const y = ys[i];
      if (!isFiniteNumber(x) || !isFiniteNumber(y)) continue;
      nodes.push({
        kind: "circle",
        cx: scaleLinear(x, xScale),
        cy: scaleLinear(y, yScale),
        r: 3,
        fill: toPaint(color),
      });
    }
  }
  return nodes;
}

function buildPieNodes(data: ResolvedChartData, layout: ChartLayout, theme: ChartTheme): Node[] {
  const values = data.series[0]?.values?.filter(Number.isFinite) ?? [];
  const total = values.reduce((a, b) => a + b, 0);
  if (!(total > 0)) return [];

  const cx = layout.plotAreaRect.x + layout.plotAreaRect.width / 2;
  const cy = layout.plotAreaRect.y + layout.plotAreaRect.height / 2;
  const r = Math.min(layout.plotAreaRect.width, layout.plotAreaRect.height) * 0.35;

  let angle = -Math.PI / 2;
  const nodes: Node[] = [];

  for (let i = 0; i < values.length; i += 1) {
    const v = values[i];
    const slice = (v / total) * Math.PI * 2;
    const next = angle + slice;
    const builder = path().moveTo(cx, cy).arc(cx, cy, r, angle, next).closePath();
    const color = theme.seriesColors[i % theme.seriesColors.length] ?? theme.axis;
    nodes.push({
      kind: "path",
      path: builder.build(),
      fill: toPaint(color),
    });
    angle = next;
  }

  return nodes;
}

function buildPieLegendNodes(data: ResolvedChartData, layout: ChartLayout, theme: ChartTheme): Node[] {
  if (!layout.legendRect) return [];
  const labels = data.series[0]?.categories ?? [];
  const values = data.series[0]?.values?.filter(Number.isFinite) ?? [];
  if (!values.length) return [];

  const padding = 8;
  const markerSize = 10;
  const markerGap = 6;
  const entryGap = 4;
  const lineHeight = theme.fontSize * 1.2;
  const font = toSceneFont(theme.fontFamily, theme.fontSize);

  const nodes: Node[] = [];

  for (let i = 0; i < values.length; i += 1) {
    const y = layout.legendRect.y + padding + i * (lineHeight + entryGap);
    const markerRect = {
      x: layout.legendRect.x + padding,
      y: y + (lineHeight - markerSize) / 2,
      width: markerSize,
      height: markerSize,
    };
    const labelX = markerRect.x + markerRect.width + markerGap;
    const label = String(labels[i] ?? "");
    const color = theme.seriesColors[i % theme.seriesColors.length] ?? theme.axis;
    nodes.push({
      kind: "rect",
      x: markerRect.x,
      y: markerRect.y,
      width: markerRect.width,
      height: markerRect.height,
      fill: toPaint(color),
      stroke: { paint: toPaint(theme.border), width: 0.5 },
    });
    nodes.push({
      kind: "text",
      x: labelX,
      y,
      text: label,
      font,
      fill: toPaint(theme.label),
      align: "start",
      baseline: "top",
      maxWidth: Math.max(0, layout.legendRect.width - (labelX - layout.legendRect.x) - padding),
    });
  }

  return nodes;
}

function buildPlotScene(model: ChartModel, data: ResolvedChartData, layout: ChartLayout, theme: ChartTheme): Node[] {
  const plotRect = layout.plotAreaRect;
  const clipNode = (children: Node[]): Node => ({
    kind: "clip",
    clip: { kind: "rect", x: plotRect.x, y: plotRect.y, width: plotRect.width, height: plotRect.height },
    children,
  });

  const kind = model.chartType.kind;
  let nodes: Node[] = [];

  if (kind === "bar") nodes = buildBarNodes(data, layout, theme);
  else if (kind === "line") nodes = buildLineNodes(model, data, layout, theme);
  else if (kind === "scatter") nodes = buildScatterNodes(data, layout, theme);
  else if (kind === "pie") nodes = buildPieNodes(data, layout, theme);

  if (nodes.length === 0) return [];
  return [clipNode(nodes)];
}

function buildPlaceholderNodes(label: string, sizePx: SizePx, theme: ChartTheme): Node[] {
  const font = toSceneFont(theme.fontFamily, theme.fontSize);
  const baseline: TextBaseline = "middle";
  return [
    {
      kind: "rect",
      x: 0,
      y: 0,
      width: sizePx.width,
      height: sizePx.height,
      fill: toPaint(theme.background),
      stroke: toStroke(theme.border, 1),
    },
    {
      kind: "text",
      x: sizePx.width / 2,
      y: sizePx.height / 2,
      text: label,
      font,
      fill: toPaint(theme.label),
      align: "center",
      baseline,
    },
  ];
}

function buildScene(model: ChartModel, data: ResolvedChartData, theme: ChartTheme, sizePx: SizePx): Scene {
  const layoutModel = applyResolvedDataToModel(model, data);
  const layout = computeChartLayout(layoutModel, buildLayoutTheme(theme), { x: 0, y: 0, width: sizePx.width, height: sizePx.height });

  const nodes: Node[] = [];

  // Background.
  nodes.push({
    kind: "rect",
    x: 0,
    y: 0,
    width: sizePx.width,
    height: sizePx.height,
    fill: toPaint(theme.background),
    stroke: toStroke(theme.border, 1),
  });

  // Gridlines behind plot.
  if (model.chartType.kind !== "pie") {
    nodes.push(...buildGridlineNodes(layout, theme));
  }

  const plotNodes = buildPlotScene(model, data, layout, theme);
  if (plotNodes.length === 0) {
    // If we couldn't render the chart kind, show a placeholder but keep title.
    const placeholder =
      typeof model.options?.placeholder === "string" && model.options.placeholder.trim() !== ""
        ? model.options.placeholder
        : null;
    const unknownName = typeof (model as any)?.chartType?.name === "string" ? String((model as any).chartType.name).trim() : "";
    const label =
      placeholder ??
      (model.chartType.kind === "unknown"
        ? unknownName
          ? `Unsupported chart: ${unknownName}`
          : "Unsupported chart"
        : `Empty ${model.chartType.kind} chart`);
    nodes.push(...buildPlaceholderNodes(label, sizePx, theme));
  } else {
    nodes.push(...plotNodes);
  }

  // Axes/ticks above plot.
  if (model.chartType.kind !== "pie") {
    nodes.push(...buildAxisAndTickNodes(layout, theme));
  }

  // Title and legend on top of plot.
  nodes.push(...buildTitleNodes(layout, theme));
  nodes.push(...buildLegendNodes(model, layout, theme));
  if (model.chartType.kind === "pie") {
    nodes.push(...buildPieLegendNodes(data, layout, theme));
  }

  return { nodes };
}

export function renderChartToSvg(model: ChartModel, data: ResolvedChartData, theme: ChartTheme, sizePx: SizePx): string {
  const resolvedTheme = resolveTheme(theme);
  const scene = buildScene(model, data, resolvedTheme, sizePx);
  return renderSceneToSvg(scene, { width: sizePx.width, height: sizePx.height });
}

export function renderChartToCanvas(
  ctx: CanvasRenderingContext2D,
  model: ChartModel,
  data: ResolvedChartData,
  theme: ChartTheme,
  sizePx: SizePx
): void {
  const resolvedTheme = resolveThemeForCanvas(resolveTheme(theme));
  const scene = buildScene(model, data, resolvedTheme, sizePx);
  ctx.save();
  ctx.clearRect(0, 0, sizePx.width, sizePx.height);
  renderSceneToCanvas(scene, ctx);
  ctx.restore();
}
