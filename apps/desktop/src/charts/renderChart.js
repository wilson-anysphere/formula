function fmt(n) {
  if (!Number.isFinite(n)) return "0";
  // Keep output deterministic across platforms by rounding at a fixed precision.
  const rounded = Math.round(n * 100) / 100;
  // Avoid "-0".
  const normalized = Object.is(rounded, -0) ? 0 : rounded;
  if (Number.isInteger(normalized)) return String(normalized);
  return normalized.toFixed(2);
}

function escapeXml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeAttr(value) {
  return escapeXml(value).replaceAll('"', "&quot;");
}

function num(value) {
  const n = typeof value === "number" ? value : Number(value);
  return Number.isFinite(n) ? n : Number.NaN;
}

function clamp(v, min, max) {
  if (v < min) return min;
  if (v > max) return max;
  return v;
}

function niceNum(range, round) {
  const exponent = Math.floor(Math.log10(range || 1));
  const fraction = range / Math.pow(10, exponent);
  let niceFraction;
  if (round) {
    if (fraction < 1.5) niceFraction = 1;
    else if (fraction < 3) niceFraction = 2;
    else if (fraction < 7) niceFraction = 5;
    else niceFraction = 10;
  } else {
    if (fraction <= 1) niceFraction = 1;
    else if (fraction <= 2) niceFraction = 2;
    else if (fraction <= 5) niceFraction = 5;
    else niceFraction = 10;
  }
  return niceFraction * Math.pow(10, exponent);
}

function generateTicks(minValue, maxValue, desiredCount) {
  const count = Math.max(2, desiredCount ?? 5);
  const min = Number.isFinite(minValue) ? minValue : 0;
  const max = Number.isFinite(maxValue) ? maxValue : 0;
  const range = niceNum(max - min || 1, false);
  const step = niceNum(range / (count - 1), true);
  const niceMin = Math.floor(min / step) * step;
  const niceMax = Math.ceil(max / step) * step;

  const ticks = [];
  // Bound tick count defensively in case of float weirdness.
  const maxTicks = 1000;
  for (let v = niceMin, i = 0; v <= niceMax + step / 2 && i < maxTicks; v += step, i += 1) {
    // Normalize values like 1.2000000000000002
    const rounded = Math.round(v / step) * step;
    ticks.push(rounded);
  }
  return ticks;
}

function formatTick(v) {
  if (!Number.isFinite(v)) return "0";
  const rounded = Math.round(v * 100) / 100;
  if (Number.isInteger(rounded)) return String(rounded);
  return rounded.toFixed(2);
}

/**
 * @typedef {{ x: number; y: number; width: number; height: number }} Rect
 */

/**
 * @typedef {{
 *   type: "group";
 *   children: SceneNode[];
 * }} GroupNode
 */

/**
 * @typedef {{
 *   type: "rect";
 *   x: number;
 *   y: number;
 *   width: number;
 *   height: number;
 *   fill?: string;
 *   stroke?: string;
 *   strokeWidth?: number;
 * }} RectNode
 */

/**
 * @typedef {{
 *   type: "line";
 *   x1: number;
 *   y1: number;
 *   x2: number;
 *   y2: number;
 *   stroke?: string;
 *   strokeWidth?: number;
 *   dash?: number[];
 * }} LineNode
 */

/**
 * @typedef {{
 *   type: "polyline";
 *   points: Array<{ x: number; y: number }>;
 *   fill?: string;
 *   stroke?: string;
 *   strokeWidth?: number;
 * }} PolylineNode
 */

/**
 * @typedef {{
 *   type: "circle";
 *   cx: number;
 *   cy: number;
 *   r: number;
 *   fill?: string;
 *   stroke?: string;
 *   strokeWidth?: number;
 * }} CircleNode
 */

/**
 * @typedef {{
 *   type: "path";
 *   d: string;
 *   fill?: string;
 *   stroke?: string;
 *   strokeWidth?: number;
 * }} PathNode
 */

/**
 * @typedef {{
 *   type: "text";
 *   x: number;
 *   y: number;
 *   text: string;
 *   fill?: string;
 *   fontFamily?: string;
 *   fontSize?: number;
 *   textAnchor?: "start" | "middle" | "end";
 *   dominantBaseline?: "auto" | "middle" | "hanging" | "alphabetic";
 * }} TextNode
 */

/**
 * @typedef {GroupNode | RectNode | LineNode | PolylineNode | CircleNode | PathNode | TextNode} SceneNode
 */

/**
 * @typedef {{
 *   chart: Rect;
 *   plot: Rect;
 *   title?: Rect;
 *   legend?: Rect;
 *   axis?: {
 *     x?: Rect;
 *     y?: Rect;
 *   };
 * }} ChartLayout
 */

/**
 * @typedef {{
 *   background: string;
 *   border: string;
 *   axis: string;
 *   gridline: string;
 *   title: string;
 *   label: string;
 *   fontFamily: string;
 *   fontSize: number;
 *   seriesColors: string[];
 * }} ChartTheme
 */

export const defaultChartTheme = {
  background: "var(--chart-bg)",
  border: "var(--chart-border)",
  axis: "var(--chart-axis)",
  gridline: "var(--chart-gridline, var(--chart-axis))",
  title: "var(--chart-title)",
  label: "var(--chart-label)",
  fontFamily: "sans-serif",
  fontSize: 10,
  seriesColors: [
    "var(--chart-series-1)",
    "var(--chart-series-2)",
    "var(--chart-series-3)",
    "var(--chart-series-4)",
  ],
};

function resolveTheme(theme) {
  if (!theme) return defaultChartTheme;
  return {
    ...defaultChartTheme,
    ...theme,
    seriesColors: Array.isArray(theme.seriesColors) && theme.seriesColors.length > 0 ? theme.seriesColors : defaultChartTheme.seriesColors,
  };
}

/**
 * Prefer cached values from the chart model, but allow a caller to supply live data
 * for series that have missing caches.
 *
 * @param {any} model
 * @param {any} [liveData]
 * @returns {{ series: Array<{ name: string | null; categories: string[]; values: number[]; xValues: number[]; yValues: number[] }> }}
 */
export function resolveChartData(model, liveData) {
  const seriesModels = Array.isArray(model?.series) ? model.series : [];
  const liveSeries = Array.isArray(liveData?.series) ? liveData.series : [];
  const out = [];

  const count = Math.max(seriesModels.length, liveSeries.length);
  for (let i = 0; i < count; i += 1) {
    const m = seriesModels[i] ?? {};
    const l = liveSeries[i] ?? {};
    const name = m.name ?? l.name ?? null;

    const categoriesRaw = m.categories?.strCache ?? l.categories ?? [];
    const valuesRaw = m.values?.numCache ?? l.values ?? [];
    const xRaw = m.xValues?.numCache ?? l.xValues ?? [];
    const yRaw = m.yValues?.numCache ?? l.yValues ?? [];

    out.push({
      name: name == null ? null : String(name),
      categories: Array.isArray(categoriesRaw) ? categoriesRaw.map((v) => String(v ?? "")) : [],
      values: Array.isArray(valuesRaw) ? valuesRaw.map(num) : [],
      xValues: Array.isArray(xRaw) ? xRaw.map(num) : [],
      yValues: Array.isArray(yRaw) ? yRaw.map(num) : [],
    });
  }

  return { series: out };
}

function computeChartLayout(model, sizePx) {
  const width = Math.max(1, sizePx.width);
  const height = Math.max(1, sizePx.height);
  const padding = 8;
  const hasTitle = Boolean(model?.title);
  const titleHeight = hasTitle ? 20 : 0;

  const showLegend = model?.legend?.show !== false;
  const legendPos = model?.legend?.position ?? "right";
  const legendWidth = showLegend && legendPos === "right" ? 96 : 0;
  const legendHeight = showLegend && legendPos === "bottom" ? 48 : 0;
  const legendGap = showLegend ? 8 : 0;

  const kind = model?.chartType?.kind ?? model?.kind ?? "unknown";
  const needsAxes = kind === "bar" || kind === "line" || kind === "scatter";
  const axisLeft = needsAxes ? 44 : 0;
  const axisBottom = needsAxes ? 28 : 0;

  const contentX = padding;
  const contentY = padding + titleHeight;
  const contentW = width - padding * 2 - (legendPos === "right" ? legendWidth + legendGap : 0);
  const contentH = height - contentY - padding - (legendPos === "bottom" ? legendHeight + legendGap : 0);

  const plotX = contentX + (needsAxes ? axisLeft : 0);
  const plotY = contentY;
  const plotW = Math.max(1, contentW - (needsAxes ? axisLeft : 0));
  const plotH = Math.max(1, contentH - (needsAxes ? axisBottom : 0));

  /** @type {ChartLayout} */
  const layout = {
    chart: { x: 0, y: 0, width, height },
    plot: { x: plotX, y: plotY, width: plotW, height: plotH },
  };

  if (hasTitle) {
    layout.title = { x: 0, y: 0, width, height: titleHeight };
  }

  if (showLegend && legendPos === "right") {
    layout.legend = {
      x: width - padding - legendWidth,
      y: contentY,
      width: legendWidth,
      height: Math.max(1, contentH),
    };
  } else if (showLegend && legendPos === "bottom") {
    layout.legend = {
      x: contentX,
      y: height - padding - legendHeight,
      width: Math.max(1, contentW),
      height: legendHeight,
    };
  }

  if (needsAxes) {
    layout.axis = {
      y: { x: contentX, y: plotY, width: axisLeft, height: plotH },
      x: { x: plotX, y: plotY + plotH, width: plotW, height: axisBottom },
    };
  }

  return layout;
}

function placeholderScene(label, theme, sizePx) {
  const t = resolveTheme(theme);
  /** @type {GroupNode} */
  const root = { type: "group", children: [] };
  root.children.push({
    type: "rect",
    x: 0,
    y: 0,
    width: sizePx.width,
    height: sizePx.height,
    fill: "var(--chart-placeholder-bg)",
    stroke: "var(--chart-placeholder-border)",
    strokeWidth: 1,
  });
  root.children.push({
    type: "text",
    x: sizePx.width / 2,
    y: sizePx.height / 2,
    text: label,
    textAnchor: "middle",
    dominantBaseline: "middle",
    fontFamily: t.fontFamily,
    fontSize: 12,
    fill: t.label,
  });
  return root;
}

function buildAxesScene({ kind, layout, theme, categories, xTicks, yTicks, xScale, yScale, showYGridlines }) {
  const nodes = [];
  if (!layout.axis?.x || !layout.axis?.y) return nodes;
  const axisColor = theme.axis;
  const labelColor = theme.label;
  const fontFamily = theme.fontFamily;
  const fontSize = theme.fontSize;

  const plot = layout.plot;
  const x0 = plot.x;
  const y0 = plot.y + plot.height;

  // Axis lines
  nodes.push({ type: "line", x1: x0, y1: y0, x2: x0 + plot.width, y2: y0, stroke: axisColor, strokeWidth: 1 });
  nodes.push({ type: "line", x1: x0, y1: plot.y, x2: x0, y2: y0, stroke: axisColor, strokeWidth: 1 });

  // Y ticks + labels + gridlines
  if (Array.isArray(yTicks)) {
    for (const tv of yTicks) {
      const py = yScale(tv);
      // tick mark
      nodes.push({ type: "line", x1: x0 - 4, y1: py, x2: x0, y2: py, stroke: axisColor, strokeWidth: 1 });
      // label
      nodes.push({
        type: "text",
        x: x0 - 6,
        y: py,
        text: formatTick(tv),
        textAnchor: "end",
        dominantBaseline: "middle",
        fontFamily,
        fontSize,
        fill: labelColor,
      });

      if (showYGridlines) {
        nodes.push({
          type: "line",
          x1: x0,
          y1: py,
          x2: x0 + plot.width,
          y2: py,
          stroke: theme.gridline,
          strokeWidth: 1,
          dash: [2, 2],
        });
      }
    }
  }

  // X ticks + labels
  if (kind === "scatter" && Array.isArray(xTicks)) {
    for (const tv of xTicks) {
      const px = xScale(tv);
      nodes.push({ type: "line", x1: px, y1: y0, x2: px, y2: y0 + 4, stroke: axisColor, strokeWidth: 1 });
      nodes.push({
        type: "text",
        x: px,
        y: y0 + 14,
        text: formatTick(tv),
        textAnchor: "middle",
        dominantBaseline: "middle",
        fontFamily,
        fontSize,
        fill: labelColor,
      });
    }
  } else if ((kind === "bar" || kind === "line") && Array.isArray(categories)) {
    for (let i = 0; i < categories.length; i += 1) {
      const cx = xScale(i + 0.5);
      nodes.push({ type: "line", x1: cx, y1: y0, x2: cx, y2: y0 + 4, stroke: axisColor, strokeWidth: 1 });
      nodes.push({
        type: "text",
        x: cx,
        y: y0 + 14,
        text: String(categories[i] ?? ""),
        textAnchor: "middle",
        dominantBaseline: "middle",
        fontFamily,
        fontSize,
        fill: labelColor,
      });
    }
  }

  return nodes;
}

function buildLegendScene({ layout, entries, colors, theme }) {
  if (!layout.legend) return [];
  if (!entries.length) return [];

  const nodes = [];
  const { x, y, width } = layout.legend;
  const fontFamily = theme.fontFamily;
  const fontSize = theme.fontSize;
  const labelColor = theme.label;

  const padding = 8;
  const rowH = 14;
  const swatch = 10;

  for (let i = 0; i < entries.length; i += 1) {
    const rowY = y + padding + i * rowH;
    nodes.push({
      type: "rect",
      x: x + padding,
      y: rowY,
      width: swatch,
      height: swatch,
      fill: colors[i % colors.length],
      stroke: theme.border,
      strokeWidth: 0.5,
    });
    nodes.push({
      type: "text",
      x: x + padding + swatch + 6,
      y: rowY + swatch / 2,
      text: entries[i],
      textAnchor: "start",
      dominantBaseline: "middle",
      fontFamily,
      fontSize,
      fill: labelColor,
    });
    // Stop if we would draw outside the legend area.
    if (rowY + rowH > y + layout.legend.height) break;
  }

  // Optional legend border for visibility in tests / debug.
  nodes.push({
    type: "rect",
    x,
    y,
    width,
    height: layout.legend.height,
    fill: "transparent",
    stroke: "transparent",
    strokeWidth: 0,
  });

  return nodes;
}

function buildChartScene(model, data, theme, sizePx) {
  const t = resolveTheme(theme);
  const layout = computeChartLayout(model, sizePx);
  const kind = model?.chartType?.kind ?? model?.kind ?? "unknown";

  /** @type {GroupNode} */
  const root = { type: "group", children: [] };

  // Chart background.
  root.children.push({
    type: "rect",
    x: 0,
    y: 0,
    width: sizePx.width,
    height: sizePx.height,
    fill: t.background,
    stroke: t.border,
    strokeWidth: 1,
  });

  // Title
  if (model?.title) {
    root.children.push({
      type: "text",
      x: sizePx.width / 2,
      y: 14,
      text: String(model.title),
      textAnchor: "middle",
      dominantBaseline: "middle",
      fontFamily: t.fontFamily,
      fontSize: 12,
      fill: t.title,
    });
  }

  // Series lookup helpers
  const series = Array.isArray(data?.series) ? data.series : [];
  const seriesColors = t.seriesColors;

  if (kind === "bar") {
    const categories =
      series[0]?.categories?.length > 0
        ? series[0].categories
        : Array.from({ length: series[0]?.values?.length ?? 0 }, (_, i) => String(i + 1));
    const catCount = Math.max(1, categories.length);
    const seriesCount = Math.max(1, series.length);

    const valuesAll = series.flatMap((s) => s.values).filter(Number.isFinite);
    if (valuesAll.length === 0) return placeholderScene("Empty bar chart", t, sizePx);

    const minVal = Math.min(0, ...valuesAll);
    const maxVal = Math.max(0, ...valuesAll);
    const yTicks = generateTicks(minVal, maxVal, model?.axes?.value?.tickCount ?? 5);

    const plot = layout.plot;
    const yScale = (v) => plot.y + plot.height - ((v - minVal) / (maxVal - minVal || 1)) * plot.height;

    const groupW = plot.width / catCount;
    const barW = Math.max(1, (groupW * 0.8) / seriesCount);
    const baseX = plot.x;

    const zeroY = yScale(0);
    for (let ci = 0; ci < catCount; ci += 1) {
      for (let si = 0; si < seriesCount; si += 1) {
        const v = series[si]?.values?.[ci];
        if (!Number.isFinite(v)) continue;

        const x = baseX + ci * groupW + groupW * 0.1 + si * barW;
        const y = Math.min(zeroY, yScale(v));
        const h = Math.abs(zeroY - yScale(v));
        root.children.push({
          type: "rect",
          x,
          y,
          width: barW,
          height: h,
          fill: seriesColors[si % seriesColors.length],
        });
      }
    }

    const showGrid = Boolean(model?.axes?.value?.majorGridlines);
    const xScale = (catIndex) => plot.x + catIndex * groupW;
    root.children.push(
      ...buildAxesScene({
        kind: "bar",
        layout,
        theme: t,
        categories,
        xTicks: null,
        yTicks,
        xScale,
        yScale,
        showYGridlines: showGrid,
      })
    );

    const legendEntries = series.map((s, i) => s.name ?? `Series ${i + 1}`);
    root.children.push(
      ...buildLegendScene({
        layout,
        entries: legendEntries,
        colors: seriesColors,
        theme: t,
      })
    );

    return root;
  }

  if (kind === "line") {
    const categories =
      series[0]?.categories?.length > 0
        ? series[0].categories
        : Array.from({ length: series[0]?.values?.length ?? 0 }, (_, i) => String(i + 1));
    const catCount = Math.max(1, categories.length);
    const seriesCount = Math.max(1, series.length);
    const valuesAll = series.flatMap((s) => s.values).filter(Number.isFinite);
    if (valuesAll.length === 0) return placeholderScene("Empty line chart", t, sizePx);

    const minVal = Math.min(0, ...valuesAll);
    const maxVal = Math.max(0, ...valuesAll);
    const yTicks = generateTicks(minVal, maxVal, model?.axes?.value?.tickCount ?? 5);

    const plot = layout.plot;
    const yScale = (v) => plot.y + plot.height - ((v - minVal) / (maxVal - minVal || 1)) * plot.height;
    const groupW = plot.width / catCount;
    const xScale = (catIndex) => plot.x + catIndex * groupW;

    for (let si = 0; si < seriesCount; si += 1) {
      const pts = [];
      for (let ci = 0; ci < catCount; ci += 1) {
        const v = series[si]?.values?.[ci];
        if (!Number.isFinite(v)) continue;
        pts.push({ x: xScale(ci + 0.5), y: yScale(v) });
      }
      root.children.push({
        type: "polyline",
        points: pts,
        fill: "none",
        stroke: seriesColors[si % seriesColors.length],
        strokeWidth: 2,
      });

      if (model?.options?.markers) {
        for (const p of pts) {
          root.children.push({
            type: "circle",
            cx: p.x,
            cy: p.y,
            r: 3,
            fill: seriesColors[si % seriesColors.length],
          });
        }
      }
    }

    root.children.push(
      ...buildAxesScene({
        kind: "line",
        layout,
        theme: t,
        categories,
        xTicks: null,
        yTicks,
        xScale,
        yScale,
        showYGridlines: false,
      })
    );

    const legendEntries = series.map((s, i) => s.name ?? `Series ${i + 1}`);
    root.children.push(...buildLegendScene({ layout, entries: legendEntries, colors: seriesColors, theme: t }));
    return root;
  }

  if (kind === "scatter") {
    const pointsBySeries = series.map((s) => {
      const pts = [];
      const xs = Array.isArray(s.xValues) ? s.xValues : [];
      const ys = Array.isArray(s.yValues) ? s.yValues : [];
      for (let i = 0; i < Math.min(xs.length, ys.length); i += 1) {
        const x = xs[i];
        const y = ys[i];
        if (Number.isFinite(x) && Number.isFinite(y)) pts.push({ x, y });
      }
      return pts;
    });

    const allPoints = pointsBySeries.flat();
    if (allPoints.length === 0) return placeholderScene("Empty scatter chart", t, sizePx);

    const minX = Math.min(...allPoints.map((p) => p.x));
    const maxX = Math.max(...allPoints.map((p) => p.x));
    const minY = Math.min(...allPoints.map((p) => p.y));
    const maxY = Math.max(...allPoints.map((p) => p.y));

    const plot = layout.plot;
    const xScale = (v) => plot.x + ((v - minX) / (maxX - minX || 1)) * plot.width;
    const yScale = (v) => plot.y + plot.height - ((v - minY) / (maxY - minY || 1)) * plot.height;

    const xTicks = generateTicks(minX, maxX, model?.axes?.x?.tickCount ?? 5);
    const yTicks = generateTicks(minY, maxY, model?.axes?.y?.tickCount ?? 5);

    for (let si = 0; si < pointsBySeries.length; si += 1) {
      const color = seriesColors[si % seriesColors.length];
      for (const p of pointsBySeries[si]) {
        root.children.push({
          type: "circle",
          cx: xScale(p.x),
          cy: yScale(p.y),
          r: 3,
          fill: color,
        });
      }
    }

    root.children.push(
      ...buildAxesScene({
        kind: "scatter",
        layout,
        theme: t,
        categories: null,
        xTicks,
        yTicks,
        xScale,
        yScale,
        showYGridlines: false,
      })
    );

    const legendEntries = series.map((s, i) => s.name ?? `Series ${i + 1}`);
    root.children.push(...buildLegendScene({ layout, entries: legendEntries, colors: seriesColors, theme: t }));
    return root;
  }

  if (kind === "pie") {
    const valuesRaw = series[0]?.values ?? [];
    const values = valuesRaw.filter(Number.isFinite);
    const labels = series[0]?.categories?.length ? series[0].categories : values.map((_, i) => String(i + 1));

    const total = values.reduce((a, b) => a + b, 0);
    if (!(total > 0)) return placeholderScene("Empty pie chart", t, sizePx);

    const plot = layout.plot;
    const cx = plot.x + plot.width / 2;
    const cy = plot.y + plot.height / 2;
    const r = Math.min(plot.width, plot.height) * 0.35;

    let angle = -Math.PI / 2;
    const legendEntries = [];
    const sliceColors = [];
    for (let i = 0; i < values.length; i += 1) {
      const v = values[i];
      const slice = (v / total) * Math.PI * 2;
      const next = angle + slice;

      const x1 = cx + r * Math.cos(angle);
      const y1 = cy + r * Math.sin(angle);
      const x2 = cx + r * Math.cos(next);
      const y2 = cy + r * Math.sin(next);
      const large = slice > Math.PI ? 1 : 0;

      const path = [
        `M ${fmt(cx)} ${fmt(cy)}`,
        `L ${fmt(x1)} ${fmt(y1)}`,
        `A ${fmt(r)} ${fmt(r)} 0 ${large} 1 ${fmt(x2)} ${fmt(y2)}`,
        "Z",
      ].join(" ");

      const color = seriesColors[i % seriesColors.length];
      root.children.push({ type: "path", d: path, fill: color });

      legendEntries.push(String(labels[i] ?? ""));
      sliceColors.push(color);
      angle = next;
    }

    root.children.push(...buildLegendScene({ layout, entries: legendEntries, colors: sliceColors, theme: t }));
    return root;
  }

  return placeholderScene(
    `Unsupported chart (${model?.chartType?.name ?? model?.chartType?.kind ?? "unknown"})`,
    t,
    sizePx
  );
}

function renderSceneNodeToSvg(node) {
  switch (node.type) {
    case "group":
      return `<g>${node.children.map(renderSceneNodeToSvg).join("")}</g>`;
    case "rect": {
      const attrs = [
        `x="${fmt(node.x)}"`,
        `y="${fmt(node.y)}"`,
        `width="${fmt(node.width)}"`,
        `height="${fmt(node.height)}"`,
      ];
      if (node.fill) attrs.push(`fill="${escapeAttr(node.fill)}"`);
      if (node.stroke) attrs.push(`stroke="${escapeAttr(node.stroke)}"`);
      if (node.strokeWidth != null) attrs.push(`stroke-width="${fmt(node.strokeWidth)}"`);
      return `<rect ${attrs.join(" ")}/>`;
    }
    case "line": {
      const attrs = [
        `x1="${fmt(node.x1)}"`,
        `y1="${fmt(node.y1)}"`,
        `x2="${fmt(node.x2)}"`,
        `y2="${fmt(node.y2)}"`,
      ];
      if (node.stroke) attrs.push(`stroke="${escapeAttr(node.stroke)}"`);
      if (node.strokeWidth != null) attrs.push(`stroke-width="${fmt(node.strokeWidth)}"`);
      if (Array.isArray(node.dash) && node.dash.length > 0) {
        attrs.push(`stroke-dasharray="${node.dash.map(fmt).join(",")}"`);
      }
      return `<line ${attrs.join(" ")}/>`;
    }
    case "polyline": {
      const pts = node.points.map((p) => `${fmt(p.x)},${fmt(p.y)}`).join(" ");
      const attrs = [`fill="${escapeAttr(node.fill ?? "none")}"`, `points="${pts}"`];
      if (node.stroke) attrs.push(`stroke="${escapeAttr(node.stroke)}"`);
      if (node.strokeWidth != null) attrs.push(`stroke-width="${fmt(node.strokeWidth)}"`);
      return `<polyline ${attrs.join(" ")}/>`;
    }
    case "circle": {
      const attrs = [`cx="${fmt(node.cx)}"`, `cy="${fmt(node.cy)}"`, `r="${fmt(node.r)}"`];
      if (node.fill) attrs.push(`fill="${escapeAttr(node.fill)}"`);
      if (node.stroke) attrs.push(`stroke="${escapeAttr(node.stroke)}"`);
      if (node.strokeWidth != null) attrs.push(`stroke-width="${fmt(node.strokeWidth)}"`);
      return `<circle ${attrs.join(" ")}/>`;
    }
    case "path": {
      const attrs = [`d="${escapeAttr(node.d)}"`];
      if (node.fill) attrs.push(`fill="${escapeAttr(node.fill)}"`);
      if (node.stroke) attrs.push(`stroke="${escapeAttr(node.stroke)}"`);
      if (node.strokeWidth != null) attrs.push(`stroke-width="${fmt(node.strokeWidth)}"`);
      return `<path ${attrs.join(" ")}/>`;
    }
    case "text": {
      const attrs = [`x="${fmt(node.x)}"`, `y="${fmt(node.y)}"`];
      if (node.textAnchor) attrs.push(`text-anchor="${node.textAnchor}"`);
      if (node.dominantBaseline) attrs.push(`dominant-baseline="${node.dominantBaseline}"`);
      if (node.fontFamily) attrs.push(`font-family="${escapeAttr(node.fontFamily)}"`);
      if (node.fontSize != null) attrs.push(`font-size="${fmt(node.fontSize)}"`);
      if (node.fill) attrs.push(`fill="${escapeAttr(node.fill)}"`);
      return `<text ${attrs.join(" ")}>${escapeXml(node.text)}</text>`;
    }
    default:
      return "";
  }
}

/**
 * Render the chart model + resolved data to an SVG string.
 *
 * @param {any} model
 * @param {any} data
 * @param {ChartTheme} theme
 * @param {{ width: number; height: number }} sizePx
 * @returns {string}
 */
export function renderChartToSvg(model, data, theme, sizePx) {
  const w = sizePx?.width ?? 320;
  const h = sizePx?.height ?? 200;
  const scene = buildChartScene(model, data, theme, { width: w, height: h });
  const kind = model?.chartType?.kind ?? model?.kind ?? "unknown";
  return [
    `<svg xmlns="http://www.w3.org/2000/svg" data-chart-kind="${escapeAttr(kind)}" width="${fmt(w)}" height="${fmt(h)}" viewBox="0 0 ${fmt(w)} ${fmt(h)}">`,
    renderSceneNodeToSvg(scene),
    `</svg>`,
  ].join("");
}

function applyStroke(ctx, node) {
  ctx.lineWidth = node.strokeWidth ?? 1;
  if (node.stroke) ctx.strokeStyle = node.stroke;
  if (Array.isArray(node.dash)) ctx.setLineDash(node.dash);
  else ctx.setLineDash([]);
}

function applyFill(ctx, node) {
  if (node.fill) ctx.fillStyle = node.fill;
}

function renderSceneNodeToCanvas(ctx, node) {
  switch (node.type) {
    case "group":
      for (const child of node.children) renderSceneNodeToCanvas(ctx, child);
      return;
    case "rect":
      if (node.fill && node.fill !== "transparent") {
        applyFill(ctx, node);
        ctx.fillRect(node.x, node.y, node.width, node.height);
      }
      if (node.stroke && node.stroke !== "transparent") {
        applyStroke(ctx, node);
        ctx.strokeRect(node.x, node.y, node.width, node.height);
      }
      return;
    case "line":
      ctx.beginPath();
      ctx.moveTo(node.x1, node.y1);
      ctx.lineTo(node.x2, node.y2);
      applyStroke(ctx, node);
      ctx.stroke();
      return;
    case "polyline":
      if (!node.points.length) return;
      ctx.beginPath();
      ctx.moveTo(node.points[0].x, node.points[0].y);
      for (let i = 1; i < node.points.length; i += 1) {
        ctx.lineTo(node.points[i].x, node.points[i].y);
      }
      applyStroke(ctx, node);
      ctx.stroke();
      return;
    case "circle":
      ctx.beginPath();
      ctx.arc(node.cx, node.cy, node.r, 0, Math.PI * 2);
      if (node.fill) {
        applyFill(ctx, node);
        ctx.fill();
      }
      if (node.stroke) {
        applyStroke(ctx, node);
        ctx.stroke();
      }
      return;
    case "path": {
      // Use Path2D when available (browser). In tests (node env) this typically doesn't run.
      if (typeof Path2D === "undefined") return;
      const path = new Path2D(node.d);
      if (node.fill) {
        applyFill(ctx, node);
        ctx.fill(path);
      }
      if (node.stroke) {
        applyStroke(ctx, node);
        ctx.stroke(path);
      }
      return;
    }
    case "text": {
      ctx.font = `${node.fontSize ?? 10}px ${node.fontFamily ?? "sans-serif"}`;
      ctx.fillStyle = node.fill ?? "black";
      ctx.textAlign = node.textAnchor === "middle" ? "center" : node.textAnchor === "end" ? "right" : "left";
      // dominant-baseline doesn't map cleanly; approximate.
      ctx.textBaseline = "middle";
      ctx.fillText(node.text, node.x, node.y);
      return;
    }
    default:
      return;
  }
}

/**
 * Render the chart model + resolved data onto an existing 2D canvas context.
 *
 * @param {CanvasRenderingContext2D} ctx
 * @param {any} model
 * @param {any} data
 * @param {ChartTheme} theme
 * @param {{ width: number; height: number }} sizePx
 */
export function renderChartToCanvas(ctx, model, data, theme, sizePx) {
  const w = sizePx?.width ?? ctx.canvas.width;
  const h = sizePx?.height ?? ctx.canvas.height;
  ctx.save();
  ctx.clearRect(0, 0, w, h);
  const scene = buildChartScene(model, data, theme, { width: w, height: h });
  renderSceneNodeToCanvas(ctx, scene);
  ctx.restore();
}

