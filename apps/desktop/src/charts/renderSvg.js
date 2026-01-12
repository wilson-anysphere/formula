import { parseA1Range } from "./a1.js";

function fmt(n) {
  if (!Number.isFinite(n)) return "0";
  const rounded = Math.round(n * 100) / 100;
  if (Number.isInteger(rounded)) return String(rounded);
  return rounded.toFixed(2);
}

function flatten(range2d) {
  if (!Array.isArray(range2d)) return [];
  const out = [];
  for (const row of range2d) {
    if (!Array.isArray(row)) continue;
    for (const value of row) out.push(value);
  }
  return out;
}

function defaultSeriesColors() {
  return [
    "var(--chart-series-1)",
    "var(--chart-series-2)",
    "var(--chart-series-3)",
    "var(--chart-series-4)",
  ];
}

function resolveSeriesColors(theme) {
  if (theme && Array.isArray(theme.seriesColors) && theme.seriesColors.length > 0) {
    return theme.seriesColors;
  }
  return defaultSeriesColors();
}

export function createMatrixRangeProvider(sheets) {
  return {
    getRange(rangeRef) {
      const parsed = parseA1Range(rangeRef);
      if (!parsed) return [];
      const sheetName = parsed.sheetName ?? Object.keys(sheets)[0];
      const sheet = sheets[sheetName];
      if (!sheet) return [];

      const out = [];
      for (let r = parsed.startRow; r <= parsed.endRow; r += 1) {
        const row = [];
        for (let c = parsed.startCol; c <= parsed.endCol; c += 1) {
          row.push(sheet[r]?.[c] ?? null);
        }
        out.push(row);
      }
      return out;
    },
  };
}

export function resolveSeries(chart, provider) {
  const seriesOut = [];
  if (!chart || !Array.isArray(chart.series)) return seriesOut;

  for (const ser of chart.series) {
    const categories = ser.categories ? flatten(provider.getRange(ser.categories)) : [];
    const values = ser.values ? flatten(provider.getRange(ser.values)) : [];
    const xValues = ser.xValues ? flatten(provider.getRange(ser.xValues)) : [];
    const yValues = ser.yValues ? flatten(provider.getRange(ser.yValues)) : [];

    seriesOut.push({
      name: ser.name ?? null,
      categories,
      values,
      xValues,
      yValues,
    });
  }
  return seriesOut;
}

export function placeholderSvg({ width, height, label }) {
  return [
    `<svg xmlns="http://www.w3.org/2000/svg" width="${fmt(width)}" height="${fmt(height)}" viewBox="0 0 ${fmt(width)} ${fmt(height)}">`,
    `<rect x="0" y="0" width="${fmt(width)}" height="${fmt(height)}" fill="var(--chart-placeholder-bg)" stroke="var(--chart-placeholder-border)"/>`,
    `<text x="${fmt(width / 2)}" y="${fmt(height / 2)}" text-anchor="middle" dominant-baseline="middle" font-family="sans-serif" font-size="12" fill="var(--chart-label)">${label}</text>`,
    `</svg>`,
  ].join("");
}

function renderBarLineSvg({ width, height, title, kind, series, seriesColors }) {
  const margin = { left: 36, right: 10, top: 22, bottom: 24 };
  const plotW = Math.max(1, width - margin.left - margin.right);
  const plotH = Math.max(1, height - margin.top - margin.bottom);

  const categories = series[0]?.categories?.length
    ? series[0].categories.map((v) => String(v ?? ""))
    : Array.from({ length: series[0]?.values?.length ?? 0 }, (_, i) => String(i + 1));

  const numericValues = series.map((s) => s.values.map((v) => (typeof v === "number" ? v : Number(v))));
  let maxVal = 0;
  for (const row of numericValues) {
    for (const v of row) {
      if (Number.isFinite(v) && v > maxVal) maxVal = v;
    }
  }

  const svg = [];
  svg.push(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${fmt(width)}" height="${fmt(height)}" viewBox="0 0 ${fmt(width)} ${fmt(height)}">`
  );
  svg.push(`<rect x="0" y="0" width="${fmt(width)}" height="${fmt(height)}" fill="var(--chart-bg)" stroke="var(--chart-border)"/>`);

  if (title) {
    svg.push(
      `<text x="${fmt(width / 2)}" y="14" text-anchor="middle" font-family="sans-serif" font-size="12" fill="var(--chart-title)">${title}</text>`
    );
  }

  const originX = margin.left;
  const originY = margin.top + plotH;
  svg.push(`<line x1="${fmt(originX)}" y1="${fmt(originY)}" x2="${fmt(originX + plotW)}" y2="${fmt(originY)}" stroke="var(--chart-axis)"/>`);
  svg.push(`<line x1="${fmt(originX)}" y1="${fmt(margin.top)}" x2="${fmt(originX)}" y2="${fmt(originY)}" stroke="var(--chart-axis)"/>`);

  const catCount = Math.max(1, categories.length);
  const seriesCount = Math.max(1, series.length);
  const groupW = plotW / catCount;

  const colors = Array.isArray(seriesColors) && seriesColors.length > 0 ? seriesColors : defaultSeriesColors();

  if (kind === "bar") {
    const barW = Math.max(1, (groupW * 0.8) / seriesCount);
    for (let ci = 0; ci < catCount; ci += 1) {
      for (let si = 0; si < seriesCount; si += 1) {
        const v = numericValues[si]?.[ci];
        if (!Number.isFinite(v) || maxVal === 0) continue;
        const h = (v / maxVal) * plotH;
        const x = originX + ci * groupW + groupW * 0.1 + si * barW;
        const y = originY - h;
        svg.push(`<rect x="${fmt(x)}" y="${fmt(y)}" width="${fmt(barW)}" height="${fmt(h)}" fill="${colors[si % colors.length]}"/>`);
      }
    }
  } else if (kind === "area") {
    for (let si = 0; si < seriesCount; si += 1) {
      const points = [];
      for (let ci = 0; ci < catCount; ci += 1) {
        const v = numericValues[si]?.[ci];
        const x = originX + (ci + 0.5) * groupW;
        const y = originY - (Number.isFinite(v) && maxVal !== 0 ? (v / maxVal) * plotH : 0);
        points.push(`${fmt(x)},${fmt(y)}`);
      }

      const firstX = originX + 0.5 * groupW;
      const lastX = originX + (catCount - 0.5) * groupW;
      const areaPoints = [`${fmt(firstX)},${fmt(originY)}`, ...points, `${fmt(lastX)},${fmt(originY)}`];
      svg.push(
        `<polygon points="${areaPoints.join(" ")}" fill="${colors[si % colors.length]}" fill-opacity="0.25" stroke="none"/>`
      );
      svg.push(`<polyline fill="none" stroke="${colors[si % colors.length]}" stroke-width="2" points="${points.join(" ")}"/>`);
    }
  } else {
    for (let si = 0; si < seriesCount; si += 1) {
      const points = [];
      for (let ci = 0; ci < catCount; ci += 1) {
        const v = numericValues[si]?.[ci];
        const x = originX + (ci + 0.5) * groupW;
        const y = originY - (Number.isFinite(v) && maxVal !== 0 ? (v / maxVal) * plotH : 0);
        points.push(`${fmt(x)},${fmt(y)}`);
      }
      svg.push(`<polyline fill="none" stroke="${colors[si % colors.length]}" stroke-width="2" points="${points.join(" ")}"/>`);
    }
  }

  svg.push(`</svg>`);
  return svg.join("");
}

function renderPieSvg({ width, height, title, series, seriesColors }) {
  const values = series[0]?.values?.map((v) => (typeof v === "number" ? v : Number(v))).filter(Number.isFinite) ?? [];
  const labelsRaw = series[0]?.categories?.map((v) => String(v ?? "")) ?? [];
  const labels = labelsRaw.length ? labelsRaw : values.map((_, i) => String(i + 1));
  const total = values.reduce((a, b) => a + b, 0);

  if (total <= 0) {
    return placeholderSvg({ width, height, label: "Empty pie chart" });
  }

  const cx = width / 2;
  const cy = height / 2 + 6;
  const r = Math.min(width, height) * 0.35;
  const colors = Array.isArray(seriesColors) && seriesColors.length > 0 ? seriesColors : defaultSeriesColors();

  let angle = -Math.PI / 2;
  const svg = [];
  svg.push(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${fmt(width)}" height="${fmt(height)}" viewBox="0 0 ${fmt(width)} ${fmt(height)}">`
  );
  svg.push(`<rect x="0" y="0" width="${fmt(width)}" height="${fmt(height)}" fill="var(--chart-bg)" stroke="var(--chart-border)"/>`);
  if (title) {
    svg.push(
      `<text x="${fmt(width / 2)}" y="14" text-anchor="middle" font-family="sans-serif" font-size="12" fill="var(--chart-title)">${title}</text>`
    );
  }

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
    svg.push(`<path d="${path}" fill="${colors[i % colors.length]}"/>`);
    angle = next;
  }

  if (labels.length) {
    svg.push(
      `<text x="${fmt(6)}" y="${fmt(height - 6)}" font-family="sans-serif" font-size="10" fill="var(--chart-label)">${labels.join(", ")}</text>`
    );
  }

  svg.push(`</svg>`);
  return svg.join("");
}

function renderScatterSvg({ width, height, title, series, seriesColors }) {
  const margin = { left: 36, right: 10, top: 22, bottom: 24 };
  const plotW = Math.max(1, width - margin.left - margin.right);
  const plotH = Math.max(1, height - margin.top - margin.bottom);

  const perSeries = series.map((ser) => {
    const xs = ser.xValues?.map((v) => (typeof v === "number" ? v : Number(v))) ?? [];
    const ys = ser.yValues?.map((v) => (typeof v === "number" ? v : Number(v))) ?? [];
    const points = [];
    for (let i = 0; i < Math.min(xs.length, ys.length); i += 1) {
      if (Number.isFinite(xs[i]) && Number.isFinite(ys[i])) points.push({ x: xs[i], y: ys[i] });
    }
    return points;
  });

  const points = perSeries.flat();

  if (points.length === 0) {
    return placeholderSvg({ width, height, label: "Empty scatter chart" });
  }

  let minX = points[0].x;
  let maxX = points[0].x;
  let minY = points[0].y;
  let maxY = points[0].y;
  for (let i = 1; i < points.length; i += 1) {
    const p = points[i];
    if (p.x < minX) minX = p.x;
    if (p.x > maxX) maxX = p.x;
    if (p.y < minY) minY = p.y;
    if (p.y > maxY) maxY = p.y;
  }

  const scaleX = (x) =>
    margin.left + ((x - minX) / (maxX - minX || 1)) * plotW;
  const scaleY = (y) =>
    margin.top + plotH - ((y - minY) / (maxY - minY || 1)) * plotH;

  const svg = [];
  svg.push(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${fmt(width)}" height="${fmt(height)}" viewBox="0 0 ${fmt(width)} ${fmt(height)}">`
  );
  svg.push(`<rect x="0" y="0" width="${fmt(width)}" height="${fmt(height)}" fill="var(--chart-bg)" stroke="var(--chart-border)"/>`);
  if (title) {
    svg.push(
      `<text x="${fmt(width / 2)}" y="14" text-anchor="middle" font-family="sans-serif" font-size="12" fill="var(--chart-title)">${title}</text>`
    );
  }

  const originX = margin.left;
  const originY = margin.top + plotH;
  svg.push(`<line x1="${fmt(originX)}" y1="${fmt(originY)}" x2="${fmt(originX + plotW)}" y2="${fmt(originY)}" stroke="var(--chart-axis)"/>`);
  svg.push(`<line x1="${fmt(originX)}" y1="${fmt(margin.top)}" x2="${fmt(originX)}" y2="${fmt(originY)}" stroke="var(--chart-axis)"/>`);

  const colors = Array.isArray(seriesColors) && seriesColors.length > 0 ? seriesColors : defaultSeriesColors();
  for (let si = 0; si < perSeries.length; si += 1) {
    const color = colors[si % colors.length] ?? colors[0];
    for (const p of perSeries[si]) {
      svg.push(`<circle cx="${fmt(scaleX(p.x))}" cy="${fmt(scaleY(p.y))}" r="3" fill="${color}"/>`);
    }
  }

  svg.push(`</svg>`);
  return svg.join("");
}

export function renderChartSvg(chart, provider, opts) {
  const width = opts?.width ?? 320;
  const height = opts?.height ?? 200;
  const seriesColors = resolveSeriesColors(opts?.theme);

  if (!chart || !chart.chartType || !chart.chartType.kind) {
    return placeholderSvg({ width, height, label: "Missing chart" });
  }

  const series = resolveSeries(chart, provider ?? { getRange: () => [] });

  switch (chart.chartType.kind) {
    case "bar":
      return renderBarLineSvg({ width, height, title: chart.title, kind: "bar", series, seriesColors });
    case "line":
      return renderBarLineSvg({ width, height, title: chart.title, kind: "line", series, seriesColors });
    case "area":
      return renderBarLineSvg({ width, height, title: chart.title, kind: "area", series, seriesColors });
    case "pie":
      return renderPieSvg({ width, height, title: chart.title, series, seriesColors });
    case "scatter":
      return renderScatterSvg({ width, height, title: chart.title, series, seriesColors });
    case "unknown":
    default:
      return placeholderSvg({
        width,
        height,
        label: `Unsupported chart (${chart.chartType.name ?? chart.chartType.kind})`,
      });
  }
}

/**
 * Migration helper: render an SVG from a ChartModel-backed chart (prefering cached
 * series values inside the model).
 *
 * @param {any} model
 * @param {any} [liveData]
 * @param {{ width?: number; height?: number; theme?: any }} [opts]
 */
export async function renderChartSvgFromModel(model, liveData, opts) {
  const width = opts?.width ?? 320;
  const height = opts?.height ?? 200;
  const mod = await import("./renderChart.ts");
  const theme = opts?.theme ?? mod.defaultChartTheme;
  const data = mod.resolveChartData(model, liveData);
  return mod.renderChartToSvg(model, data, theme, { width, height });
}
