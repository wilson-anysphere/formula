import type { ChartAxisModel, ChartDataCache, ChartLegendModel, ChartModel, ChartSeriesModel, ChartTypeModel } from "./layout/types";

/**
 * Best-effort converter from the JSON-serialized Rust `formula_model::charts::ChartModel`
 * into the UI chart model used by `computeChartLayout` / `renderChartToCanvas`.
 *
 * This function is intentionally tolerant: it should never throw, even when the
 * backend payload changes or is partially missing.
 */
export function formulaModelChartModelToUiChartModel(input: any): ChartModel {
  try {
    const chartType = parseChartType(input?.chartKind ?? input?.chart_kind);
    const title = parseTextModelToPlainString(input?.title);
    const legend = parseLegend(input?.legend);
    const axes = parseAxes(input?.axes);
    const series = parseSeries(input?.series);

    return {
      chartType,
      ...(title != null ? { title } : {}),
      ...(legend != null ? { legend } : {}),
      ...(axes != null ? { axes } : {}),
      series,
    };
  } catch {
    return { chartType: { kind: "unknown" }, series: [] };
  }
}

function parseChartType(value: unknown): ChartTypeModel {
  const parsed = parseChartKind(value);
  if (!parsed) return { kind: "unknown" };
  const kind = normalizeChartKind(parsed.kind);
  if (kind === "unknown") {
    const name = parsed.name ?? parsed.rawName;
    return name ? { kind, name } : { kind };
  }
  return { kind };
}

function parseChartKind(value: unknown): { kind: string; name?: string; rawName?: string } | null {
  if (!value) return null;
  if (typeof value === "string") return { kind: value };
  if (typeof value !== "object") return null;
  const obj: any = value;
  const kind = typeof obj.kind === "string" ? obj.kind : typeof obj.type === "string" ? obj.type : null;
  if (!kind) return null;
  const name = typeof obj.name === "string" ? obj.name : undefined;
  const rawName = typeof obj.rawName === "string" ? obj.rawName : undefined;
  return { kind, ...(name ? { name } : {}), ...(rawName ? { rawName } : {}) };
}

function normalizeChartKind(kind: string): ChartTypeModel["kind"] {
  const k = kind.trim().toLowerCase();
  if (k === "bar") return "bar";
  if (k === "line") return "line";
  if (k === "pie") return "pie";
  if (k === "scatter") return "scatter";
  return "unknown";
}

function parseTextModelToPlainString(value: unknown): string | null {
  if (value == null) return null;
  if (typeof value === "string") return value;
  if (typeof value !== "object") return null;
  const obj: any = value;

  const richText = obj.richText ?? obj.rich_text;
  const fromRich = parseRichTextToPlainString(richText);
  if (fromRich != null) return fromRich;

  // Back-compat: some producers may serialize a plain `text` field.
  if (typeof obj.text === "string") return obj.text;

  // Best-effort: fall back to formula if no rich text is available.
  if (typeof obj.formula === "string") return obj.formula;

  return null;
}

function parseRichTextToPlainString(value: unknown): string | null {
  if (value == null) return null;
  if (typeof value === "string") return value;
  if (typeof value !== "object") return null;
  const obj: any = value;
  if (typeof obj.text === "string") return obj.text;
  return null;
}

function parseLegend(value: unknown): ChartLegendModel | null {
  if (!value || typeof value !== "object") return null;
  const obj: any = value;
  const position = parseLegendPosition(obj.position);
  const overlay = typeof obj.overlay === "boolean" ? obj.overlay : null;

  const out: ChartLegendModel = {};
  if (position != null) out.position = position;
  if (overlay != null) out.overlay = overlay;

  return Object.keys(out).length ? out : null;
}

function parseLegendPosition(value: unknown): ChartLegendModel["position"] | null {
  if (typeof value !== "string") return null;
  const pos = value.trim();
  if (!pos) return null;

  // Rust enum serializes to camelCase strings (e.g. "topRight").
  switch (pos) {
    case "right":
    case "r":
      return "right";
    case "left":
    case "l":
      return "left";
    case "top":
    case "t":
      return "top";
    case "bottom":
    case "b":
      return "bottom";
    case "topRight":
      // v1 layout only supports right; treat topRight as right.
      return "right";
    default:
      return null;
  }
}

function parseAxes(value: unknown): ChartAxisModel[] | null {
  if (!Array.isArray(value)) return null;
  const axes = value
    .map((axis) => parseAxis(axis))
    .filter((axis): axis is ChartAxisModel => axis != null);
  return axes.length ? axes : null;
}

function parseAxis(value: unknown): ChartAxisModel | null {
  if (!value || typeof value !== "object") return null;
  const obj: any = value;

  const kind = normalizeAxisKind(obj.kind);
  const position = normalizeAxisPosition(obj.position) ?? defaultAxisPosition(kind);
  const id = obj.id != null ? String(obj.id) : null;

  const scaling = parseAxisScaling(obj.scaling);
  const numFmt = obj.numFmt ?? obj.num_fmt;
  const formatCode = typeof numFmt?.formatCode === "string" ? numFmt.formatCode : typeof numFmt?.format_code === "string" ? numFmt.format_code : null;

  const majorGridlines = obj.majorGridlines === true || obj.major_gridlines === true ? true : undefined;

  const axis: ChartAxisModel = {
    kind,
    position,
    ...(id ? { id } : {}),
    ...(scaling ? { scaling } : {}),
    ...(formatCode ? { formatCode } : {}),
    ...(majorGridlines ? { majorGridlines } : {}),
  };
  return axis;
}

function normalizeAxisKind(value: unknown): ChartAxisModel["kind"] {
  if (typeof value !== "string") return "value";
  const kind = value.trim();
  if (kind === "category" || kind === "catAx") return "category";
  if (kind === "value" || kind === "valAx") return "value";
  // Rust can serialize "unknown"; prefer value axis as a safe default.
  return "value";
}

function normalizeAxisPosition(value: unknown): ChartAxisModel["position"] | null {
  if (typeof value !== "string") return null;
  const pos = value.trim();
  if (pos === "left" || pos === "l") return "left";
  if (pos === "right" || pos === "r") return "right";
  if (pos === "top" || pos === "t") return "top";
  if (pos === "bottom" || pos === "b") return "bottom";
  return null;
}

function defaultAxisPosition(kind: ChartAxisModel["kind"]): ChartAxisModel["position"] {
  return kind === "category" ? "bottom" : "left";
}

function parseAxisScaling(value: unknown): ChartAxisModel["scaling"] | null {
  if (!value || typeof value !== "object") return null;
  const obj: any = value;
  const min = toFiniteNumberOrNull(obj.min);
  const max = toFiniteNumberOrNull(obj.max);

  // Rust uses `reverse: bool`.
  const reverse = obj.reverse === true;

  // Only include scaling if it carries information. Keeping this sparse makes the
  // resulting model easier to diff/debug.
  if (min == null && max == null && !reverse) return null;

  return {
    ...(min != null ? { min } : {}),
    ...(max != null ? { max } : {}),
    ...(reverse ? { reverseOrder: true } : {}),
  };
}

function parseSeries(value: unknown): ChartSeriesModel[] {
  if (!Array.isArray(value)) return [];
  const series = value
    .map((s) => parseSeriesModel(s))
    .filter((s): s is ChartSeriesModel => s != null);
  return series;
}

function parseSeriesModel(value: unknown): ChartSeriesModel | null {
  if (!value || typeof value !== "object") return null;
  const obj: any = value;

  const name = parseTextModelToPlainString(obj.name);
  // `formula_model::charts::SeriesModel` can represent categories as either
  // `categories` (text) or `categoriesNum` (numeric, e.g. date serials).
  // Prefer text categories when available, otherwise fall back to numeric.
  const categories =
    parseSeriesTextData(obj.categories) ??
    // `parseSeriesNumberData` returns the same cache/ref shape, but with numeric
    // coercion. Cast is safe because UI categories accept `string | number`.
    (parseSeriesNumberData(obj.categoriesNum ?? obj.categories_num) as ChartDataCache<string | number> | null);
  const values = parseSeriesNumberData(obj.values);
  const xValues = parseSeriesData(obj.xValues ?? obj.x_values);
  const yValues = parseSeriesData(obj.yValues ?? obj.y_values);

  const out: ChartSeriesModel = {
    ...(name != null ? { name } : {}),
    ...(categories != null ? { categories } : {}),
    ...(values != null ? { values } : {}),
    ...(xValues != null ? { xValues } : {}),
    ...(yValues != null ? { yValues } : {}),
  };

  // If we couldn't extract anything useful, drop the series.
  return Object.keys(out).length ? out : null;
}

function parseSeriesTextData(value: unknown): ChartDataCache<string | number> | null {
  const { cache, ref } = extractCacheAndRef<string | number>(value, { coerce: (v) => (v == null ? null : String(v)) });
  if (cache) return { cache };
  if (ref) return { ref };
  return null;
}

function parseSeriesNumberData(value: unknown): ChartDataCache<number | string> | null {
  const { cache, ref } = extractCacheAndRef<number | string>(value, {
    coerce: (v) => {
      if (v == null) return null;
      if (typeof v === "number") return Number.isFinite(v) ? v : null;
      const n = Number(v);
      return Number.isFinite(n) ? n : null;
    },
  });
  if (cache) return { cache };
  if (ref) return { ref };
  return null;
}

function parseSeriesData(value: unknown): ChartDataCache<number | string> | null {
  if (!value) return null;
  if (Array.isArray(value)) {
    // Best-effort: assume numeric-ish.
    return { cache: value.map((v) => (v == null ? null : (typeof v === "number" ? v : Number(v)))) };
  }
  if (typeof value !== "object") return null;
  const obj: any = value;
  const kind = typeof obj.kind === "string" ? obj.kind : null;
  if (kind === "text") return parseSeriesTextData(obj) as ChartDataCache<number | string>;
  if (kind === "number") return parseSeriesNumberData(obj);

  // Some producers may omit the tagged union wrapper; try both shapes.
  return parseSeriesNumberData(obj) ?? (parseSeriesTextData(obj) as ChartDataCache<number | string> | null);
}

function extractCacheAndRef<T>(
  value: unknown,
  opts: { coerce: (v: unknown) => T | null }
): { cache: Array<T | null> | null; ref: string | null } {
  if (!value) return { cache: null, ref: null };

  if (Array.isArray(value)) {
    return { cache: value.map(opts.coerce), ref: null };
  }

  if (typeof value !== "object") return { cache: null, ref: null };
  const obj: any = value;

  const rawCache = obj.cache;
  const cache = Array.isArray(rawCache) ? rawCache.map(opts.coerce) : null;

  const formula = typeof obj.formula === "string" ? obj.formula : null;
  const ref = formula && formula.trim() ? formula.trim() : null;

  return { cache, ref };
}

function toFiniteNumberOrNull(value: unknown): number | null {
  if (typeof value !== "number") return null;
  return Number.isFinite(value) ? value : null;
}
