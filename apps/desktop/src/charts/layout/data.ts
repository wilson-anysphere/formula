import type { ChartDataCache, ChartSeriesModel } from "./types";

function extractCache<T>(data: ChartDataCache<T> | null | undefined): Array<T | null> {
  if (!data) return [];
  if (Array.isArray(data)) return data;
  if (typeof data === "object" && "cache" in data && Array.isArray(data.cache)) return data.cache;
  return [];
}

export function extractSeriesStrings(
  series: ChartSeriesModel,
  key: "categories"
): string[] {
  const values = extractCache(series[key]);
  return values.map((v) => (v == null ? "" : String(v)));
}

export function extractSeriesNumbers(
  series: ChartSeriesModel,
  key: "values" | "xValues" | "yValues"
): number[] {
  const values = extractCache(series[key]);
  const out: number[] = [];
  for (const v of values) {
    const n = typeof v === "number" ? v : Number(v);
    if (Number.isFinite(n)) out.push(n);
  }
  return out;
}
