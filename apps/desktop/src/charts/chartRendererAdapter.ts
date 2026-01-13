import { resolveCssVar } from "../theme/cssVars.js";
import type { ChartRenderer } from "../drawings/overlay";
import type { Rect } from "../drawings/types";
import {
  defaultChartTheme,
  renderChartToCanvas,
  resolveChartData,
  type ChartModel,
  type ChartTheme,
  type ResolvedChartData,
} from "./renderChart";

export interface ChartStore {
  getChartModel(chartId: string): ChartModel | undefined;
  getChartData(chartId: string): Partial<ResolvedChartData> | undefined;
  getChartTheme(chartId: string): Partial<ChartTheme> | undefined;
  /**
   * Optional monotonic revision counter used to avoid re-rendering charts on scroll.
   *
   * When present, ChartRendererAdapter only re-renders the chart's cached surface when:
   * - the chart rect size changes, or
   * - `getChartRevision(chartId)` changes.
   */
  getChartRevision?(chartId: string): number;
}

type Surface = {
  canvas: HTMLCanvasElement | OffscreenCanvas;
  ctx: CanvasRenderingContext2D;
  width: number;
  height: number;
  /**
   * Cache key for stores that provide an explicit revision counter.
   *
   * For stores without revisions we fall back to comparing model/data/theme
   * identities (see `renderToCanvas`).
   */
  revision: number;
  modelRef: ChartModel | null;
  dataRef: Partial<ResolvedChartData> | undefined;
  themeSig: string;
};

type ParsedVar = { name: string; fallback: string | null };

function parseCssVar(input: string): ParsedVar | null {
  const value = input.trim();
  if (!value.startsWith("var(")) return null;

  const open = value.indexOf("(");
  if (open === -1) return null;

  let depth = 0;
  let close = -1;
  for (let i = open; i < value.length; i += 1) {
    const ch = value[i];
    if (ch === "(") depth += 1;
    else if (ch === ")") {
      depth -= 1;
      if (depth === 0) {
        close = i;
        break;
      }
    }
  }

  if (close === -1) return null;
  if (value.slice(close + 1).trim() !== "") return null;

  const body = value.slice(open + 1, close);
  let split = -1;
  depth = 0;
  for (let i = 0; i < body.length; i += 1) {
    const ch = body[i];
    if (ch === "(") depth += 1;
    else if (ch === ")") depth = Math.max(0, depth - 1);
    else if (ch === "," && depth === 0) {
      split = i;
      break;
    }
  }

  const rawName = (split === -1 ? body : body.slice(0, split)).trim();
  if (!rawName.startsWith("--")) return null;

  const fallback = split === -1 ? null : body.slice(split + 1).trim() || null;
  return { name: rawName, fallback };
}

function resolveCssColor(input: string, fallback: string): string {
  const value = input.trim();
  if (!value) return fallback;

  const parsed = parseCssVar(value);
  if (!parsed) return value;

  const resolved = resolveCssVar(parsed.name, { fallback: "" });
  if (resolved) return resolved;

  if (parsed.fallback) return resolveCssColor(parsed.fallback, fallback);
  return fallback;
}

function resolveThemeForCanvas(theme: ChartTheme): ChartTheme {
  const axis = resolveCssColor(theme.axis, "black");
  return {
    ...theme,
    background: resolveCssColor(theme.background, "white"),
    border: resolveCssColor(theme.border, "black"),
    axis,
    gridline: resolveCssColor(theme.gridline, axis),
    title: resolveCssColor(theme.title, "black"),
    label: resolveCssColor(theme.label, "black"),
    seriesColors: theme.seriesColors.map((color) => resolveCssColor(color, "black")),
  };
}

function getContextScale(ctx: CanvasRenderingContext2D): number {
  try {
    const transform = ctx.getTransform?.();
    const scale = transform?.a;
    if (typeof scale === "number" && Number.isFinite(scale) && scale > 0) return scale;
  } catch {
    // Ignore missing DOMMatrix support.
  }
  return 1;
}

function themeSignature(patch: Partial<ChartTheme> | undefined): string {
  if (!patch) return "";
  const entries: string[] = [];
  for (const key of Object.keys(patch).sort()) {
    const value = (patch as any)[key] as unknown;
    if (key === "seriesColors") {
      if (Array.isArray(value)) {
        entries.push(`seriesColors=${value.join(",")}`);
      }
      continue;
    }
    if (value == null) continue;
    entries.push(`${key}=${String(value)}`);
  }
  return entries.join("|");
}

function createCanvasSurface(width: number, height: number): { canvas: HTMLCanvasElement | OffscreenCanvas; ctx: CanvasRenderingContext2D } {
  if (typeof OffscreenCanvas !== "undefined") {
    const canvas = new OffscreenCanvas(width, height);
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("chart offscreen 2d context not available");
    return { canvas, ctx: ctx as unknown as CanvasRenderingContext2D };
  }

  if (typeof document !== "undefined") {
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("chart canvas 2d context not available");
    return { canvas, ctx };
  }

  throw new Error("chart rendering requires OffscreenCanvas or document");
}

export class ChartRendererAdapter implements ChartRenderer {
  private readonly surfaces = new Map<string, Surface>();

  constructor(private readonly store: ChartStore) {}

  renderToCanvas(ctx: CanvasRenderingContext2D, chartId: string, rect: Rect): void {
    const scale = getContextScale(ctx);
    const widthPx = Math.max(1, Math.round(rect.width * scale));
    const heightPx = Math.max(1, Math.round(rect.height * scale));

    const surface = this.getSurface(chartId, widthPx, heightPx);

    const revisionRaw = this.store.getChartRevision?.(chartId);
    const revision = typeof revisionRaw === "number" && Number.isFinite(revisionRaw) ? revisionRaw : null;

    if (revision != null) {
      const needsRender = surface.revision !== revision;
      if (needsRender) {
        const model = this.store.getChartModel(chartId);
        if (!model) throw new Error(`Chart not found: ${chartId}`);

        const liveData = this.store.getChartData(chartId);
        const data = resolveChartData(model, liveData);

        const themePatch = this.store.getChartTheme(chartId);
        const mergedTheme: ChartTheme = {
          ...defaultChartTheme,
          ...(themePatch ?? {}),
          seriesColors:
            themePatch?.seriesColors && themePatch.seriesColors.length > 0
              ? themePatch.seriesColors
              : defaultChartTheme.seriesColors,
        };
        const theme = resolveThemeForCanvas(mergedTheme);

        renderChartToCanvas(surface.ctx, model, data, theme, { width: widthPx, height: heightPx });
        surface.revision = revision;
        surface.modelRef = model;
        surface.dataRef = liveData;
        surface.themeSig = themeSignature(themePatch);
      }
    } else {
      // Back-compat: for stores without revisions, use identity/small signatures so charts
      // still update when their model/theme/data changes (without re-rendering on scroll).
      const model = this.store.getChartModel(chartId);
      if (!model) throw new Error(`Chart not found: ${chartId}`);
      const liveData = this.store.getChartData(chartId);
      const themePatch = this.store.getChartTheme(chartId);
      const sig = themeSignature(themePatch);

      const needsRender =
        Number.isNaN(surface.revision) ||
        surface.modelRef !== model ||
        surface.dataRef !== liveData ||
        surface.themeSig !== sig;

      if (needsRender) {
        const data = resolveChartData(model, liveData);
        const mergedTheme: ChartTheme = {
          ...defaultChartTheme,
          ...(themePatch ?? {}),
          seriesColors:
            themePatch?.seriesColors && themePatch.seriesColors.length > 0
              ? themePatch.seriesColors
              : defaultChartTheme.seriesColors,
        };
        const theme = resolveThemeForCanvas(mergedTheme);

        renderChartToCanvas(surface.ctx, model, data, theme, { width: widthPx, height: heightPx });
        surface.revision = 0;
        surface.modelRef = model;
        surface.dataRef = liveData;
        surface.themeSig = sig;
      }
    }
    ctx.drawImage(surface.canvas as any, rect.x, rect.y, rect.width, rect.height);
  }

  /**
   * Drop cached offscreen surfaces for charts that are no longer needed.
   *
   * SpreadsheetApp uses this to avoid retaining surfaces for charts on other sheets (or
   * deleted charts) after a sheet switch / chart removal.
   */
  pruneSurfaces(keep: ReadonlySet<string>): void {
    if (!keep || keep.size === 0) {
      this.surfaces.clear();
      return;
    }
    for (const id of this.surfaces.keys()) {
      if (keep.has(id)) continue;
      this.surfaces.delete(id);
    }
  }

  destroy(): void {
    // Drop references to offscreen surfaces so their backing buffers can be GC'd.
    this.surfaces.clear();
  }

  private getSurface(chartId: string, width: number, height: number): Surface {
    const existing = this.surfaces.get(chartId);
    if (existing) {
      if (existing.width !== width || existing.height !== height) {
        existing.canvas.width = width;
        existing.canvas.height = height;
        existing.width = width;
        existing.height = height;
        // Force a rerender after resize.
        existing.revision = Number.NaN;
        existing.modelRef = null;
        existing.dataRef = undefined;
        existing.themeSig = "";
      }
      return existing;
    }

    const created = createCanvasSurface(width, height);
    const surface: Surface = { ...created, width, height, revision: Number.NaN, modelRef: null, dataRef: undefined, themeSig: "" };
    this.surfaces.set(chartId, surface);
    return surface;
  }
}
