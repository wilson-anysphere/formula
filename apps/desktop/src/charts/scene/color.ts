import { clamp01, formatNumber } from "./format.js";
import type { Paint } from "./types.js";

export interface RGBA {
  r: number;
  g: number;
  b: number;
  a: number;
}

type CssParserCtx = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

let cssParserCtx: CssParserCtx | null | undefined;

function getCssParserCtx(): CssParserCtx | null {
  if (cssParserCtx !== undefined) return cssParserCtx;

  if (typeof OffscreenCanvas !== "undefined") {
    try {
      const canvas = new OffscreenCanvas(1, 1);
      cssParserCtx = canvas.getContext("2d");
      return cssParserCtx ?? null;
    } catch {
      cssParserCtx = null;
      return null;
    }
  }

  if (typeof document !== "undefined") {
    try {
      const canvas = document.createElement("canvas");
      cssParserCtx = canvas.getContext("2d");
      return cssParserCtx ?? null;
    } catch {
      cssParserCtx = null;
      return null;
    }
  }

  cssParserCtx = null;
  return null;
}

function clampChannel(value: number): number {
  if (!Number.isFinite(value)) return 0;
  return Math.max(0, Math.min(255, Math.round(value)));
}

function normalizeHex(hex: string): string | null {
  const raw = hex.trim().replace(/^#/, "");
  if (/^[0-9a-fA-F]+$/.test(raw) === false) return null;
  if (raw.length === 3 || raw.length === 4) {
    return raw
      .split("")
      .map((ch) => ch + ch)
      .join("");
  }
  if (raw.length === 6 || raw.length === 8) return raw;
  return null;
}

export function parseColor(input: string): RGBA | null {
  const value = input.trim();
  if (!value) return null;

  if (value === "transparent") return { r: 0, g: 0, b: 0, a: 0 };

  if (value.startsWith("#")) {
    const normalized = normalizeHex(value);
    if (!normalized) return null;

    const r = parseInt(normalized.slice(0, 2), 16);
    const g = parseInt(normalized.slice(2, 4), 16);
    const b = parseInt(normalized.slice(4, 6), 16);
    const a = normalized.length === 8 ? parseInt(normalized.slice(6, 8), 16) / 255 : 1;
    return { r, g, b, a: clamp01(a) };
  }

  const rgbMatch =
    /^rgba?\(\s*([+-]?\d+(?:\.\d+)?)\s*(?:,|\s)\s*([+-]?\d+(?:\.\d+)?)\s*(?:,|\s)\s*([+-]?\d+(?:\.\d+)?)(?:\s*(?:,|\/)\s*([+-]?\d+(?:\.\d+)?%?))?\s*\)$/i.exec(
      value
    );
  if (rgbMatch) {
    const r = clampChannel(Number(rgbMatch[1]));
    const g = clampChannel(Number(rgbMatch[2]));
    const b = clampChannel(Number(rgbMatch[3]));
    const alphaToken = rgbMatch[4];
    let a = 1;
    if (alphaToken != null) {
      if (alphaToken.endsWith("%")) a = Number(alphaToken.slice(0, -1)) / 100;
      else a = Number(alphaToken);
    }
    return { r, g, b, a: clamp01(a) };
  }

  // Last-chance fallback: use the platform's CSS color parser via Canvas.
  // This supports named colors and formats like `hsl(...)` when available, but
  // will not resolve CSS variables (which should be handled upstream).
  const ctx = getCssParserCtx();
  if (ctx) {
    ctx.fillStyle = "#010203";
    const sentinel = ctx.fillStyle;
    ctx.fillStyle = value;
    const parsed = ctx.fillStyle;
    if (parsed !== sentinel) {
      // Some environments (notably unit tests with a stubbed canvas context) do not normalize
      // `fillStyle` assignments. In that case `parsed === value`, and recursing would loop
      // indefinitely (e.g. "black" -> "black" -> ...).
      //
      // If the platform normalized the value (e.g. "black" -> "rgb(0, 0, 0)"), recurse once
      // so we can parse via the fast-path regexes above.
      if (parsed !== value) return parseColor(parsed);
    }
  }

  return null;
}

export function rgbaToHex(rgba: RGBA): string {
  const r = clampChannel(rgba.r).toString(16).padStart(2, "0");
  const g = clampChannel(rgba.g).toString(16).padStart(2, "0");
  const b = clampChannel(rgba.b).toString(16).padStart(2, "0");
  return `#${r}${g}${b}`;
}

export function rgbaToCss(rgba: RGBA): string {
  const r = clampChannel(rgba.r);
  const g = clampChannel(rgba.g);
  const b = clampChannel(rgba.b);
  const a = clamp01(rgba.a);
  return `rgba(${r},${g},${b},${formatNumber(a)})`;
}

export function paintToRgba(paint: Paint): RGBA | null {
  const parsed = parseColor(paint.color);
  if (!parsed) return null;
  const opacity = paint.opacity == null ? 1 : clamp01(paint.opacity);
  return { ...parsed, a: clamp01(parsed.a * opacity) };
}
