import type { FontSpec } from "./types.js";

export function fontSpecToCss(font: FontSpec): string {
  const style = font.style ?? "normal";
  const weight = font.weight ?? "normal";
  return `${style} ${weight} ${font.sizePx}px ${font.family}`;
}

export function approximateTextWidth(text: string, font: FontSpec): number {
  return text.length * font.sizePx * 0.6;
}

export function measureTextWidth(
  text: string,
  font: FontSpec,
  options?: { ctx?: CanvasRenderingContext2D; providedWidth?: number }
): number {
  if (options?.providedWidth != null) return options.providedWidth;
  if (options?.ctx) {
    const prev = options.ctx.font;
    options.ctx.font = fontSpecToCss(font);
    const width = options.ctx.measureText(text).width;
    options.ctx.font = prev;
    return width;
  }

  if (typeof OffscreenCanvas !== "undefined") {
    const canvas = new OffscreenCanvas(1, 1);
    const ctx = canvas.getContext("2d");
    if (ctx) {
      ctx.font = fontSpecToCss(font);
      return ctx.measureText(text).width;
    }
  }

  return approximateTextWidth(text, font);
}

