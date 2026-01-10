import type { RichText } from "./types.js";

export interface RenderRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface RenderRichTextOptions {
  padding?: number;
  align?: "left" | "center" | "right" | "start" | "end";
  verticalAlign?: "top" | "middle" | "bottom";
  wrapMode?: "none" | "word" | "char";
  direction?: "ltr" | "rtl" | "auto";
  fontFamily?: string;
  fontSizePx?: number;
  color?: string;
}

export function renderRichText(
  ctx: CanvasRenderingContext2D,
  richText: RichText,
  rect: RenderRect,
  options?: RenderRichTextOptions,
): void;
