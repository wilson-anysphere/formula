import type { RichText } from "./types.js";

export interface RenderRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface RenderRichTextOptions {
  padding?: number;
  align?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";
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

