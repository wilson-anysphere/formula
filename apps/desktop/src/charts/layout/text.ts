import type { FontSpec, Rect, TextBlockLayout, TextLineLayout } from "./types";
import { round } from "./geometry";

export interface TextMeasure {
  width: number;
  height: number;
}

export function estimateLineHeight(font: FontSpec): number {
  return font.sizePx * 1.2;
}

export function estimateTextWidth(text: string, font: FontSpec): number {
  // Deterministic heuristic for environments without Canvas (node tests).
  // Approximate average glyph width as ~0.6em (common for sans fonts).
  return text.length * font.sizePx * 0.6;
}

export function wrapTextToWidth(text: string, font: FontSpec, maxWidth: number): string[] {
  const clean = String(text ?? "").replace(/\s+/g, " ").trim();
  if (!clean) return [];

  const words = clean.split(" ");
  const lines: string[] = [];
  let current = "";

  function flush() {
    const s = current.trim();
    if (s) lines.push(s);
    current = "";
  }

  for (const word of words) {
    if (!current) {
      current = word;
      continue;
    }

    const candidate = `${current} ${word}`;
    if (estimateTextWidth(candidate, font) <= maxWidth) {
      current = candidate;
      continue;
    }

    flush();
    // Word may still be too long - break it deterministically.
    if (estimateTextWidth(word, font) <= maxWidth) {
      current = word;
      continue;
    }

    let chunk = "";
    for (const ch of word) {
      const next = chunk + ch;
      if (estimateTextWidth(next, font) > maxWidth && chunk) {
        lines.push(chunk);
        chunk = ch;
      } else {
        chunk = next;
      }
    }
    current = chunk;
  }

  flush();
  return lines;
}

export function layoutTextBlock(args: {
  rect: Rect;
  font: FontSpec;
  align: "start" | "center" | "end";
  lines: string[];
  paddingY?: number;
}): TextBlockLayout {
  const paddingY = args.paddingY ?? 0;
  const lineHeight = estimateLineHeight(args.font);
  const totalHeight = args.lines.length * lineHeight;
  const startY = args.rect.y + paddingY + Math.max(0, (args.rect.height - paddingY * 2 - totalHeight) / 2);

  /** @type {TextLineLayout[]} */
  const outLines: TextLineLayout[] = [];

  for (let i = 0; i < args.lines.length; i += 1) {
    const text = args.lines[i];
    const width = estimateTextWidth(text, args.font);
    const height = lineHeight;
    let x = args.rect.x;
    if (args.align === "center") x = args.rect.x + (args.rect.width - width) / 2;
    if (args.align === "end") x = args.rect.x + (args.rect.width - width);
    const y = startY + i * lineHeight;
    outLines.push({
      text,
      x: round(x),
      y: round(y),
      width: round(width),
      height: round(height),
    });
  }

  return {
    rect: args.rect,
    font: args.font,
    lineHeight: round(lineHeight),
    align: args.align,
    lines: outLines,
  };
}

