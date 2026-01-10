import { resolveCssVar } from "../../../theme/cssVars.js";

/**
 * @param {import('./types.js').RichTextRunStyle | undefined} style
 * @param {{fontFamily: string, fontSizePx: number}} defaults
 */
function fontStringForStyle(style, defaults) {
  const parts = [];
  if (style?.italic) parts.push("italic");
  if (style?.bold) parts.push("bold");

  const size = style?.size_100pt != null
    ? pointsToPx(style.size_100pt / 100)
    : defaults.fontSizePx;
  const family = style?.font ?? defaults.fontFamily;
  parts.push(`${size}px`);
  parts.push(family);
  return parts.join(" ");
}

function pointsToPx(points) {
  // Excel point sizes are typically interpreted at 96DPI.
  return (points * 96) / 72;
}

function engineColorToCanvasColor(color) {
  if (typeof color !== "string") return undefined;
  if (!color.startsWith("#")) return color;
  if (color.length !== 9) return color;

  // Engine colors are serialized as `#AARRGGBB`.
  const hex = color.slice(1);
  const a = Number.parseInt(hex.slice(0, 2), 16) / 255;
  const r = Number.parseInt(hex.slice(2, 4), 16);
  const g = Number.parseInt(hex.slice(4, 6), 16);
  const b = Number.parseInt(hex.slice(6, 8), 16);

  if (Number.isNaN(a) || Number.isNaN(r) || Number.isNaN(g) || Number.isNaN(b)) {
    return color;
  }

  if (a >= 1) {
    return `rgb(${r}, ${g}, ${b})`;
  }
  return `rgba(${r}, ${g}, ${b}, ${a})`;
}

function buildCodePointIndex(text) {
  /** @type {number[]} */
  const offsets = [0];
  let utf16Offset = 0;
  for (const ch of text) {
    utf16Offset += ch.length;
    offsets.push(utf16Offset);
  }
  return offsets;
}

function sliceByCodePointRange(text, offsets, start, end) {
  const len = offsets.length - 1;
  const s = Math.max(0, Math.min(len, start));
  const e = Math.max(s, Math.min(len, end));
  return text.slice(offsets[s], offsets[e]);
}

/**
 * Render rich text into a cell rectangle.
 *
 * This is designed to be called by the grid renderer.
 *
 * @param {CanvasRenderingContext2D} ctx
 * @param {import('./types.js').RichText} richText
 * @param {{x: number, y: number, width: number, height: number}} rect
 * @param {{
 *   padding?: number,
 *   align?: 'left' | 'center' | 'right',
 *   verticalAlign?: 'top' | 'middle' | 'bottom',
 *   fontFamily?: string,
 *   fontSizePx?: number,
 *   color?: string,
 * }} [options]
 */
export function renderRichText(ctx, richText, rect, options = {}) {
  const padding = options.padding ?? 2;
  const align = options.align ?? "left";
  const verticalAlign = options.verticalAlign ?? "middle";

  const defaults = {
    fontFamily: options.fontFamily ?? "Calibri",
    fontSizePx: options.fontSizePx ?? 12,
  };
  const defaultColor = options.color ?? resolveCssVar("--text-primary", { fallback: "CanvasText" });

  const offsets = buildCodePointIndex(richText.text);
  const textLen = offsets.length - 1;

  const runs = Array.isArray(richText.runs) && richText.runs.length > 0
    ? richText.runs
    : [{ start: 0, end: textLen, style: undefined }];

  // Pre-measure width per run for alignment.
  const measured = runs.map((run) => {
    const text = sliceByCodePointRange(richText.text, offsets, run.start, run.end);
    const style = run.style;
    const font = fontStringForStyle(style, defaults);
    ctx.font = font;
    const width = ctx.measureText(text).width;
    return { text, style, font, width };
  });

  const totalWidth = measured.reduce((acc, m) => acc + m.width, 0);
  let cursorX = rect.x + padding;
  if (align === "center") {
    cursorX = rect.x + (rect.width - totalWidth) / 2;
  } else if (align === "right") {
    cursorX = rect.x + rect.width - padding - totalWidth;
  }

  let baselineY = rect.y + rect.height / 2;
  if (verticalAlign === "top") baselineY = rect.y + padding;
  if (verticalAlign === "bottom") baselineY = rect.y + rect.height - padding;

  ctx.save();
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, rect.width, rect.height);
  ctx.clip();

  ctx.textBaseline = verticalAlign === "middle" ? "middle" : (verticalAlign === "top" ? "top" : "bottom");

  for (const fragment of measured) {
    ctx.font = fragment.font;
    ctx.fillStyle = engineColorToCanvasColor(fragment.style?.color) ?? defaultColor;
    ctx.fillText(fragment.text, cursorX, baselineY);

    if (fragment.style?.underline && fragment.style.underline !== "none" && fragment.text.length > 0) {
      // Basic underline rendering. Canvas doesn't support underline natively.
      const fontSize = fragment.style?.size_100pt != null
        ? pointsToPx(fragment.style.size_100pt / 100)
        : defaults.fontSizePx;
      const underlineOffset = Math.max(1, Math.round(fontSize * 0.08));
      const underlineY =
        ctx.textBaseline === "top"
          ? baselineY + fontSize + underlineOffset
          : ctx.textBaseline === "bottom"
            ? baselineY - underlineOffset
            : baselineY + Math.round(fontSize * 0.5);

      ctx.beginPath();
      ctx.strokeStyle = ctx.fillStyle;
      ctx.lineWidth = Math.max(1, Math.round(fontSize / 16));
      ctx.moveTo(cursorX, underlineY);
      ctx.lineTo(cursorX + fragment.width, underlineY);
      ctx.stroke();
    }

    cursorX += fragment.width;
  }

  ctx.restore();
}
