import { resolveCssVar } from "../../../theme/cssVars.js";
import { detectBaseDirection, resolveAlign, toCanvasFontString } from "@formula/text-layout";
import { getSharedTextLayoutEngine } from "../textLayout.js";

/**
 * @param {import('./types.js').RichTextRunStyle | undefined} style
 * @param {{fontFamily: string, fontSizePx: number}} defaults
 */
function fontSpecForStyle(style, defaults) {
  const sizePx = style?.size_100pt != null
    ? pointsToPx(style.size_100pt / 100)
    : defaults.fontSizePx;
  return {
    family: style?.font ?? defaults.fontFamily,
    sizePx,
    weight: style?.bold ? "bold" : "normal",
    style: style?.italic ? "italic" : "normal",
  };
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
 *   align?: 'left' | 'center' | 'right' | 'start' | 'end',
 *   verticalAlign?: 'top' | 'middle' | 'bottom',
 *   wrapMode?: 'none' | 'word' | 'char',
 *   direction?: 'ltr' | 'rtl' | 'auto',
 *   fontFamily?: string,
 *   fontSizePx?: number,
 *   color?: string,
 * }} [options]
 */
export function renderRichText(ctx, richText, rect, options = {}) {
  const padding = options.padding ?? 2;
  const align = options.align ?? "start";
  const verticalAlign = options.verticalAlign ?? "middle";
  const wrapMode = options.wrapMode ?? "none";
  const direction = options.direction ?? "auto";

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

  const engine = getSharedTextLayoutEngine(ctx);

  /** @type {Array<{ text: string, font: any, color: string | undefined, underline: boolean }>} */
  const layoutRuns = runs.map((run) => {
    const text = sliceByCodePointRange(richText.text, offsets, run.start, run.end);
    const font = fontSpecForStyle(run.style, defaults);
    return {
      text,
      font,
      color: engineColorToCanvasColor(run.style?.color),
      underline: Boolean(run.style?.underline && run.style.underline !== "none"),
    };
  });

  const fullText = layoutRuns.map((r) => r.text).join("");
  const hasExplicitNewline = /[\r\n]/.test(fullText);

  const availableWidth = Math.max(0, rect.width - padding * 2);
  const availableHeight = Math.max(0, rect.height - padding * 2);
  const maxFontSizePx = layoutRuns.reduce((acc, run) => Math.max(acc, run.font.sizePx), defaults.fontSizePx);
  const lineHeight = Math.ceil(maxFontSizePx * 1.2);
  const maxLines = Math.max(1, Math.floor(availableHeight / lineHeight));

  ctx.save();
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, rect.width, rect.height);
  ctx.clip();
  ctx.textAlign = "left";
  ctx.textBaseline = "alphabetic";

  if (wrapMode === "none" && !hasExplicitNewline) {
    // Fast path: single-line rich text (no wrapping). Uses cached measurements.
    const baseDir = direction === "auto" ? detectBaseDirection(fullText) : direction;
    const resolvedAlign =
      align === "left" || align === "right" || align === "center"
        ? align
        : resolveAlign(align, baseDir);

    const fragments = layoutRuns
      .map((fragment) => {
        const measurement = engine.measure(fragment.text, fragment.font);
        return { ...fragment, measurement, width: measurement.width };
      })
      .filter((fragment) => fragment.text.length > 0);

    const totalWidth = fragments.reduce((acc, fragment) => acc + fragment.width, 0);
    let cursorX = rect.x + padding;
    if (resolvedAlign === "center") {
      cursorX = rect.x + padding + (availableWidth - totalWidth) / 2;
    } else if (resolvedAlign === "right") {
      cursorX = rect.x + padding + (availableWidth - totalWidth);
    }

    const lineAscent = fragments.reduce((acc, fragment) => Math.max(acc, fragment.measurement.ascent), 0);
    const lineDescent = fragments.reduce((acc, fragment) => Math.max(acc, fragment.measurement.descent), 0);

    let baselineY = rect.y + padding + lineAscent;
    if (verticalAlign === "middle") {
      baselineY = rect.y + rect.height / 2 + (lineAscent - lineDescent) / 2;
    } else if (verticalAlign === "bottom") {
      baselineY = rect.y + rect.height - padding - lineDescent;
    }

    for (const fragment of fragments) {
      ctx.font = toCanvasFontString(fragment.font);
      ctx.fillStyle = fragment.color ?? defaultColor;
      ctx.fillText(fragment.text, cursorX, baselineY);

      if (fragment.underline) {
        const fontSize = fragment.font.sizePx;
        const underlineOffset = Math.max(1, Math.round(fontSize * 0.08));
        const underlineY = baselineY + underlineOffset;

        ctx.beginPath();
        ctx.strokeStyle = ctx.fillStyle;
        ctx.lineWidth = Math.max(1, Math.round(fontSize / 16));
        ctx.moveTo(cursorX, underlineY);
        ctx.lineTo(cursorX + fragment.width, underlineY);
        ctx.stroke();
      }

      cursorX += fragment.width;
    }
  } else {
    // Full layout: wrapping and/or explicit newlines.
    const layout = engine.layout({
      runs: layoutRuns.map((r) => ({ text: r.text, font: r.font, color: r.color, underline: r.underline })),
      text: undefined,
      font: { family: defaults.fontFamily, sizePx: defaults.fontSizePx },
      maxWidth: availableWidth,
      wrapMode,
      align,
      direction,
      lineHeightPx: lineHeight,
      maxLines,
    });

    let originY = rect.y + padding;
    if (verticalAlign === "middle") {
      originY = rect.y + padding + Math.max(0, (availableHeight - layout.height) / 2);
    } else if (verticalAlign === "bottom") {
      originY = rect.y + rect.height - padding - layout.height;
    }

    const originX = rect.x + padding;

    for (let i = 0; i < layout.lines.length; i++) {
      const line = layout.lines[i];
      let cursorX = originX + line.x;
      const baselineY = originY + i * layout.lineHeight + line.ascent;

      for (const run of line.runs) {
        const measurement = engine.measure(run.text, run.font);
        ctx.font = toCanvasFontString(run.font);
        ctx.fillStyle = run.color ?? defaultColor;
        ctx.fillText(run.text, cursorX, baselineY);

        if (run.underline) {
          const fontSize = run.font.sizePx;
          const underlineOffset = Math.max(1, Math.round(fontSize * 0.08));
          const underlineY = baselineY + underlineOffset;

          ctx.beginPath();
          ctx.strokeStyle = ctx.fillStyle;
          ctx.lineWidth = Math.max(1, Math.round(fontSize / 16));
          ctx.moveTo(cursorX, underlineY);
          ctx.lineTo(cursorX + measurement.width, underlineY);
          ctx.stroke();
        }

        cursorX += measurement.width;
      }
    }
  }

  ctx.restore();
}
