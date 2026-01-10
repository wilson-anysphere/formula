/**
 * @param {import('./types.js').RichTextRunStyle | undefined} style
 * @param {{fontFamily: string, fontSizePx: number}} defaults
 */
function fontStringForStyle(style, defaults) {
  const parts = [];
  if (style?.italic) parts.push("italic");
  if (style?.bold) parts.push("bold");

  const size = style?.size ?? defaults.fontSizePx;
  const family = style?.font ?? defaults.fontFamily;
  parts.push(`${size}px`);
  parts.push(family);
  return parts.join(" ");
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
  const defaultColor = options.color ?? "#000000";

  const runs = Array.isArray(richText.runs) && richText.runs.length > 0
    ? richText.runs
    : [{ start: 0, end: richText.text.length, style: undefined }];

  // Pre-measure width per run for alignment.
  const measured = runs.map((run) => {
    const text = richText.text.slice(run.start, run.end);
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
    ctx.fillStyle = fragment.style?.color ?? defaultColor;
    ctx.fillText(fragment.text, cursorX, baselineY);

    if (fragment.style?.underline && fragment.text.length > 0) {
      // Basic underline rendering. Canvas doesn't support underline natively.
      const fontSize = fragment.style?.size ?? defaults.fontSizePx;
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

