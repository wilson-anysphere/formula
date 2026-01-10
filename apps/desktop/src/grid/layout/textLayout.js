/**
 * @typedef {"general" | "left" | "center" | "right"} HorizontalAlign
 * @typedef {"top" | "center" | "bottom"} VerticalAlign
 * @typedef {{ x: number; y: number; width: number; height: number }} Rect
 * @typedef {{ text: string; x: number; y: number; width: number }} TextLineLayout
 * @typedef {{
 *   wrap: boolean;
 *   horizontalAlign: HorizontalAlign;
 *   verticalAlign: VerticalAlign;
 *   rotationDeg: number;
 *   paddingX?: number;
 *   paddingY?: number;
 *   lineHeight: number;
 * }} TextLayoutStyle
 * @typedef {(text: string) => number} TextMeasure
 * @typedef {{
 *   isCellEmpty(row: number, col: number): boolean;
 *   getColWidth(col: number): number;
 * }} OverflowProbe
 * @typedef {{
 *   clipRect: Rect;
 *   drawRect: Rect;
 *   lines: TextLineLayout[];
 *   rotationDeg: number;
 *   origin: { x: number; y: number };
 * }} TextLayoutResult
 */

/**
 * Compute a render-time text layout for a cell.
 *
 * This is a pure layout function: the caller provides a `measure(text)` function
 * (typically from Canvas `measureText`) and optional neighbor-cell probes for
 * Excel-like overflow behavior.
 *
 * @param {{
 *   text: string;
 *   row: number;
 *   col: number;
 *   cellRect: Rect;
 *   mergedRect?: Rect;
 *   style: TextLayoutStyle;
 *   measure: TextMeasure;
 *   overflow?: OverflowProbe;
 *   maxOverflowCols?: number;
 * }} args
 * @returns {TextLayoutResult}
 */
export function layoutCellText(args) {
  const {
    text,
    row,
    col,
    cellRect,
    mergedRect,
    style,
    measure,
    overflow,
    maxOverflowCols = 256,
  } = args;

  const paddingX = style.paddingX ?? 4;
  const paddingY = style.paddingY ?? 2;
  const baseRect = mergedRect ?? cellRect;

  // Excel clips rotated text to the cell bounds; overflow logic gets very tricky with rotation.
  // For now we disable overflow expansion when rotated.
  const canOverflow = style.rotationDeg === 0 && !style.wrap && overflow != null;

  const clipRect = { ...baseRect };
  const drawRect = canOverflow
    ? expandOverflowRect({
        text,
        row,
        col,
        baseRect,
        horizontalAlign: style.horizontalAlign,
        paddingX,
        measure,
        overflow,
        maxOverflowCols,
      })
    : baseRect;

  const maxLineWidth = Math.max(0, baseRect.width - paddingX * 2);
  const linesText = style.wrap ? wrapText(text, maxLineWidth, measure) : [text];
  const linesWidth = linesText.map((t) => measure(t));
  const contentHeight = linesText.length * style.lineHeight;

  const startY = verticalStart(baseRect, paddingY, contentHeight, style.verticalAlign);
  /** @type {TextLineLayout[]} */
  const lines = linesText.map((t, i) => {
    const lineWidth = linesWidth[i];
    const x = horizontalX(
      baseRect,
      drawRect,
      paddingX,
      lineWidth,
      style.horizontalAlign,
      canOverflow,
    );
    const y = startY + i * style.lineHeight;
    return { text: t, x, y, width: lineWidth };
  });

  // Rotation origin: Excel rotates around the center of the cell's content box.
  const origin = {
    x: baseRect.x + baseRect.width / 2,
    y: baseRect.y + baseRect.height / 2,
  };

  return { clipRect, drawRect, lines, rotationDeg: style.rotationDeg, origin };
}

function expandOverflowRect(args) {
  const { text, row, col, baseRect, horizontalAlign, paddingX, measure, overflow, maxOverflowCols } =
    args;

  const textWidth = measure(text) + paddingX * 2;
  if (textWidth <= baseRect.width) return baseRect;

  if (horizontalAlign === "center") {
    // Excel clips centered text; it does not overflow into neighbor cells.
    return baseRect;
  }

  /** @type {"left" | "right"} */
  const direction = horizontalAlign === "right" ? "left" : "right";

  let extra = 0;
  for (let i = 1; i <= maxOverflowCols; i++) {
    const targetCol = direction === "right" ? col + i : col - i;
    if (targetCol < 0) break;
    if (!overflow.isCellEmpty(row, targetCol)) break;
    extra += overflow.getColWidth(targetCol);
    if (baseRect.width + extra >= textWidth) break;
  }

  if (extra === 0) return baseRect;

  if (direction === "right") {
    return { x: baseRect.x, y: baseRect.y, width: baseRect.width + extra, height: baseRect.height };
  }
  return {
    x: baseRect.x - extra,
    y: baseRect.y,
    width: baseRect.width + extra,
    height: baseRect.height,
  };
}

function wrapText(text, maxWidth, measure) {
  if (maxWidth <= 0) return [text];

  const paragraphs = text.split(/\r?\n/);
  /** @type {string[]} */
  const lines = [];

  for (const para of paragraphs) {
    // Excel treats consecutive spaces as significant; we wrap on runs of whitespace.
    const words = para.split(/(\s+)/).filter((w) => w.length > 0);
    let current = "";

    for (const word of words) {
      const candidate = current.length === 0 ? word : current + word;
      if (measure(candidate) <= maxWidth) {
        current = candidate;
        continue;
      }

      if (current.length > 0) {
        lines.push(current);
        current = "";
      }

      if (measure(word) <= maxWidth) {
        current = word;
        continue;
      }

      // Hard-break long tokens.
      let token = "";
      for (const ch of word) {
        const next = token + ch;
        if (measure(next) > maxWidth && token.length > 0) {
          lines.push(token);
          token = ch;
        } else {
          token = next;
        }
      }
      current = token;
    }

    if (current.length > 0) lines.push(current);
  }

  return lines.length === 0 ? [""] : lines;
}

function verticalStart(rect, paddingY, contentHeight, align) {
  const top = rect.y + paddingY;
  const bottom = rect.y + rect.height - paddingY - contentHeight;
  if (align === "top") return top;
  if (align === "bottom") return Math.max(top, bottom);
  return rect.y + (rect.height - contentHeight) / 2;
}

function horizontalX(baseRect, drawRect, paddingX, lineWidth, align, canOverflow) {
  const effectiveRect = canOverflow ? drawRect : baseRect;
  const left = effectiveRect.x + paddingX;
  const right = effectiveRect.x + effectiveRect.width - paddingX - lineWidth;

  if (align === "right") return Math.max(left, right);
  if (align === "center") {
    return effectiveRect.x + (effectiveRect.width - lineWidth) / 2;
  }
  return left;
}

