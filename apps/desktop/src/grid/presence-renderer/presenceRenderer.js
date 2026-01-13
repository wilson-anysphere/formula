import { t } from "../../i18n/index.js";
import { resolveCssVar } from "../../theme/cssVars.js";

function parseHexColor(color) {
  const match = /^#?([0-9a-f]{6})$/i.exec(color);
  if (!match) return null;
  const value = Number.parseInt(match[1], 16);
  return {
    r: (value >> 16) & 255,
    g: (value >> 8) & 255,
    b: value & 255,
  };
}

function pickTextColor(backgroundColor) {
  const rgb = parseHexColor(backgroundColor);
  if (!rgb) return "#ffffff";
  const luma = (0.2126 * rgb.r + 0.7152 * rgb.g + 0.0722 * rgb.b) / 255;
  return luma > 0.6 ? "#000000" : "#ffffff";
}

function normalizeRange(range) {
  if (!range || typeof range !== "object") return null;

  let startRow;
  let startCol;
  let endRow;
  let endCol;

  if (
    typeof range.startRow === "number" &&
    typeof range.startCol === "number" &&
    typeof range.endRow === "number" &&
    typeof range.endCol === "number"
  ) {
    startRow = range.startRow;
    startCol = range.startCol;
    endRow = range.endRow;
    endCol = range.endCol;
  } else if (
    range.start &&
    range.end &&
    typeof range.start.row === "number" &&
    typeof range.start.col === "number" &&
    typeof range.end.row === "number" &&
    typeof range.end.col === "number"
  ) {
    startRow = range.start.row;
    startCol = range.start.col;
    endRow = range.end.row;
    endCol = range.end.col;
  } else {
    return null;
  }

  const normalizedStartRow = Math.min(startRow, endRow);
  const normalizedEndRow = Math.max(startRow, endRow);
  const normalizedStartCol = Math.min(startCol, endCol);
  const normalizedEndCol = Math.max(startCol, endCol);

  return {
    startRow: Math.trunc(normalizedStartRow),
    startCol: Math.trunc(normalizedStartCol),
    endRow: Math.trunc(normalizedEndRow),
    endCol: Math.trunc(normalizedEndCol),
  };
}

function rectForRange(getCellRect, range) {
  const normalized = normalizeRange(range);
  if (!normalized) return null;

  const startRect = getCellRect(normalized.startRow, normalized.startCol);
  const endRect = getCellRect(normalized.endRow, normalized.endCol);
  if (!startRect || !endRect) return null;

  const x1 = Math.min(startRect.x, endRect.x);
  const y1 = Math.min(startRect.y, endRect.y);
  const x2 = Math.max(startRect.x + startRect.width, endRect.x + endRect.width);
  const y2 = Math.max(startRect.y + startRect.height, endRect.y + endRect.height);

  return { x: x1, y: y1, width: x2 - x1, height: y2 - y1 };
}

export class PresenceRenderer {
  constructor(options) {
    const {
      selectionFillAlpha = 0.12,
      selectionStrokeAlpha = 0.9,
      cursorStrokeWidth = 2,
      badgePaddingX = 6,
      badgePaddingY = 3,
      badgeOffsetX = 8,
      badgeOffsetY = -18,
      font,
    } = options ?? {};

    this.selectionFillAlpha = selectionFillAlpha;
    this.selectionStrokeAlpha = selectionStrokeAlpha;
    this.cursorStrokeWidth = cursorStrokeWidth;
    this.badgePaddingX = badgePaddingX;
    this.badgePaddingY = badgePaddingY;
    this.badgeOffsetX = badgeOffsetX;
    this.badgeOffsetY = badgeOffsetY;
    const fontFamily = resolveCssVar("--font-sans", { fallback: "system-ui, sans-serif" });
    this.font = font ?? `12px ${fontFamily}`;

    this._badgeWidthCache = new Map();
  }

  clear(ctx) {
    // `clearRect` is affected by the current transform. Reset to identity to
    // clear the full backing store regardless of DPR scaling.
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
    ctx.restore();
  }

  render(ctx, presences, options) {
    const { getCellRect } = options ?? {};
    if (typeof getCellRect !== "function") return;

    ctx.save();
    ctx.font = this.font;
    ctx.lineWidth = this.cursorStrokeWidth;
    ctx.textBaseline = "top";

    for (const presence of presences) {
      const color = presence.color ?? "#4c8bf5";

      if (Array.isArray(presence.selections)) {
        ctx.fillStyle = color;
        ctx.strokeStyle = color;

        /** @type {Array<{ x: number; y: number; width: number; height: number }>} */
        const rects = [];
        for (const selection of presence.selections) {
          const rect = rectForRange(getCellRect, selection);
          if (!rect) continue;
          rects.push(rect);
        }

        if (rects.length > 0) {
          ctx.globalAlpha = this.selectionFillAlpha;
          for (const rect of rects) {
            ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
          }

          ctx.globalAlpha = this.selectionStrokeAlpha;
          for (const rect of rects) {
            // Sub-pixel alignment for crisp borders.
            ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width - 1, rect.height - 1);
          }
        }

        ctx.globalAlpha = 1;
      }

      if (presence.cursor) {
        const cursorRect = getCellRect(presence.cursor.row, presence.cursor.col);
        if (!cursorRect) continue;

        ctx.globalAlpha = 1;
        ctx.strokeStyle = color;
        ctx.strokeRect(
          cursorRect.x + this.cursorStrokeWidth / 2,
          cursorRect.y + this.cursorStrokeWidth / 2,
          cursorRect.width - this.cursorStrokeWidth,
          cursorRect.height - this.cursorStrokeWidth,
        );

        const name = presence.name ?? t("presence.anonymous");
        const cacheKey = `${this.font}::${name}`;
        let textWidth = this._badgeWidthCache.get(cacheKey);
        if (textWidth === undefined) {
          textWidth = ctx.measureText(name).width;
          this._badgeWidthCache.set(cacheKey, textWidth);
        }

        const badgeWidth = textWidth + this.badgePaddingX * 2;
        const badgeHeight = 14 + this.badgePaddingY * 2;
        const badgeX = cursorRect.x + cursorRect.width + this.badgeOffsetX;
        const badgeY = cursorRect.y + this.badgeOffsetY;

        ctx.fillStyle = color;
        ctx.fillRect(badgeX, badgeY, badgeWidth, badgeHeight);
        ctx.fillStyle = pickTextColor(color);
        ctx.fillText(name, badgeX + this.badgePaddingX, badgeY + this.badgePaddingY);
      }
    }

    ctx.restore();
  }
}
