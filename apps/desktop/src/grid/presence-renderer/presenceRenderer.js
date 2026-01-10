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
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);

  return { startRow, startCol, endRow, endCol };
}

function rectForRange(getCellRect, range) {
  const normalized = normalizeRange(range);
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
      font = "12px system-ui, sans-serif",
    } = options ?? {};

    this.selectionFillAlpha = selectionFillAlpha;
    this.selectionStrokeAlpha = selectionStrokeAlpha;
    this.cursorStrokeWidth = cursorStrokeWidth;
    this.badgePaddingX = badgePaddingX;
    this.badgePaddingY = badgePaddingY;
    this.badgeOffsetX = badgeOffsetX;
    this.badgeOffsetY = badgeOffsetY;
    this.font = font;

    this._badgeWidthCache = new Map();
  }

  clear(ctx) {
    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
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

        for (const selection of presence.selections) {
          const rect = rectForRange(getCellRect, selection);
          if (!rect) continue;

          ctx.globalAlpha = this.selectionFillAlpha;
          ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
          ctx.globalAlpha = this.selectionStrokeAlpha;
          ctx.strokeRect(rect.x, rect.y, rect.width, rect.height);
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

        const name = presence.name ?? "Anonymous";
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
