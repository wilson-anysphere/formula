/**
 * Diff overlay rendering utilities.
 *
 * - `computeDiffHighlights` is useful for grid engines that already have their
 *   own rendering pipeline.
 * - `DiffOverlayRenderer` is a small canvas overlay renderer aligned with the
 *   presence overlay renderer style.
 */

import { resolveCssVar } from "../../theme/cssVars.js";

/**
 * @param {import("../../versioning/index.js").DiffResult} diff
 */
export function computeDiffHighlights(diff) {
  /** @type {Map<string, "added" | "removed" | "modified" | "formatOnly" | "moved">} */
  const highlights = new Map();

  for (const change of diff.added) highlights.set(`${change.cell.row},${change.cell.col}`, "added");
  for (const change of diff.removed)
    highlights.set(`${change.cell.row},${change.cell.col}`, "removed");
  for (const change of diff.modified)
    highlights.set(`${change.cell.row},${change.cell.col}`, "modified");
  for (const change of diff.formatOnly)
    highlights.set(`${change.cell.row},${change.cell.col}`, "formatOnly");
  for (const move of diff.moved) {
    highlights.set(`${move.oldLocation.row},${move.oldLocation.col}`, "moved");
    highlights.set(`${move.newLocation.row},${move.newLocation.col}`, "moved");
  }

  return highlights;
}

/**
 * @typedef {{ x: number, y: number, width: number, height: number }} Rect
 * @typedef {{ getCellRect: (row: number, col: number) => Rect | null }} DiffOverlayOptions
 */

function withAlpha(color, alpha) {
  // Accept CSS color strings (e.g. hex or computed channel values) by falling back to a
  // `globalAlpha` approach. This exists for callers that want to pass raw token values.
  return { color, alpha };
}

function readCssVarColor(varName, root = null, fallback = "transparent") {
  return resolveCssVar(varName, { root, fallback });
}

function drawCellHighlight(ctx, rect, style, options) {
  const { fill = null, stroke = null, strokeWidth = 2, removed = false } = options ?? {};

  ctx.save();

  if (fill) {
    ctx.fillStyle = fill.color;
    ctx.globalAlpha = fill.alpha;
    ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
  }

  if (stroke) {
    ctx.strokeStyle = stroke.color;
    ctx.globalAlpha = stroke.alpha;
    ctx.lineWidth = strokeWidth;
    ctx.strokeRect(rect.x + strokeWidth / 2, rect.y + strokeWidth / 2, rect.width - strokeWidth, rect.height - strokeWidth);
  }

  if (removed) {
    ctx.globalAlpha = 0.9;
    ctx.strokeStyle = style.removedStroke;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(rect.x, rect.y + rect.height / 2);
    ctx.lineTo(rect.x + rect.width, rect.y + rect.height / 2);
    ctx.stroke();
  }

  ctx.restore();
}

export class DiffOverlayRenderer {
  constructor(options) {
    const root = options?.cssVarRoot ?? globalThis?.document?.documentElement ?? null;

    const palette = {
      added: readCssVarColor("--success", root),
      removed: readCssVarColor("--error", root),
      modified: readCssVarColor("--warning", root),
      format: readCssVarColor("--accent", root),
      moved: readCssVarColor("--accent", root),
    };

    const {
      addedFill = palette.added,
      addedStroke = palette.added,
      removedFill = palette.removed,
      removedStroke = palette.removed,
      modifiedFill = palette.modified,
      modifiedStroke = palette.modified,
      formatFill = palette.format,
      formatStroke = palette.format,
      movedFill = palette.moved,
      movedStroke = palette.moved,
      fillAlpha = 0.18,
      strokeAlpha = 0.85,
      strokeWidth = 2,
    } = options ?? {};

    this.addedFill = addedFill;
    this.addedStroke = addedStroke;
    this.removedFill = removedFill;
    this.removedStroke = removedStroke;
    this.removedStrikethrough = removedStroke;
    this.modifiedFill = modifiedFill;
    this.modifiedStroke = modifiedStroke;
    this.formatFill = formatFill;
    this.formatStroke = formatStroke;
    this.movedFill = movedFill;
    this.movedStroke = movedStroke;
    this.fillAlpha = fillAlpha;
    this.strokeAlpha = strokeAlpha;
    this.strokeWidth = strokeWidth;
  }

  clear(ctx) {
    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
  }

  /**
   * @param {CanvasRenderingContext2D} ctx
   * @param {import("../../versioning/index.js").DiffResult} diff
   * @param {DiffOverlayOptions} options
   */
  render(ctx, diff, options) {
    const { getCellRect } = options ?? {};
    if (typeof getCellRect !== "function") return;

    const style = {
      removedStroke: this.removedStrikethrough,
    };

    // Added
    for (const change of diff.added) {
      const rect = getCellRect(change.cell.row, change.cell.col);
      if (!rect) continue;
      drawCellHighlight(ctx, rect, style, {
        fill: withAlpha(this.addedFill, this.fillAlpha),
        stroke: withAlpha(this.addedStroke, this.strokeAlpha),
        strokeWidth: this.strokeWidth,
      });
    }

    // Removed
    for (const change of diff.removed) {
      const rect = getCellRect(change.cell.row, change.cell.col);
      if (!rect) continue;
      drawCellHighlight(ctx, rect, style, {
        fill: withAlpha(this.removedFill, this.fillAlpha),
        stroke: withAlpha(this.removedStroke, this.strokeAlpha),
        strokeWidth: this.strokeWidth,
        removed: true,
      });
    }

    // Modified
    for (const change of diff.modified) {
      const rect = getCellRect(change.cell.row, change.cell.col);
      if (!rect) continue;
      drawCellHighlight(ctx, rect, style, {
        fill: withAlpha(this.modifiedFill, this.fillAlpha),
        stroke: withAlpha(this.modifiedStroke, this.strokeAlpha),
        strokeWidth: this.strokeWidth,
      });
    }

    // Format-only
    for (const change of diff.formatOnly) {
      const rect = getCellRect(change.cell.row, change.cell.col);
      if (!rect) continue;
      drawCellHighlight(ctx, rect, style, {
        fill: withAlpha(this.formatFill, this.fillAlpha),
        stroke: withAlpha(this.formatStroke, this.strokeAlpha),
        strokeWidth: this.strokeWidth,
      });
    }

    // Moved: render old location as removed-ish, new location as added-ish.
    for (const move of diff.moved) {
      const oldRect = getCellRect(move.oldLocation.row, move.oldLocation.col);
      if (oldRect) {
        drawCellHighlight(ctx, oldRect, style, {
          fill: withAlpha(this.movedFill, this.fillAlpha),
          stroke: withAlpha(this.movedStroke, this.strokeAlpha),
          strokeWidth: this.strokeWidth,
          removed: true,
        });
      }

      const newRect = getCellRect(move.newLocation.row, move.newLocation.col);
      if (newRect) {
        drawCellHighlight(ctx, newRect, style, {
          fill: withAlpha(this.movedFill, this.fillAlpha),
          stroke: withAlpha(this.movedStroke, this.strokeAlpha),
          strokeWidth: this.strokeWidth,
        });
      }
    }
  }
}
