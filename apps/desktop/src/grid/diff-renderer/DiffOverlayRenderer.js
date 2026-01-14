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

function readCssVarColor(varName, root = null, fallback = "transparent") {
  return resolveCssVar(varName, { root, fallback });
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
    // `clearRect` is affected by the current transform. Reset to identity to
    // clear the full backing store regardless of DPR scaling.
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
    ctx.restore();
  }

  /**
   * @param {CanvasRenderingContext2D} ctx
   * @param {import("../../versioning/index.js").DiffResult} diff
   * @param {DiffOverlayOptions} options
   */
  render(ctx, diff, options) {
    const { getCellRect } = options ?? {};
    if (typeof getCellRect !== "function") return;

    const strokeWidth = this.strokeWidth;
    const halfStroke = strokeWidth / 2;

    ctx.save();
    ctx.lineWidth = strokeWidth;

    const fillAlpha = this.fillAlpha;
    const strokeAlpha = this.strokeAlpha;
    const removedStroke = this.removedStrikethrough;
    const removedLineAlpha = 0.9;

    // Added
    ctx.fillStyle = this.addedFill;
    ctx.strokeStyle = this.addedStroke;
    for (const change of diff.added) {
      const rect = getCellRect(change.cell.row, change.cell.col);
      if (!rect) continue;
      ctx.globalAlpha = fillAlpha;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      ctx.globalAlpha = strokeAlpha;
      ctx.strokeRect(rect.x + halfStroke, rect.y + halfStroke, rect.width - strokeWidth, rect.height - strokeWidth);
    }

    // Removed
    ctx.fillStyle = this.removedFill;
    for (const change of diff.removed) {
      const rect = getCellRect(change.cell.row, change.cell.col);
      if (!rect) continue;
      ctx.globalAlpha = fillAlpha;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      ctx.globalAlpha = strokeAlpha;
      ctx.lineWidth = strokeWidth;
      ctx.strokeStyle = this.removedStroke;
      ctx.strokeRect(rect.x + halfStroke, rect.y + halfStroke, rect.width - strokeWidth, rect.height - strokeWidth);

      ctx.globalAlpha = removedLineAlpha;
      ctx.strokeStyle = removedStroke;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(rect.x, rect.y + rect.height / 2);
      ctx.lineTo(rect.x + rect.width, rect.y + rect.height / 2);
      ctx.stroke();
    }

    // Modified
    ctx.fillStyle = this.modifiedFill;
    ctx.strokeStyle = this.modifiedStroke;
    ctx.lineWidth = strokeWidth;
    for (const change of diff.modified) {
      const rect = getCellRect(change.cell.row, change.cell.col);
      if (!rect) continue;
      ctx.globalAlpha = fillAlpha;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      ctx.globalAlpha = strokeAlpha;
      ctx.strokeRect(rect.x + halfStroke, rect.y + halfStroke, rect.width - strokeWidth, rect.height - strokeWidth);
    }

    // Format-only
    ctx.fillStyle = this.formatFill;
    ctx.strokeStyle = this.formatStroke;
    ctx.lineWidth = strokeWidth;
    for (const change of diff.formatOnly) {
      const rect = getCellRect(change.cell.row, change.cell.col);
      if (!rect) continue;
      ctx.globalAlpha = fillAlpha;
      ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
      ctx.globalAlpha = strokeAlpha;
      ctx.strokeRect(rect.x + halfStroke, rect.y + halfStroke, rect.width - strokeWidth, rect.height - strokeWidth);
    }

    // Moved: render old location as removed-ish, new location as added-ish.
    ctx.fillStyle = this.movedFill;
    ctx.strokeStyle = this.movedStroke;
    for (const move of diff.moved) {
      const oldRect = getCellRect(move.oldLocation.row, move.oldLocation.col);
      if (oldRect) {
        ctx.globalAlpha = fillAlpha;
        ctx.fillRect(oldRect.x, oldRect.y, oldRect.width, oldRect.height);
        ctx.globalAlpha = strokeAlpha;
        ctx.lineWidth = strokeWidth;
        ctx.strokeStyle = this.movedStroke;
        ctx.strokeRect(
          oldRect.x + halfStroke,
          oldRect.y + halfStroke,
          oldRect.width - strokeWidth,
          oldRect.height - strokeWidth,
        );

        ctx.globalAlpha = removedLineAlpha;
        ctx.strokeStyle = removedStroke;
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(oldRect.x, oldRect.y + oldRect.height / 2);
        ctx.lineTo(oldRect.x + oldRect.width, oldRect.y + oldRect.height / 2);
        ctx.stroke();
      }

      const newRect = getCellRect(move.newLocation.row, move.newLocation.col);
      if (newRect) {
        ctx.globalAlpha = fillAlpha;
        ctx.fillRect(newRect.x, newRect.y, newRect.width, newRect.height);
        ctx.globalAlpha = strokeAlpha;
        ctx.lineWidth = strokeWidth;
        ctx.strokeStyle = this.movedStroke;
        ctx.strokeRect(newRect.x + halfStroke, newRect.y + halfStroke, newRect.width - strokeWidth, newRect.height - strokeWidth);
      }
    }

    ctx.restore();
  }
}
