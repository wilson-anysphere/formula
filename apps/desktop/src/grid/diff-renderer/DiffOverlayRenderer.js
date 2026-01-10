/**
 * Given a semantic diff result, compute per-cell highlight type. The grid
 * renderer can use this to draw overlays (green/red/yellow).
 */

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

