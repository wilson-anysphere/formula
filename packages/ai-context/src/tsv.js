import { isCellEmpty } from "./a1.js";
import { throwIfAborted } from "./abort.js";

/**
 * Convert a sub-range of a sheet's value matrix to TSV.
 *
 * This is intentionally streaming-ish: it only reads up to `maxRows` rows from `values`
 * rather than allocating a full `slice2D()` copy of the entire range.
 *
 * @param {unknown[][]} values
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} range
 * @param {{ maxRows: number, signal?: AbortSignal }} options
 */
export function valuesRangeToTsv(values, range, options) {
  const signal = options.signal;
  const shouldCheckAbort = Boolean(signal);
  const lines = [];
  const totalRows = range.endRow - range.startRow + 1;
  const limit = Math.min(totalRows, options.maxRows);

  for (let rOffset = 0; rOffset < limit; rOffset++) {
    if (shouldCheckAbort) throwIfAborted(signal);
    const row = values[range.startRow + rOffset];
    if (!Array.isArray(row)) {
      lines.push("");
      continue;
    }

    // Preserve `slice2D(...)+matrixToTsv(...)` ragged-row semantics:
    // only include columns that exist in the source row slice.
    const rowLen = row.length;
    if (rowLen <= range.startCol) {
      lines.push("");
      continue;
    }

    const sliceLen = Math.max(0, Math.min(rowLen, range.endCol + 1) - range.startCol);
    if (sliceLen === 0) {
      lines.push("");
      continue;
    }

    /** @type {string[]} */
    const cells = new Array(sliceLen);
    for (let cOffset = 0; cOffset < sliceLen; cOffset++) {
      // Avoid calling `throwIfAborted` for every cell when no signal is provided.
      // When a signal exists, check periodically to keep cancellation responsive
      // even for very wide ranges.
      if (shouldCheckAbort && (cOffset & 0x7f) === 0) throwIfAborted(signal);
      const v = row[range.startCol + cOffset];
      cells[cOffset] = isCellEmpty(v) ? "" : String(v);
    }
    lines.push(cells.join("\t"));
  }

  if (totalRows > limit) lines.push(`… (${totalRows - limit} more rows)`);
  return lines.join("\n");
}

