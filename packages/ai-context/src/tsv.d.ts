/**
 * Convert a sub-range of a sheet's value matrix to TSV.
 *
 * This is intentionally streaming-ish: it only reads up to `maxRows` rows from `values`
 * rather than allocating a full `slice2D()` copy of the entire range.
 *
 * Output matches the previous `slice2D(...)+matrixToTsv(...)` behavior:
 * - tab-separated values
 * - empty cells -> ""
 * - ragged rows do not emit trailing tabs beyond the source row length
 * - ellipsis line when truncated
 */
export function valuesRangeToTsv(
  values: unknown[][],
  range: { startRow: number; startCol: number; endRow: number; endCol: number },
  options: { maxRows: number; signal?: AbortSignal },
): string;

