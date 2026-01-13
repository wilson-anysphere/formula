/**
 * @param {{ r0: number, c0: number, r1: number, c1: number }} rect
 */
export function rectSize(rect) {
  return (rect.r1 - rect.r0 + 1) * (rect.c1 - rect.c0 + 1);
}

/**
 * @param {{ r0: number, c0: number, r1: number, c1: number }} a
 * @param {{ r0: number, c0: number, r1: number, c1: number }} b
 */
export function rectIntersectionArea(a, b) {
  const r0 = Math.max(a.r0, b.r0);
  const c0 = Math.max(a.c0, b.c0);
  const r1 = Math.min(a.r1, b.r1);
  const c1 = Math.min(a.c1, b.c1);
  if (r1 < r0 || c1 < c0) return 0;
  return (r1 - r0 + 1) * (c1 - c0 + 1);
}

/**
 * Convert a zero-based (row, col) coordinate to an A1 cell address.
 *
 * @param {number} row0
 * @param {number} col0
 */
export function cellToA1(row0, col0) {
  function colToLetters(colIdx0) {
    let col = colIdx0 + 1;
    let s = "";
    while (col > 0) {
      const mod = (col - 1) % 26;
      s = String.fromCharCode(65 + mod) + s;
      col = Math.floor((col - 1) / 26);
    }
    return s;
  }

  return `${colToLetters(col0)}${row0 + 1}`;
}

/**
 * @param {{ r0: number, c0: number, r1: number, c1: number }} rect
 */
export function rectToA1(rect) {
  const a = cellToA1(rect.r0, rect.c0);
  const b = cellToA1(rect.r1, rect.c1);
  return a === b ? a : `${a}:${b}`;
}
