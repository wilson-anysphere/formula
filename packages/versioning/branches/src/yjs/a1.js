/**
 * Encode a zero-based row/col into an A1 address (e.g. row=0,col=0 -> "A1").
 *
 * @param {number} row
 * @param {number} col
 * @returns {string}
 */
export function rowColToA1(row, col) {
  if (!Number.isInteger(row) || row < 0) throw new Error(`Invalid row: ${row}`);
  if (!Number.isInteger(col) || col < 0) throw new Error(`Invalid col: ${col}`);

  let c = col + 1;
  let letters = "";
  while (c > 0) {
    const rem = (c - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    c = Math.floor((c - 1) / 26);
  }

  return `${letters}${row + 1}`;
}

/**
 * Decode an A1 address into a zero-based row/col.
 *
 * @param {string} a1
 * @returns {{ row: number, col: number }}
 */
export function a1ToRowCol(a1) {
  if (typeof a1 !== "string") throw new Error(`Invalid A1 address: ${String(a1)}`);
  const m = a1.match(/^\$?([A-Za-z]+)\$?(\d+)$/);
  if (!m) throw new Error(`Invalid A1 address: ${a1}`);

  const letters = m[1].toUpperCase();
  const rowOneBased = Number(m[2]);
  if (!Number.isInteger(rowOneBased) || rowOneBased <= 0) throw new Error(`Invalid A1 address: ${a1}`);

  let colOneBased = 0;
  for (let i = 0; i < letters.length; i += 1) {
    const code = letters.charCodeAt(i);
    if (code < 65 || code > 90) throw new Error(`Invalid A1 address: ${a1}`);
    colOneBased = colOneBased * 26 + (code - 64);
  }

  return { row: rowOneBased - 1, col: colOneBased - 1 };
}

