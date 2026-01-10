export function columnNumberToLetters(colNumber) {
  if (!Number.isInteger(colNumber) || colNumber <= 0) {
    throw new Error(`Invalid column number: ${colNumber}`);
  }

  let col = colNumber;
  let letters = "";
  while (col > 0) {
    const remainder = (col - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    col = Math.floor((col - 1) / 26);
  }
  return letters;
}

export function columnLettersToNumber(colLetters) {
  const letters = String(colLetters || "").toUpperCase();
  if (!/^[A-Z]+$/.test(letters)) {
    throw new Error(`Invalid column letters: ${colLetters}`);
  }

  let colNumber = 0;
  for (const ch of letters) {
    colNumber = colNumber * 26 + (ch.charCodeAt(0) - 64);
  }
  return colNumber;
}

export function normalizeA1Address(address) {
  const str = String(address || "").trim().toUpperCase();
  if (!/^[A-Z]+[0-9]+$/.test(str)) {
    throw new Error(`Invalid A1 address: ${address}`);
  }
  return str;
}

export function a1ToRowCol(address) {
  const a1 = normalizeA1Address(address);
  const match = /^(?<col>[A-Z]+)(?<row>[0-9]+)$/.exec(a1);
  const row = Number(match.groups.row);
  const col = columnLettersToNumber(match.groups.col);
  return { row, col };
}

export function rowColToA1(row, col) {
  if (!Number.isInteger(row) || row <= 0) {
    throw new Error(`Invalid row: ${row}`);
  }
  const letters = columnNumberToLetters(col);
  return `${letters}${row}`;
}

