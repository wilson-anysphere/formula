export function columnLettersToIndex(letters) {
  let col = 0;
  for (const ch of letters.toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) return null;
    col = col * 26 + (code - 64);
  }
  return col - 1;
}

export function parseCellRef(cell) {
  const match = /^\$?([A-Z]+)\$?([0-9]+)$/.exec(cell.toUpperCase());
  if (!match) return null;
  const col = columnLettersToIndex(match[1]);
  if (col == null) return null;
  const row = Number.parseInt(match[2], 10) - 1;
  if (!Number.isFinite(row) || row < 0) return null;
  return { col, row };
}

function parseSheetAndRef(rangeRef) {
  const match = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(rangeRef);
  if (!match) return { sheetName: undefined, ref: rangeRef };

  const rawSheet = match[1] ?? match[2];
  const sheetName = rawSheet ? rawSheet.replace(/''/g, "'") : undefined;
  return { sheetName, ref: match[3] };
}

export function parseA1Range(rangeRef) {
  const { sheetName, ref } = parseSheetAndRef(rangeRef.trim());
  const [startRef, endRef] = ref.split(":", 2);
  const start = parseCellRef(startRef);
  const end = parseCellRef(endRef ?? startRef);
  if (!start || !end) return null;

  return {
    sheetName,
    startCol: Math.min(start.col, end.col),
    startRow: Math.min(start.row, end.row),
    endCol: Math.max(start.col, end.col),
    endRow: Math.max(start.row, end.row),
  };
}
