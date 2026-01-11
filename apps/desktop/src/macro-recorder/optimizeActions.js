import { formatRangeAddress, parseCellAddress } from "../../../../packages/scripting/src/a1.js";

function cellKey(row, col) {
  return `${row},${col}`;
}

function buildDenseRect(
  actions,
) {
  if (actions.length < 2) return null;

  let minRow = Number.POSITIVE_INFINITY;
  let maxRow = Number.NEGATIVE_INFINITY;
  let minCol = Number.POSITIVE_INFINITY;
  let maxCol = Number.NEGATIVE_INFINITY;
  const values = new Map();

  for (const action of actions) {
    const { row, col } = parseCellAddress(action.address);
    minRow = Math.min(minRow, row);
    maxRow = Math.max(maxRow, row);
    minCol = Math.min(minCol, col);
    maxCol = Math.max(maxCol, col);
    values.set(cellKey(row, col), action.value);
  }

  const rows = maxRow - minRow + 1;
  const cols = maxCol - minCol + 1;
  if (rows * cols !== actions.length) return null;

  const matrix = [];
  for (let r = 0; r < rows; r++) {
    const row = [];
    for (let c = 0; c < cols; c++) {
      const v = values.get(cellKey(minRow + r, minCol + c));
      if (v === undefined) return null;
      row.push(v);
    }
    matrix.push(row);
  }

  return {
    type: "setRangeValues",
    sheetName: actions[0].sheetName,
    address: formatRangeAddress({ startRow: minRow, startCol: minCol, endRow: maxRow, endCol: maxCol }),
    values: matrix,
  };
}

function optimizeCellRuns(actions) {
  const rect = buildDenseRect(actions);
  if (rect) return [rect];

  // Fallback: merge per-row contiguous column segments.
  const byRow = new Map();
  for (const action of actions) {
    const coord = parseCellAddress(action.address);
    const bucket = byRow.get(coord.row) ?? [];
    bucket.push({ col: coord.col, action });
    byRow.set(coord.row, bucket);
  }

  const out = [];
  for (const [row, bucket] of [...byRow.entries()].sort(([a], [b]) => a - b)) {
    bucket.sort((a, b) => a.col - b.col);
    let start = 0;
    while (start < bucket.length) {
      let end = start + 1;
      while (end < bucket.length && bucket[end].col === bucket[end - 1].col + 1) end++;
      const segment = bucket.slice(start, end);
      if (segment.length === 1) {
        out.push(segment[0].action);
      } else {
        const startCol = segment[0].col;
        const endCol = segment[segment.length - 1].col;
        out.push({
          type: "setRangeValues",
          sheetName: segment[0].action.sheetName,
          address: formatRangeAddress({ startRow: row, startCol, endRow: row, endCol }),
          values: [segment.map((s) => s.action.value)],
        });
      }
      start = end;
    }
  }

  // Preserve original ordering across rows by sorting by (row, col).
  const ordering = new Map();
  actions.forEach((action, idx) => ordering.set(action.address, idx));
  out.sort((a, b) => {
    const aIdx = "address" in a ? ordering.get(a.address) ?? 0 : 0;
    const bIdx = "address" in b ? ordering.get(b.address) ?? 0 : 0;
    return aIdx - bIdx;
  });

  return out;
}

function collapseSelections(actions) {
  const out = [];
  for (const action of actions) {
    if (action.type === "setSelection") {
      const previous = out[out.length - 1];
      if (previous?.type === "setSelection" && previous.sheetName === action.sheetName) {
        out[out.length - 1] = action;
        continue;
      }
    }
    out.push(action);
  }
  return out;
}

export function optimizeMacroActions(actions) {
  const collapsedSelections = collapseSelections(actions);

  const out = [];
  let i = 0;
  while (i < collapsedSelections.length) {
    const action = collapsedSelections[i];
    if (action.type !== "setCellValue") {
      out.push(action);
      i += 1;
      continue;
    }

    const run = [action];
    let j = i + 1;
    while (j < collapsedSelections.length) {
      const next = collapsedSelections[j];
      if (next.type !== "setCellValue" || next.sheetName !== action.sheetName) break;
      run.push(next);
      j += 1;
    }

    out.push(...optimizeCellRuns(run));
    i = j;
  }

  return out;
}
