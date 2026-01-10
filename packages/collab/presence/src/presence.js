export const PRESENCE_VERSION = 1;

function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function isFiniteNumber(value) {
  return typeof value === "number" && Number.isFinite(value);
}

function toInt(value) {
  if (!isFiniteNumber(value)) return null;
  return Math.trunc(value);
}

function normalizeRange(range) {
  if (!isRecord(range)) return null;

  // Support two common range shapes:
  // - { startRow, startCol, endRow, endCol } (wire format)
  // - { start: { row, col }, end: { row, col } } (local UI format)

  let startRow;
  let startCol;
  let endRow;
  let endCol;

  if (
    isFiniteNumber(range.startRow) &&
    isFiniteNumber(range.startCol) &&
    isFiniteNumber(range.endRow) &&
    isFiniteNumber(range.endCol)
  ) {
    startRow = range.startRow;
    startCol = range.startCol;
    endRow = range.endRow;
    endCol = range.endCol;
  } else if (
    isRecord(range.start) &&
    isRecord(range.end) &&
    isFiniteNumber(range.start.row) &&
    isFiniteNumber(range.start.col) &&
    isFiniteNumber(range.end.row) &&
    isFiniteNumber(range.end.col)
  ) {
    startRow = range.start.row;
    startCol = range.start.col;
    endRow = range.end.row;
    endCol = range.end.col;
  } else {
    return null;
  }

  const normalizedStartRow = Math.min(startRow, endRow);
  const normalizedEndRow = Math.max(startRow, endRow);
  const normalizedStartCol = Math.min(startCol, endCol);
  const normalizedEndCol = Math.max(startCol, endCol);

  return {
    startRow: Math.trunc(normalizedStartRow),
    startCol: Math.trunc(normalizedStartCol),
    endRow: Math.trunc(normalizedEndRow),
    endCol: Math.trunc(normalizedEndCol),
  };
}

export function serializePresenceState(state) {
  const selections = Array.isArray(state.selections)
    ? state.selections.map((range) => normalizeRange(range)).filter((range) => range !== null)
    : [];

  const cursor =
    state.cursor && isFiniteNumber(state.cursor.row) && isFiniteNumber(state.cursor.col)
      ? { row: Math.trunc(state.cursor.row), col: Math.trunc(state.cursor.col) }
      : null;

  return {
    v: PRESENCE_VERSION,
    id: state.id,
    name: state.name,
    color: state.color,
    sheet: state.activeSheet,
    cursor,
    selections,
    lastActive: state.lastActive,
  };
}

export function deserializePresenceState(payload) {
  if (!isRecord(payload)) return null;
  if (payload.v !== PRESENCE_VERSION) return null;

  if (typeof payload.id !== "string" || payload.id.length === 0) return null;
  if (typeof payload.name !== "string") return null;
  if (typeof payload.color !== "string" || payload.color.length === 0) return null;
  if (typeof payload.sheet !== "string" || payload.sheet.length === 0) return null;

  const lastActive = isFiniteNumber(payload.lastActive) ? payload.lastActive : null;
  if (lastActive === null) return null;

  let cursor = null;
  if (payload.cursor !== null && payload.cursor !== undefined) {
    if (!isRecord(payload.cursor)) return null;
    const row = toInt(payload.cursor.row);
    const col = toInt(payload.cursor.col);
    if (row === null || col === null) return null;
    cursor = { row, col };
  }

  const selections = [];
  if (payload.selections !== null && payload.selections !== undefined) {
    if (!Array.isArray(payload.selections)) return null;
    for (const selection of payload.selections) {
      if (!isRecord(selection)) return null;

      // Accept both range shapes for tolerance (see `normalizeRange`).
      const normalized = normalizeRange(selection);
      if (!normalized) return null;
      selections.push(normalized);
    }
  }

  return {
    id: payload.id,
    name: payload.name,
    color: payload.color,
    activeSheet: payload.sheet,
    cursor,
    selections,
    lastActive,
  };
}
