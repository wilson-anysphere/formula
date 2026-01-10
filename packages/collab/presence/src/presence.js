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
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);

  return { startRow, startCol, endRow, endCol };
}

export function serializePresenceState(state) {
  const selections = Array.isArray(state.selections)
    ? state.selections.map((range) => normalizeRange(range))
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
      const startRow = toInt(selection.startRow);
      const startCol = toInt(selection.startCol);
      const endRow = toInt(selection.endRow);
      const endCol = toInt(selection.endCol);
      if (startRow === null || startCol === null || endRow === null || endCol === null) return null;
      selections.push(normalizeRange({ startRow, startCol, endRow, endCol }));
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

