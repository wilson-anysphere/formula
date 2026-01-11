export const DOCUMENT_ROLES = Object.freeze([
  "owner",
  "admin",
  "editor",
  "commenter",
  "viewer"
]);

const ROLE_SET = new Set(DOCUMENT_ROLES);

export function assertValidRole(role) {
  if (!ROLE_SET.has(role)) {
    throw new Error(`Invalid role "${role}"`);
  }
}

export const ROLE_CAPABILITIES = Object.freeze({
  owner: Object.freeze({ read: true, edit: true, comment: true, share: true }),
  admin: Object.freeze({ read: true, edit: true, comment: true, share: true }),
  editor: Object.freeze({ read: true, edit: true, comment: true, share: false }),
  commenter: Object.freeze({ read: true, edit: false, comment: true, share: false }),
  viewer: Object.freeze({ read: true, edit: false, comment: false, share: false })
});

export function roleCanRead(role) {
  return Boolean(ROLE_CAPABILITIES[role]?.read);
}

export function roleCanEdit(role) {
  return Boolean(ROLE_CAPABILITIES[role]?.edit);
}

export function roleCanComment(role) {
  return Boolean(ROLE_CAPABILITIES[role]?.comment);
}

export function roleCanShare(role) {
  return Boolean(ROLE_CAPABILITIES[role]?.share);
}

export function normalizeRange(range) {
  if (!range || typeof range !== "object") {
    throw new Error("range must be an object");
  }
  const sheetId = range.sheetId ?? range.sheetName ?? "Sheet1";
  const startRow = Number(range.startRow);
  const endRow = Number(range.endRow);
  const startCol = Number(range.startCol);
  const endCol = Number(range.endCol);

  for (const [name, value] of [
    ["startRow", startRow],
    ["endRow", endRow],
    ["startCol", startCol],
    ["endCol", endCol]
  ]) {
    if (!Number.isInteger(value) || value < 0) {
      throw new Error(`range.${name} must be a non-negative integer`);
    }
  }

  if (endRow < startRow) throw new Error("range.endRow must be >= startRow");
  if (endCol < startCol) throw new Error("range.endCol must be >= startCol");

  return { sheetId, startRow, endRow, startCol, endCol };
}

export function cellInRange(cell, range) {
  return (
    cell.sheetId === range.sheetId &&
    cell.row >= range.startRow &&
    cell.row <= range.endRow &&
    cell.col >= range.startCol &&
    cell.col <= range.endCol
  );
}

export function normalizeRestriction(restriction) {
  if (!restriction || typeof restriction !== "object") {
    throw new Error("restriction must be an object");
  }

  // The API returns a flattened shape for range restrictions:
  // `{ sheetName, startRow, startCol, endRow, endCol, readAllowlist, editAllowlist }`
  // while older clients may send `{ range: { ... }, readAllowlist, editAllowlist }`.
  const range = normalizeRange(restriction.range ?? restriction);

  let readAllowlist = undefined;
  if (restriction.readAllowlist !== undefined) {
    if (!Array.isArray(restriction.readAllowlist)) {
      throw new Error("restriction.readAllowlist must be an array when provided");
    }
    const normalized = Array.from(new Set(restriction.readAllowlist.map(String)));
    // API payloads always include allowlist arrays (often empty). Treat an empty
    // list as "no restriction" so edit-only restrictions don't implicitly deny reads.
    readAllowlist = normalized.length > 0 ? normalized : undefined;
  }

  let editAllowlist = undefined;
  if (restriction.editAllowlist !== undefined) {
    if (!Array.isArray(restriction.editAllowlist)) {
      throw new Error("restriction.editAllowlist must be an array when provided");
    }
    const normalized = Array.from(new Set(restriction.editAllowlist.map(String)));
    editAllowlist = normalized.length > 0 ? normalized : undefined;
  }

  return {
    id: restriction.id ?? createId(),
    range,
    readAllowlist,
    editAllowlist,
    createdAt: restriction.createdAt ? new Date(restriction.createdAt) : new Date()
  };
}

function restrictionAllows(userId, allowlist) {
  if (allowlist === undefined) return true;
  if (!userId) return false;
  return allowlist.includes(userId);
}

export function getCellPermissions({ role, restrictions, userId, cell }) {
  if (!roleCanRead(role)) {
    return { canRead: false, canEdit: false };
  }

  const normalizedCell = {
    sheetId: cell.sheetId ?? cell.sheetName ?? "Sheet1",
    row: Number(cell.row),
    col: Number(cell.col)
  };

  if (
    !Number.isInteger(normalizedCell.row) ||
    normalizedCell.row < 0 ||
    !Number.isInteger(normalizedCell.col) ||
    normalizedCell.col < 0
  ) {
    throw new Error("cell.row and cell.col must be non-negative integers");
  }

  let canRead = true;
  let canEdit = roleCanEdit(role);

  for (const restriction of restrictions ?? []) {
    const normalizedRestriction = normalizeRestriction(restriction);
    if (!cellInRange(normalizedCell, normalizedRestriction.range)) continue;

    if (!restrictionAllows(userId, normalizedRestriction.readAllowlist)) {
      canRead = false;
      canEdit = false;
      break;
    }

    if (!restrictionAllows(userId, normalizedRestriction.editAllowlist)) {
      canEdit = false;
    }
  }

  if (!canRead) return { canRead: false, canEdit: false };
  return { canRead, canEdit };
}

export function maskCellValue(value) {
  // Enterprise masking: never leak even value length.
  return "###";
}

export function maskValueIfUnreadable({ value, canRead }) {
  if (canRead) return value;
  return maskCellValue(value);
}

export function maskCellUpdatesForUser({ role, restrictions, userId, updates }) {
  return updates.map((update) => {
    const { canRead } = getCellPermissions({
      role,
      restrictions,
      userId,
      cell: update.cell
    });

    return {
      ...update,
      value: maskValueIfUnreadable({ value: update.value, canRead })
    };
  });
}

function createId() {
  const globalCrypto = globalThis.crypto;
  if (globalCrypto?.randomUUID) {
    return globalCrypto.randomUUID();
  }
  return `perm_${Math.random().toString(16).slice(2)}_${Date.now()}`;
}
