const CLASSIFICATION_LEVEL = Object.freeze({
  PUBLIC: "Public",
  INTERNAL: "Internal",
  CONFIDENTIAL: "Confidential",
  RESTRICTED: "Restricted",
});

const CLASSIFICATION_LEVELS = Object.freeze([
  CLASSIFICATION_LEVEL.PUBLIC,
  CLASSIFICATION_LEVEL.INTERNAL,
  CLASSIFICATION_LEVEL.CONFIDENTIAL,
  CLASSIFICATION_LEVEL.RESTRICTED,
]);

const DEFAULT_CLASSIFICATION = Object.freeze({
  level: CLASSIFICATION_LEVEL.PUBLIC,
  labels: [],
});

const DLP_ACTION = Object.freeze({
  SHARE_EXTERNAL_LINK: "sharing.externalLink",
  EXPORT_CSV: "export.csv",
  EXPORT_PDF: "export.pdf",
  EXPORT_XLSX: "export.xlsx",
  CLIPBOARD_COPY: "clipboard.copy",
  AI_CLOUD_PROCESSING: "ai.cloudProcessing",
  EXTERNAL_CONNECTOR: "connector.external",
});

const DLP_POLICY_VERSION = 1;

function isObject(value) {
  return typeof value === "object" && value !== null;
}

function validateLevel(level) {
  if (level === null) return;
  if (typeof level !== "string" || !CLASSIFICATION_LEVELS.includes(level)) {
    throw new Error(`Invalid classification level: ${String(level)}`);
  }
}

function normalizeLabels(labelsRaw) {
  const labels = Array.isArray(labelsRaw) ? labelsRaw : [];
  return [...new Set(labels.map((l) => String(l).trim()).filter(Boolean))].sort();
}

function normalizeClassification(classification) {
  if (!classification) return { ...DEFAULT_CLASSIFICATION };
  if (!isObject(classification)) throw new Error("Classification must be an object");
  validateLevel(classification.level);
  return { level: classification.level, labels: normalizeLabels(classification.labels) };
}

function classificationRank(level) {
  const idx = CLASSIFICATION_LEVELS.indexOf(level);
  if (idx === -1) throw new Error(`Unknown classification level: ${level}`);
  return idx;
}

function compareClassification(a, b) {
  const na = normalizeClassification(a);
  const nb = normalizeClassification(b);
  const ra = classificationRank(na.level);
  const rb = classificationRank(nb.level);
  if (ra === rb) return 0;
  return ra > rb ? 1 : -1;
}

function maxClassification(a, b) {
  const na = normalizeClassification(a);
  const nb = normalizeClassification(b);
  const level = classificationRank(na.level) >= classificationRank(nb.level) ? na.level : nb.level;
  const labels = [...new Set([...(na.labels || []), ...(nb.labels || [])])].sort();
  return { level, labels };
}

function ensureSupportedPolicyVersion(version) {
  if (version === DLP_POLICY_VERSION) return;
  if (version > DLP_POLICY_VERSION) {
    throw new Error(
      `Unsupported policy version: ${version}. This Formula build supports policy versions up to ${DLP_POLICY_VERSION}. Please upgrade your client/server.`
    );
  }
  throw new Error(
    `Unsupported policy version: ${version}. This Formula build supports policy version ${DLP_POLICY_VERSION}.`
  );
}

function validateDlpPolicy(policy) {
  normalizeDlpPolicy(policy);
}

function normalizeDlpPolicy(policy) {
  if (!isObject(policy)) throw new Error("Policy must be an object");
  if (!Number.isInteger(policy.version)) throw new Error("Policy.version must be an integer");
  ensureSupportedPolicyVersion(policy.version);

  if (typeof policy.allowDocumentOverrides !== "boolean") {
    throw new Error("Policy.allowDocumentOverrides must be a boolean");
  }
  if (!isObject(policy.rules)) throw new Error("Policy.rules must be an object");

  for (const [action, rule] of Object.entries(policy.rules)) {
    if (!isObject(rule)) throw new Error(`Rule for ${action} must be an object`);
    validateLevel(rule.maxAllowed);

    if (action === DLP_ACTION.AI_CLOUD_PROCESSING) {
      if (typeof rule.allowRestrictedContent !== "boolean") {
        throw new Error("AI rule allowRestrictedContent must be boolean");
      }
      if (typeof rule.redactDisallowed !== "boolean") {
        throw new Error("AI rule redactDisallowed must be boolean");
      }
    }
  }

  return policy;
}

function normalizeRange(range) {
  if (!range || !range.start || !range.end) throw new Error("Invalid range");
  const startRow = Math.min(range.start.row, range.end.row);
  const endRow = Math.max(range.start.row, range.end.row);
  const startCol = Math.min(range.start.col, range.end.col);
  const endCol = Math.max(range.start.col, range.end.col);
  return { start: { row: startRow, col: startCol }, end: { row: endRow, col: endCol } };
}

function normalizeNonEmptyString(value, name) {
  if (typeof value !== "string" || !value.trim()) {
    throw new Error(`${name} must be a non-empty string`);
  }
  return value;
}

function normalizeNonNegativeInt(value, name) {
  if (!Number.isInteger(value) || value < 0) {
    throw new Error(`${name} must be a non-negative integer`);
  }
  return value;
}

function normalizeCellCoord(coord, name) {
  if (!isObject(coord)) throw new Error(`${name} must be an object`);
  return {
    row: normalizeNonNegativeInt(coord.row, `${name}.row`),
    col: normalizeNonNegativeInt(coord.col, `${name}.col`),
  };
}

/**
 * Normalize and validate a DLP classification selector.
 *
 * Selector scopes:
 * - document: applies to the whole document
 * - sheet: applies to an entire sheet
 * - column: applies to a sheet column (0-based columnIndex) or table column (tableId + columnId)
 * - range: applies to a rectangular selection on a sheet
 * - cell: applies to an individual cell on a sheet
 *
 * The returned selector is canonicalized to keep range coordinates stable (start <= end).
 */
function normalizeSelector(selector) {
  if (!isObject(selector)) throw new Error("Selector must be an object");
  const scope = selector.scope;
  if (typeof scope !== "string") throw new Error("Selector.scope must be a string");

  const documentId = normalizeNonEmptyString(selector.documentId, "Selector.documentId");

  switch (scope) {
    case "document":
      return { scope, documentId };
    case "sheet": {
      const sheetId = normalizeNonEmptyString(selector.sheetId, "Selector.sheetId");
      return { scope, documentId, sheetId };
    }
    case "column": {
      const sheetId = normalizeNonEmptyString(selector.sheetId, "Selector.sheetId");
      const out = { scope, documentId, sheetId };

      const columnIndex =
        typeof selector.columnIndex === "number"
          ? normalizeNonNegativeInt(selector.columnIndex, "Selector.columnIndex")
          : null;
      const columnId = selector.columnId ? normalizeNonEmptyString(selector.columnId, "Selector.columnId") : null;
      const tableId = selector.tableId ? normalizeNonEmptyString(selector.tableId, "Selector.tableId") : null;

      if (columnIndex === null && columnId === null) {
        throw new Error("Column selector must include columnIndex or columnId");
      }

      if (columnIndex !== null) out.columnIndex = columnIndex;
      if (columnId !== null) out.columnId = columnId;
      if (tableId !== null) out.tableId = tableId;
      return out;
    }
    case "cell": {
      const sheetId = normalizeNonEmptyString(selector.sheetId, "Selector.sheetId");
      const row = normalizeNonNegativeInt(selector.row, "Selector.row");
      const col = normalizeNonNegativeInt(selector.col, "Selector.col");
      const out = { scope, documentId, sheetId, row, col };
      if (selector.tableId) out.tableId = normalizeNonEmptyString(selector.tableId, "Selector.tableId");
      if (selector.columnId) out.columnId = normalizeNonEmptyString(selector.columnId, "Selector.columnId");
      return out;
    }
    case "range": {
      const sheetId = normalizeNonEmptyString(selector.sheetId, "Selector.sheetId");
      if (!isObject(selector.range)) throw new Error("Selector.range must be an object");
      const normalized = normalizeRange({
        start: normalizeCellCoord(selector.range.start, "Selector.range.start"),
        end: normalizeCellCoord(selector.range.end, "Selector.range.end"),
      });
      return { scope, documentId, sheetId, range: normalized };
    }
    default:
      throw new Error(`Unknown selector scope: ${scope}`);
  }
}

function selectorKey(selector) {
  if (!isObject(selector)) throw new Error("Selector must be an object");
  const scope = selector.scope;
  if (typeof scope !== "string") throw new Error("Selector.scope must be a string");

  switch (scope) {
    case "document":
      return `document:${selector.documentId}`;
    case "sheet":
      return `sheet:${selector.documentId}:${selector.sheetId}`;
    case "column": {
      const tablePart = selector.tableId ? `:table:${selector.tableId}` : "";
      const colPart =
        typeof selector.columnIndex === "number"
          ? `:col:${selector.columnIndex}`
          : selector.columnId
            ? `:colId:${selector.columnId}`
            : "";
      if (!colPart) throw new Error("Column selector must include columnIndex or columnId");
      return `column:${selector.documentId}:${selector.sheetId}${tablePart}${colPart}`;
    }
    case "cell":
      return `cell:${selector.documentId}:${selector.sheetId}:${selector.row},${selector.col}`;
    case "range": {
      const normalized = normalizeRange(selector.range);
      return `range:${selector.documentId}:${selector.sheetId}:${normalized.start.row},${normalized.start.col}:${normalized.end.row},${normalized.end.col}`;
    }
    default:
      throw new Error(`Unknown selector scope: ${scope}`);
  }
}

function cellInRange(cell, range) {
  return (
    cell.row >= range.start.row &&
    cell.row <= range.end.row &&
    cell.col >= range.start.col &&
    cell.col <= range.end.col
  );
}

function rangesIntersect(a, b) {
  return a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col;
}

function selectorAppliesToCell(selector, cellRef) {
  if (!selector || typeof selector !== "object") return false;
  if (selector.documentId !== cellRef.documentId) return false;
  switch (selector.scope) {
    case "document":
      return true;
    case "sheet":
      return selector.sheetId === cellRef.sheetId;
    case "column": {
      if (selector.sheetId !== cellRef.sheetId) return false;
      if (typeof selector.columnIndex === "number") return selector.columnIndex === cellRef.col;
      if (selector.tableId && selector.columnId && cellRef.tableId && cellRef.columnId) {
        return selector.tableId === cellRef.tableId && selector.columnId === cellRef.columnId;
      }
      return false;
    }
    case "range":
      if (selector.sheetId !== cellRef.sheetId) return false;
      return cellInRange({ row: cellRef.row, col: cellRef.col }, selector.range);
    case "cell":
      return (
        selector.sheetId === cellRef.sheetId && selector.row === cellRef.row && selector.col === cellRef.col
      );
    default:
      return false;
  }
}

function selectorIntersectsRange(selector, rangeRef) {
  if (!selector || typeof selector !== "object") return false;
  if (selector.documentId !== rangeRef.documentId) return false;
  switch (selector.scope) {
    case "document":
      return true;
    case "sheet":
      return selector.sheetId === rangeRef.sheetId;
    case "column":
      if (selector.sheetId !== rangeRef.sheetId) return false;
      if (typeof selector.columnIndex !== "number") return false;
      return selector.columnIndex >= rangeRef.range.start.col && selector.columnIndex <= rangeRef.range.end.col;
    case "range":
      if (selector.sheetId !== rangeRef.sheetId) return false;
      return rangesIntersect(selector.range, rangeRef.range);
    case "cell":
      if (selector.sheetId !== rangeRef.sheetId) return false;
      return cellInRange({ row: selector.row, col: selector.col }, rangeRef.range);
    default:
      return false;
  }
}

function selectorSortRank(scope) {
  switch (scope) {
    case "cell":
      return 5;
    case "range":
      return 4;
    case "column":
      return 3;
    case "sheet":
      return 2;
    case "document":
      return 1;
    default:
      return 0;
  }
}

function safeStringCompare(a, b) {
  if (a === b) return 0;
  return a < b ? -1 : 1;
}

function compareMatchedSelectors(a, b) {
  const ra = selectorSortRank(a.selector.scope);
  const rb = selectorSortRank(b.selector.scope);
  if (ra !== rb) return rb - ra;

  // Same scope: apply stable, scope-specific ordering.
  const sa = a.selector;
  const sb = b.selector;

  switch (sa.scope) {
    case "document":
      return safeStringCompare(sa.documentId, sb.documentId);
    case "sheet":
      return safeStringCompare(sa.sheetId, sb.sheetId);
    case "column": {
      const sheetCmp = safeStringCompare(sa.sheetId, sb.sheetId);
      if (sheetCmp !== 0) return sheetCmp;

      const aHasIdx = typeof sa.columnIndex === "number";
      const bHasIdx = typeof sb.columnIndex === "number";
      if (aHasIdx && bHasIdx && sa.columnIndex !== sb.columnIndex) return sa.columnIndex - sb.columnIndex;
      if (aHasIdx !== bHasIdx) return aHasIdx ? -1 : 1;

      const tableCmp = safeStringCompare(sa.tableId ?? "", sb.tableId ?? "");
      if (tableCmp !== 0) return tableCmp;
      const colCmp = safeStringCompare(sa.columnId ?? "", sb.columnId ?? "");
      if (colCmp !== 0) return colCmp;
      break;
    }
    case "range": {
      const sheetCmp = safeStringCompare(sa.sheetId, sb.sheetId);
      if (sheetCmp !== 0) return sheetCmp;

      const aArea = (sa.range.end.row - sa.range.start.row + 1) * (sa.range.end.col - sa.range.start.col + 1);
      const bArea = (sb.range.end.row - sb.range.start.row + 1) * (sb.range.end.col - sb.range.start.col + 1);
      if (aArea !== bArea) return aArea - bArea;

      if (sa.range.start.row !== sb.range.start.row) return sa.range.start.row - sb.range.start.row;
      if (sa.range.start.col !== sb.range.start.col) return sa.range.start.col - sb.range.start.col;
      if (sa.range.end.row !== sb.range.end.row) return sa.range.end.row - sb.range.end.row;
      if (sa.range.end.col !== sb.range.end.col) return sa.range.end.col - sb.range.end.col;
      break;
    }
    case "cell": {
      const sheetCmp = safeStringCompare(sa.sheetId, sb.sheetId);
      if (sheetCmp !== 0) return sheetCmp;
      if (sa.row !== sb.row) return sa.row - sb.row;
      if (sa.col !== sb.col) return sa.col - sb.col;
      break;
    }
    default:
      break;
  }

  return safeStringCompare(a.selectorKey, b.selectorKey);
}

const DEFAULT_MAX_MATCHED_SELECTORS = Number.POSITIVE_INFINITY;
const DEFAULT_MAX_RANGE_CELLS_FOR_MATCHED_SELECTORS = 1_000_000;

function rangeCellCount(range) {
  return (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
}

/**
 * Resolve an "effective" classification for a query selector (cell or range) given a set of
 * classification records.
 *
 * Semantics:
 * - Matching selectors are determined by scope + overlap rules (see `selectorAppliesToCell`
 *   and `selectorIntersectsRange`).
 * - The effective classification is the maximum classification level across all matched
 *   selectors, with labels unioned across all matched selectors.
 *
 * This is intentionally conservative for DLP enforcement: a more specific selector cannot
 * weaken a broader restriction (e.g. a Public cell inside a Restricted range remains Restricted).
 */
function resolveClassification({ querySelector, records, options = {} }) {
  const normalizedQuery = normalizeSelector(querySelector);
  if (normalizedQuery.scope !== "cell" && normalizedQuery.scope !== "range") {
    throw new Error("Query selector must have scope 'cell' or 'range'");
  }

  const includeMatchedSelectors = Boolean(options.includeMatchedSelectors);
  const maxMatchedSelectors =
    Number.isInteger(options.maxMatchedSelectors) && options.maxMatchedSelectors >= 0
      ? options.maxMatchedSelectors
      : DEFAULT_MAX_MATCHED_SELECTORS;
  const maxRangeCellsForMatchedSelectors =
    Number.isInteger(options.maxRangeCellsForMatchedSelectors) && options.maxRangeCellsForMatchedSelectors >= 0
      ? options.maxRangeCellsForMatchedSelectors
      : DEFAULT_MAX_RANGE_CELLS_FOR_MATCHED_SELECTORS;

  if (normalizedQuery.scope === "range") {
    const cells = rangeCellCount(normalizedQuery.range);
    if (includeMatchedSelectors && cells > maxRangeCellsForMatchedSelectors) {
      throw new Error(
        `Range too large to include matched selectors (${cells} cells > ${maxRangeCellsForMatchedSelectors})`
      );
    }
  }

  const docEntries = [];
  const sheetEntriesBySheet = new Map();
  const columnEntriesBySheet = new Map();
  const rangeEntriesBySheet = new Map();
  const cellEntriesBySheet = new Map();

  /**
   * @param {Map<string, any[]>} map
   * @param {string} sheetId
   * @param {{selector:any, selectorKey:string, classification:any}} entry
   */
  function addBySheet(map, sheetId, entry) {
    const existing = map.get(sheetId);
    if (existing) existing.push(entry);
    else map.set(sheetId, [entry]);
  }

  for (const record of records || []) {
    if (!record || !record.selector) continue;
    let selector;
    try {
      selector = normalizeSelector(record.selector);
    } catch {
      // Ignore invalid persisted selectors so one bad row can't break enforcement.
      continue;
    }

    const entry = {
      selector,
      selectorKey: selectorKey(selector),
      classification: normalizeClassification(record.classification),
    };

    switch (selector.scope) {
      case "document":
        docEntries.push(entry);
        break;
      case "sheet":
        addBySheet(sheetEntriesBySheet, selector.sheetId, entry);
        break;
      case "column":
        addBySheet(columnEntriesBySheet, selector.sheetId, entry);
        break;
      case "range":
        addBySheet(rangeEntriesBySheet, selector.sheetId, entry);
        break;
      case "cell":
        addBySheet(cellEntriesBySheet, selector.sheetId, entry);
        break;
      default:
        break;
    }
  }

  let effectiveClassification = { ...DEFAULT_CLASSIFICATION };
  let matchedCount = 0;
  const matchedSelectors = [];

  function consider(entry) {
    matchedCount++;
    effectiveClassification = maxClassification(effectiveClassification, entry.classification);
    if (includeMatchedSelectors && matchedSelectors.length < maxMatchedSelectors) {
      matchedSelectors.push(entry);
    }
  }

  const sheetId = normalizedQuery.sheetId;

  // Document-level selectors always apply for same-document queries.
  for (const entry of docEntries) {
    if (entry.selector.documentId === normalizedQuery.documentId) consider(entry);
  }

  if (normalizedQuery.scope === "cell") {
    const cellRef = normalizedQuery;

    for (const entry of sheetEntriesBySheet.get(sheetId) || []) {
      if (selectorAppliesToCell(entry.selector, cellRef)) consider(entry);
    }
    for (const entry of columnEntriesBySheet.get(sheetId) || []) {
      if (selectorAppliesToCell(entry.selector, cellRef)) consider(entry);
    }
    for (const entry of rangeEntriesBySheet.get(sheetId) || []) {
      if (selectorAppliesToCell(entry.selector, cellRef)) consider(entry);
    }
    for (const entry of cellEntriesBySheet.get(sheetId) || []) {
      if (selectorAppliesToCell(entry.selector, cellRef)) consider(entry);
    }
  } else {
    const rangeRef = normalizedQuery;
    for (const entry of sheetEntriesBySheet.get(sheetId) || []) {
      if (selectorIntersectsRange(entry.selector, rangeRef)) consider(entry);
    }
    for (const entry of columnEntriesBySheet.get(sheetId) || []) {
      if (selectorIntersectsRange(entry.selector, rangeRef)) consider(entry);
    }
    for (const entry of rangeEntriesBySheet.get(sheetId) || []) {
      if (selectorIntersectsRange(entry.selector, rangeRef)) consider(entry);
    }
    for (const entry of cellEntriesBySheet.get(sheetId) || []) {
      if (selectorIntersectsRange(entry.selector, rangeRef)) consider(entry);
    }
  }

  if (includeMatchedSelectors) matchedSelectors.sort(compareMatchedSelectors);

  return {
    effectiveClassification,
    matchedCount,
    matchedSelectors: includeMatchedSelectors ? matchedSelectors : undefined,
  };
}

const DLP_DECISION = Object.freeze({
  ALLOW: "allow",
  BLOCK: "block",
  REDACT: "redact",
});

const DLP_REASON_CODE = Object.freeze({
  BLOCKED_BY_POLICY: "dlp.blockedByPolicy",
  INVALID_POLICY: "dlp.invalidPolicy",
});

function ruleForAction(policy, action) {
  if (!policy || typeof policy !== "object" || !policy.rules) {
    throw new Error("Invalid policy object");
  }
  return policy.rules[action] || { maxAllowed: null };
}

function evaluatePolicy({ action, classification, policy, options = {} }) {
  const normalized = normalizeClassification(classification);
  const rule = ruleForAction(policy, action);

  const maxAllowed = rule.maxAllowed ?? null;
  if (maxAllowed === null) {
    return {
      action,
      decision: DLP_DECISION.BLOCK,
      reasonCode: DLP_REASON_CODE.BLOCKED_BY_POLICY,
      classification: normalized,
      maxAllowed,
    };
  }

  const overThreshold = compareClassification(normalized, { level: maxAllowed, labels: [] }) === 1;

  if (action === DLP_ACTION.AI_CLOUD_PROCESSING) {
    if (normalized.level === CLASSIFICATION_LEVEL.RESTRICTED && options.includeRestrictedContent) {
      if (!rule.allowRestrictedContent) {
        return {
          action,
          decision: DLP_DECISION.BLOCK,
          reasonCode: DLP_REASON_CODE.BLOCKED_BY_POLICY,
          classification: normalized,
          maxAllowed,
        };
      }
      return {
        action,
        decision: DLP_DECISION.ALLOW,
        classification: normalized,
        maxAllowed,
      };
    }

    if (!overThreshold) {
      return {
        action,
        decision: DLP_DECISION.ALLOW,
        classification: normalized,
        maxAllowed,
      };
    }

    if (rule.redactDisallowed) {
      return {
        action,
        decision: DLP_DECISION.REDACT,
        reasonCode: DLP_REASON_CODE.BLOCKED_BY_POLICY,
        classification: normalized,
        maxAllowed,
      };
    }

    return {
      action,
      decision: DLP_DECISION.BLOCK,
      reasonCode: DLP_REASON_CODE.BLOCKED_BY_POLICY,
      classification: normalized,
      maxAllowed,
    };
  }

  if (!overThreshold) {
    return {
      action,
      decision: DLP_DECISION.ALLOW,
      classification: normalized,
      maxAllowed,
    };
  }

  return {
    action,
    decision: DLP_DECISION.BLOCK,
    reasonCode: DLP_REASON_CODE.BLOCKED_BY_POLICY,
    classification: normalized,
    maxAllowed,
  };
}

function isClassificationAllowed(classification, maxAllowed) {
  if (maxAllowed === null) return false;
  const level = normalizeClassification(classification).level;
  return classificationRank(level) <= classificationRank(maxAllowed);
}

function isAllowed(action, classification, policy, options) {
  const decision = evaluatePolicy({ action, classification, policy, options });
  return decision.decision === DLP_DECISION.ALLOW || decision.decision === DLP_DECISION.REDACT;
}

const REDACTION_PLACEHOLDER = "[REDACTED]";

function redact() {
  return REDACTION_PLACEHOLDER;
}

module.exports = {
  CLASSIFICATION_LEVEL,
  CLASSIFICATION_LEVELS,
  DEFAULT_CLASSIFICATION,
  DLP_ACTION,
  DLP_POLICY_VERSION,
  DLP_DECISION,
  DLP_REASON_CODE,
  REDACTION_PLACEHOLDER,
  classificationRank,
  compareClassification,
  evaluatePolicy,
  isAllowed,
  isClassificationAllowed,
  maxClassification,
  normalizeClassification,
  normalizeDlpPolicy,
  normalizeRange,
  normalizeSelector,
  redact,
  resolveClassification,
  selectorKey,
  validateDlpPolicy,
};
