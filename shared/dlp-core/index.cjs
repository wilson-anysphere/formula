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
  redact,
  selectorKey,
  validateDlpPolicy,
};

