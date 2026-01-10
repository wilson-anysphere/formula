export const CLASSIFICATION_LEVELS = ["Public", "Internal", "Confidential", "Restricted"] as const;

export type ClassificationLevel = (typeof CLASSIFICATION_LEVELS)[number];

export interface Classification {
  level: ClassificationLevel;
  labels?: string[];
}

export interface DlpRuleBase {
  maxAllowed: ClassificationLevel | null;
}

export interface DlpAiRule extends DlpRuleBase {
  allowRestrictedContent?: boolean;
  redactDisallowed?: boolean;
}

export interface DlpPolicy {
  version: number;
  allowDocumentOverrides: boolean;
  rules: Record<string, DlpRuleBase | DlpAiRule>;
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function validateLevel(level: unknown): asserts level is ClassificationLevel | null {
  if (level === null) return;
  if (typeof level !== "string" || !CLASSIFICATION_LEVELS.includes(level as ClassificationLevel)) {
    throw new Error(`Invalid classification level: ${String(level)}`);
  }
}

export function validateDlpPolicy(policy: unknown): asserts policy is DlpPolicy {
  if (!isObject(policy)) throw new Error("Policy must be an object");
  if (!Number.isInteger(policy.version)) throw new Error("Policy.version must be an integer");
  if (typeof policy.allowDocumentOverrides !== "boolean") {
    throw new Error("Policy.allowDocumentOverrides must be a boolean");
  }
  if (!isObject(policy.rules)) throw new Error("Policy.rules must be an object");

  for (const [action, rule] of Object.entries(policy.rules)) {
    if (!isObject(rule)) throw new Error(`Rule for ${action} must be an object`);
    validateLevel(rule.maxAllowed);

    if (action === "ai.cloudProcessing") {
      if (typeof rule.allowRestrictedContent !== "boolean") {
        throw new Error("AI rule allowRestrictedContent must be boolean");
      }
      if (typeof rule.redactDisallowed !== "boolean") {
        throw new Error("AI rule redactDisallowed must be boolean");
      }
    }
  }
}

export function normalizeClassification(classification: unknown): Classification {
  if (!isObject(classification)) throw new Error("Classification must be an object");
  validateLevel(classification.level);
  const labelsRaw = Array.isArray(classification.labels) ? classification.labels : [];
  const labels = [...new Set(labelsRaw.map((l) => String(l).trim()).filter(Boolean))].sort();
  return { level: classification.level as ClassificationLevel, labels };
}

type CellCoord = { row: number; col: number };
type CellRange = { start: CellCoord; end: CellCoord };

function normalizeRange(range: CellRange): CellRange {
  if (!range || !range.start || !range.end) throw new Error("Invalid range");
  const startRow = Math.min(range.start.row, range.end.row);
  const endRow = Math.max(range.start.row, range.end.row);
  const startCol = Math.min(range.start.col, range.end.col);
  const endCol = Math.max(range.start.col, range.end.col);
  return { start: { row: startRow, col: startCol }, end: { row: endRow, col: endCol } };
}

/**
 * Stable key used to address a classification record for upserts/deletes.
 *
 * Mirrors the `selectorKey` algorithm in `packages/security/dlp/src/selectors.js` so the
 * client and server agree on identifiers.
 */
export function selectorKey(selector: unknown): string {
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
      const range = normalizeRange(selector.range as CellRange);
      return `range:${selector.documentId}:${selector.sheetId}:${range.start.row},${range.start.col}:${range.end.row},${range.end.col}`;
    }
    default:
      throw new Error(`Unknown selector scope: ${scope}`);
  }
}

