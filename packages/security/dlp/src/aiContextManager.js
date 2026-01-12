import { DLP_ACTION } from "./actions.js";
import { evaluatePolicy, DLP_DECISION } from "./policyEngine.js";
import { effectiveCellClassification, effectiveRangeClassification, normalizeRange } from "./selectors.js";
import { DlpViolationError } from "./errors.js";
import { CLASSIFICATION_LEVEL, DEFAULT_CLASSIFICATION, classificationRank, maxClassification } from "./classification.js";
import dlpCore from "./core.js";

const { redact } = dlpCore;

const DEFAULT_CLASSIFICATION_RANK = classificationRank(CLASSIFICATION_LEVEL.PUBLIC);
const RESTRICTED_CLASSIFICATION_RANK = classificationRank(CLASSIFICATION_LEVEL.RESTRICTED);

/**
 * Build AI context from spreadsheet cells while respecting DLP classification.
 *
 * This module focuses on cloud LLM processing controls:
 * - Restricted cells are redacted by default (never sent to the cloud).
 * - Attempts to explicitly include Restricted content are blocked unless the policy
 *   allows it.
 * - Every request is audited with the classification + redaction decisions.
 */
export class AiContextManager {
  /**
   * @param {{
   *  classificationStore: {list(documentId:string): Array<{selector:any, classification:any}>},
   *  auditLogger: {log(event:any): void}
   * }} params
   */
  constructor({ classificationStore, auditLogger }) {
    if (!classificationStore) throw new Error("AiContextManager requires classificationStore");
    if (!auditLogger) throw new Error("AiContextManager requires auditLogger");
    this.classificationStore = classificationStore;
    this.auditLogger = auditLogger;
  }

  /**
   * @param {{
   *  documentId: string,
   *  sheetId: string,
   *  range: {start:{row:number,col:number}, end:{row:number,col:number}},
   *  cells: Array<{row:number,col:number,value:any}>,
   *  policy: any,
   *  includeRestrictedContent?: boolean
   * }} params
   */
  buildCloudContext({ documentId, sheetId, range, cells, policy, includeRestrictedContent = false }) {
    const records = this.classificationStore.list(documentId);
    const selectionClassification = effectiveRangeClassification({ documentId, sheetId, range }, records);

    const evaluation = evaluatePolicy({
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      classification: selectionClassification,
      policy,
      options: { includeRestrictedContent },
    });

    if (evaluation.decision === DLP_DECISION.BLOCK) {
      this.auditLogger.log({
        type: "ai.request",
        documentId,
        sheetId,
        range,
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        decision: evaluation,
        redactedCellCount: 0,
      });
      throw new DlpViolationError(evaluation);
    }

    const normalizedRange = normalizeRange(range);
    const byRowCol = new Map();
    for (const cell of cells) {
      byRowCol.set(`${cell.row},${cell.col}`, cell.value);
    }

    // Fast path: if the selection is allowed as a whole, no per-cell DLP checks are required.
    // (If any cell were disallowed, `selectionClassification` would exceed the threshold.)
    if (evaluation.decision === DLP_DECISION.ALLOW) {
      const lines = [];
      for (let row = normalizedRange.start.row; row <= normalizedRange.end.row; row++) {
        const rowValues = [];
        for (let col = normalizedRange.start.col; col <= normalizedRange.end.col; col++) {
          const value = byRowCol.get(`${row},${col}`);
          rowValues.push(value === undefined || value === null ? "" : String(value));
        }
        lines.push(rowValues.join("\t"));
      }

      const context = lines.join("\n");

      this.auditLogger.log({
        type: "ai.request",
        documentId,
        sheetId,
        range,
        action: DLP_ACTION.AI_CLOUD_PROCESSING,
        decision: evaluation,
        selectionClassification,
        redactedCellCount: 0,
        redactions: [],
      });

      return {
        context,
        selectionClassification,
        evaluation,
        redactions: [],
      };
    }

    const maxAllowedRank = evaluation.maxAllowed === null ? null : classificationRank(evaluation.maxAllowed);
    const policyAllowsRestrictedContent = Boolean(policy?.rules?.[DLP_ACTION.AI_CLOUD_PROCESSING]?.allowRestrictedContent);
    const restrictedAllowed = includeRestrictedContent
      ? policyAllowsRestrictedContent
      : maxAllowedRank !== null && maxAllowedRank >= RESTRICTED_CLASSIFICATION_RANK;

    const index = buildDlpRangeIndex({ documentId, sheetId, range: normalizedRange }, records, {
      // Under REDACT decisions, maxAllowed is always non-null, but keep this defensive.
      maxAllowedRank: maxAllowedRank ?? DEFAULT_CLASSIFICATION_RANK,
    });

    const redactions = [];
    let redactedCount = 0;

    // Render as a small TSV-like snippet. This matches many LLM prompt conventions and
    // is deterministic for tests.
    const lines = [];
    // Reuse a single cell ref object for per-cell classification checks to avoid allocating
    // `{documentId,sheetId,row,col}` objects in hot loops.
    const cellRef = { documentId, sheetId, row: 0, col: 0 };
    for (let row = normalizedRange.start.row; row <= normalizedRange.end.row; row++) {
      const rowValues = [];
      for (let col = normalizedRange.start.col; col <= normalizedRange.end.col; col++) {
        const coordKey = `${row},${col}`;
        const value = byRowCol.get(coordKey);
        cellRef.row = row;
        cellRef.col = col;
        const classification = effectiveCellClassificationFromIndex(index, cellRef, coordKey);
        const cellRank = classificationRank(classification.level);
        const allowed =
          cellRank === RESTRICTED_CLASSIFICATION_RANK ? restrictedAllowed : maxAllowedRank !== null && cellRank <= maxAllowedRank;

        if (!allowed) {
          // Individual cell is not allowed to be sent to the cloud (either "block" or
          // "redact"). Since the overall request is permitted via redaction, replace the
          // cell with a placeholder.
          rowValues.push(redact(value, null));
          redactedCount++;
          redactions.push({ row, col, classification });
        } else {
          rowValues.push(value === undefined || value === null ? "" : String(value));
        }
      }
      lines.push(rowValues.join("\t"));
    }

    if (redactedCount > 0) {
      lines.push(`\n${redactedCount} cells redacted due to DLP policy.`);
    }

    const context = lines.join("\n");

    this.auditLogger.log({
      type: "ai.request",
      documentId,
      sheetId,
      range,
      action: DLP_ACTION.AI_CLOUD_PROCESSING,
      decision: evaluation,
      selectionClassification,
      redactedCellCount: redactedCount,
      redactions,
    });

    return {
      context,
      selectionClassification,
      evaluation,
      redactions,
    };
  }
}

function cellInNormalizedRange(cell, range) {
  return (
    cell.row >= range.start.row &&
    cell.row <= range.end.row &&
    cell.col >= range.start.col &&
    cell.col <= range.end.col
  );
}

function rangesIntersectNormalized(a, b) {
  return a.start.row <= b.end.row && b.start.row <= a.end.row && a.start.col <= b.end.col && b.start.col <= a.end.col;
}

function buildDlpRangeIndex(ref, records, opts) {
  const selectionRange = ref.range;
  const maxAllowedRank = opts?.maxAllowedRank ?? DEFAULT_CLASSIFICATION_RANK;
  let docClassificationMax = { ...DEFAULT_CLASSIFICATION };
  let sheetClassificationMax = { ...DEFAULT_CLASSIFICATION };
  let baseClassificationMax = { ...DEFAULT_CLASSIFICATION };
  const columnClassificationByIndex = new Map();
  const cellClassificationByCoord = new Map();
  const rangeRecords = [];
  const fallbackRecords = [];

  for (const record of records || []) {
    if (!record || !record.selector || typeof record.selector !== "object") continue;
    const selector = record.selector;
    if (selector.documentId !== ref.documentId) continue;

    // Records at/below the policy `maxAllowed` threshold cannot change per-cell allow/redact
    // decisions and are ignored for performance.
    try {
      const rank = classificationRank(record.classification?.level);
      if (rank <= maxAllowedRank) continue;
    } catch {
      // Ignore invalid classifications so one bad row can't break enforcement.
      continue;
    }

    switch (selector.scope) {
      case "document": {
        docClassificationMax = maxClassification(docClassificationMax, record.classification);
        break;
      }
      case "sheet": {
        if (selector.sheetId === ref.sheetId) {
          sheetClassificationMax = maxClassification(sheetClassificationMax, record.classification);
        }
        break;
      }
      case "column": {
        if (selector.sheetId !== ref.sheetId) break;
        if (typeof selector.columnIndex === "number") {
          if (selector.columnIndex < selectionRange.start.col || selector.columnIndex > selectionRange.end.col) break;
          const existing = columnClassificationByIndex.get(selector.columnIndex);
          columnClassificationByIndex.set(
            selector.columnIndex,
            existing ? maxClassification(existing, record.classification) : record.classification,
          );
        } else {
          // Table/columnId selectors require table metadata to evaluate; AiContextManager's cell refs
          // do not include table context, so these selectors cannot apply and are ignored.
        }
        break;
      }
      case "cell": {
        if (selector.sheetId !== ref.sheetId) break;
        if (typeof selector.row !== "number" || typeof selector.col !== "number") break;
        if (
          selector.row < selectionRange.start.row ||
          selector.row > selectionRange.end.row ||
          selector.col < selectionRange.start.col ||
          selector.col > selectionRange.end.col
        ) {
          break;
        }
        const key = `${selector.row},${selector.col}`;
        const existing = cellClassificationByCoord.get(key);
        cellClassificationByCoord.set(key, existing ? maxClassification(existing, record.classification) : record.classification);
        break;
      }
      case "range": {
        if (selector.sheetId !== ref.sheetId) break;
        if (!selector.range) break;
        const normalized = normalizeRange(selector.range);
        if (!rangesIntersectNormalized(normalized, selectionRange)) break;
        rangeRecords.push({ range: normalized, classification: record.classification });
        break;
      }
      default: {
        // Unknown selector scope: ignore (selectorAppliesToCell would treat it as non-matching).
        break;
      }
    }
  }

  baseClassificationMax = maxClassification(docClassificationMax, sheetClassificationMax);

  return {
    docClassificationMax,
    sheetClassificationMax,
    baseClassificationMax,
    columnClassificationByIndex,
    cellClassificationByCoord,
    rangeRecords,
    fallbackRecords,
  };
}

function effectiveCellClassificationFromIndex(index, cellRef, coordKey) {
  let classification = { ...DEFAULT_CLASSIFICATION };

  classification = maxClassification(classification, index.baseClassificationMax);

  if (classification.level !== CLASSIFICATION_LEVEL.RESTRICTED) {
    const colClassification = index.columnClassificationByIndex.get(cellRef.col);
    if (colClassification) classification = maxClassification(classification, colClassification);
  }

  if (classification.level !== CLASSIFICATION_LEVEL.RESTRICTED) {
    const key = coordKey ?? `${cellRef.row},${cellRef.col}`;
    const cellClassification = index.cellClassificationByCoord.get(key);
    if (cellClassification) classification = maxClassification(classification, cellClassification);
  }

  if (classification.level !== CLASSIFICATION_LEVEL.RESTRICTED) {
    for (const record of index.rangeRecords) {
      if (!cellInNormalizedRange(cellRef, record.range)) continue;
      classification = maxClassification(classification, record.classification);
      if (classification.level === CLASSIFICATION_LEVEL.RESTRICTED) break;
    }
  }

  if (classification.level !== CLASSIFICATION_LEVEL.RESTRICTED && index.fallbackRecords.length > 0) {
    const fallback = effectiveCellClassification(cellRef, index.fallbackRecords);
    classification = maxClassification(classification, fallback);
  }

  return classification;
}
