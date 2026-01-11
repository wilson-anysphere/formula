import { DLP_ACTION } from "./actions.js";
import { evaluatePolicy, DLP_DECISION } from "./policyEngine.js";
import { effectiveCellClassification, effectiveRangeClassification, normalizeRange } from "./selectors.js";
import { DlpViolationError } from "./errors.js";
import dlpCore from "./core.js";

const { redact } = dlpCore;

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

    const redactions = [];
    let redactedCount = 0;

    // Render as a small TSV-like snippet. This matches many LLM prompt conventions and
    // is deterministic for tests.
    const lines = [];
    for (let row = normalizedRange.start.row; row <= normalizedRange.end.row; row++) {
      const rowValues = [];
      for (let col = normalizedRange.start.col; col <= normalizedRange.end.col; col++) {
        const value = byRowCol.get(`${row},${col}`);
        const classification = effectiveCellClassification({ documentId, sheetId, row, col }, records);
        const allowedDecision = evaluatePolicy({
          action: DLP_ACTION.AI_CLOUD_PROCESSING,
          classification,
          policy,
          options: { includeRestrictedContent },
        });

        if (allowedDecision.decision !== DLP_DECISION.ALLOW) {
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
