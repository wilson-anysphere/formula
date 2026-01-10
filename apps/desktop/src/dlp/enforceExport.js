import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { evaluatePolicy, DLP_DECISION } from "../../../../packages/security/dlp/src/policyEngine.js";
import {
  effectiveDocumentClassification,
  effectiveRangeClassification,
} from "../../../../packages/security/dlp/src/selectors.js";
import { DlpViolationError } from "../../../../packages/security/dlp/src/errors.js";

const EXPORT_ACTION_BY_FORMAT = Object.freeze({
  csv: DLP_ACTION.EXPORT_CSV,
  pdf: DLP_ACTION.EXPORT_PDF,
  xlsx: DLP_ACTION.EXPORT_XLSX,
});

/**
 * Enforce DLP policy for an export operation.
 *
 * If `range` is omitted, the export is treated as a whole-document operation.
 *
 * @param {{
 *  documentId: string,
 *  sheetId?: string,
 *  range?: {start:{row:number,col:number}, end:{row:number,col:number}},
 *  format: "csv"|"pdf"|"xlsx",
 *  classificationStore: {list(documentId:string): Array<{selector:any, classification:any}>},
 *  policy: any
 * }} params
 */
export function enforceExport({ documentId, sheetId, range, format, classificationStore, policy }) {
  const action = EXPORT_ACTION_BY_FORMAT[format];
  if (!action) throw new Error(`Unknown export format: ${format}`);

  const records = classificationStore.list(documentId);
  const selectionClassification =
    range && sheetId
      ? effectiveRangeClassification({ documentId, sheetId, range }, records)
      : effectiveDocumentClassification(documentId, records);

  const decision = evaluatePolicy({ action, classification: selectionClassification, policy });
  if (decision.decision === DLP_DECISION.BLOCK) throw new DlpViolationError(decision);
  return { decision, selectionClassification };
}

