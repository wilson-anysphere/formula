import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { evaluatePolicy, DLP_DECISION } from "../../../../packages/security/dlp/src/policyEngine.js";
import { effectiveRangeClassification } from "../../../../packages/security/dlp/src/selectors.js";
import { DlpViolationError } from "../../../../packages/security/dlp/src/errors.js";

/**
 * Enforce clipboard copy DLP policy for a selection.
 *
 * @param {{
 *  documentId: string,
 *  sheetId: string,
 *  range: {start:{row:number,col:number}, end:{row:number,col:number}},
 *  classificationStore: {list(documentId:string): Array<{selector:any, classification:any}>},
 *  policy: any
 * }} params
 */
export function enforceClipboardCopy({ documentId, sheetId, range, classificationStore, policy }) {
  const records = classificationStore.list(documentId);
  const selectionClassification = effectiveRangeClassification({ documentId, sheetId, range }, records);
  const decision = evaluatePolicy({
    action: DLP_ACTION.CLIPBOARD_COPY,
    classification: selectionClassification,
    policy,
  });

  if (decision.decision === DLP_DECISION.BLOCK) {
    throw new DlpViolationError(decision);
  }

  return { decision, selectionClassification };
}
