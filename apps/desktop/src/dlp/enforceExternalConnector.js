import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { evaluatePolicy, DLP_DECISION } from "../../../../packages/security/dlp/src/policyEngine.js";
import {
  effectiveDocumentClassification,
  effectiveRangeClassification,
} from "../../../../packages/security/dlp/src/selectors.js";
import { DlpViolationError } from "../../../../packages/security/dlp/src/errors.js";

/**
 * Enforce DLP policy for sending data to an external connector (e.g., Slack, webhook,
 * cloud storage connector).
 *
 * If `range` is omitted, the operation is treated as a whole-document connector sync.
 */
export function enforceExternalConnector({ documentId, sheetId, range, classificationStore, policy }) {
  const records = classificationStore.list(documentId);
  const selectionClassification =
    range && sheetId
      ? effectiveRangeClassification({ documentId, sheetId, range }, records)
      : effectiveDocumentClassification(documentId, records);

  const decision = evaluatePolicy({
    action: DLP_ACTION.EXTERNAL_CONNECTOR,
    classification: selectionClassification,
    policy,
  });
  if (decision.decision === DLP_DECISION.BLOCK) throw new DlpViolationError(decision);
  return { decision, selectionClassification };
}

