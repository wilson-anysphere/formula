import { DLP_ACTION } from "../../../../packages/security/dlp/src/actions.js";
import { evaluatePolicy, DLP_DECISION } from "../../../../packages/security/dlp/src/policyEngine.js";
import { effectiveDocumentClassification } from "../../../../packages/security/dlp/src/selectors.js";
import { DlpViolationError } from "../../../../packages/security/dlp/src/errors.js";

/**
 * Enforce DLP policy for creating an external sharing link.
 *
 * This is evaluated at the whole-document level: if any sensitive classification is
 * present anywhere in the document, external link sharing is restricted accordingly.
 */
export function enforceExternalShareLink({ documentId, classificationStore, policy }) {
  const records = classificationStore.list(documentId);
  const documentClassification = effectiveDocumentClassification(documentId, records);
  const decision = evaluatePolicy({
    action: DLP_ACTION.SHARE_EXTERNAL_LINK,
    classification: documentClassification,
    policy,
  });
  if (decision.decision === DLP_DECISION.BLOCK) throw new DlpViolationError(decision);
  return { decision, documentClassification };
}

