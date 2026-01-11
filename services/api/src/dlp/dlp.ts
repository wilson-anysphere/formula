import {
  CLASSIFICATION_LEVELS,
  DLP_DECISION,
  DLP_REASON_CODE,
  evaluatePolicy,
  isAllowed,
  normalizeClassification as coreNormalizeClassification,
  redact,
  selectorKey,
  validateDlpPolicy,
} from "../../../../shared/dlp-core";

import type { Classification, ClassificationLevel, DlpAiRule, DlpPolicy, DlpRuleBase } from "../../../../shared/dlp-core";

export { CLASSIFICATION_LEVELS, DLP_DECISION, DLP_REASON_CODE, evaluatePolicy, isAllowed, redact, selectorKey, validateDlpPolicy };
export type { Classification, ClassificationLevel, DlpAiRule, DlpPolicy, DlpRuleBase };

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export function normalizeClassification(classification: unknown): Classification {
  // API request validation should treat missing/invalid classifications as a hard error;
  // callers that want a permissive default can use the shared core directly.
  if (!isObject(classification)) throw new Error("Classification must be an object");
  return coreNormalizeClassification(classification);
}
