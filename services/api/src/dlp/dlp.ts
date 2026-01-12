import {
  CLASSIFICATION_LEVEL,
  CLASSIFICATION_LEVELS,
  classificationRank,
  DLP_ACTION,
  DLP_DECISION,
  DLP_POLICY_VERSION,
  DLP_REASON_CODE,
  evaluatePolicy,
  isAllowed,
  maxClassification,
  normalizeClassification as coreNormalizeClassification,
  normalizeDlpPolicy,
  normalizeSelector,
  redact,
  resolveClassification,
  selectorKey,
  validateDlpPolicy,
} from "../../../../shared/dlp-core";

import type {
  Classification,
  ClassificationLevel,
  DlpAiRule,
  DlpPolicy,
  DlpRuleBase,
  PolicyEvaluationResult,
} from "../../../../shared/dlp-core";

export {
  CLASSIFICATION_LEVEL,
  CLASSIFICATION_LEVELS,
  classificationRank,
  DLP_ACTION,
  DLP_DECISION,
  DLP_POLICY_VERSION,
  DLP_REASON_CODE,
  evaluatePolicy,
  isAllowed,
  maxClassification,
  normalizeDlpPolicy,
  normalizeSelector,
  redact,
  resolveClassification,
  selectorKey,
  validateDlpPolicy
};
export type { Classification, ClassificationLevel, DlpAiRule, DlpPolicy, DlpRuleBase, PolicyEvaluationResult };

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export function normalizeClassification(classification: unknown): Classification {
  // API request validation should treat missing/invalid classifications as a hard error;
  // callers that want a permissive default can use the shared core directly.
  if (!isObject(classification)) throw new Error("Classification must be an object");
  return coreNormalizeClassification(classification);
}
