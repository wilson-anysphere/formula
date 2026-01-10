import { compareClassification, normalizeClassification, classificationRank } from "./classification.js";
import { DLP_ACTION } from "./actions.js";

export const DLP_DECISION = Object.freeze({
  ALLOW: "allow",
  BLOCK: "block",
  REDACT: "redact",
});

export const DLP_REASON_CODE = Object.freeze({
  BLOCKED_BY_POLICY: "dlp.blockedByPolicy",
  INVALID_POLICY: "dlp.invalidPolicy",
});

/**
 * @param {any} policy
 * @param {string} action
 */
function ruleForAction(policy, action) {
  if (!policy || typeof policy !== "object" || !policy.rules) {
    throw new Error("Invalid policy object");
  }
  return policy.rules[action] || { maxAllowed: null };
}

/**
 * Evaluate whether an operation is allowed under the provided policy for data with the
 * given classification.
 *
 * @param {{
 *  action: string,
 *  classification: {level:string, labels?:string[]},
 *  policy: any,
 *  options?: {
 *    includeRestrictedContent?: boolean
 *  }
 * }} params
 */
export function evaluatePolicy({ action, classification, policy, options = {} }) {
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

  // AI cloud processing supports a "redact instead of block" mode, plus an explicit
  // exception for Restricted content when a user/admin has opted in.
  if (action === DLP_ACTION.AI_CLOUD_PROCESSING) {
    if (normalized.level === "Restricted" && options.includeRestrictedContent) {
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

/**
 * Returns true if the provided classification is allowed under a maxAllowed threshold.
 *
 * @param {{level:string}} classification
 * @param {string|null} maxAllowed
 */
export function isClassificationAllowed(classification, maxAllowed) {
  if (maxAllowed === null) return false;
  const level = normalizeClassification(classification).level;
  return classificationRank(level) <= classificationRank(maxAllowed);
}
