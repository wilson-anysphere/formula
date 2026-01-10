import { CLASSIFICATION_LEVELS, classificationRank } from "./classification.js";
import { DLP_ACTION } from "./actions.js";

export const POLICY_SOURCE = Object.freeze({
  ORG: "org",
  DOCUMENT: "document",
  EFFECTIVE: "effective",
});

/**
 * @typedef {{maxAllowed: string|null}} ThresholdRule
 */

/**
 * @typedef {{
 *  maxAllowed: string|null,
 *  allowRestrictedContent?: boolean,
 *  redactDisallowed?: boolean
 * }} AiRule
 */

/**
 * @typedef {{
 *  version: number,
 *  allowDocumentOverrides: boolean,
 *  rules: Record<string, ThresholdRule | AiRule>
 * }} DlpPolicy
 */

export function createDefaultOrgPolicy() {
  return {
    version: 1,
    allowDocumentOverrides: true,
    rules: {
      [DLP_ACTION.SHARE_EXTERNAL_LINK]: { maxAllowed: "Internal" },
      [DLP_ACTION.EXPORT_CSV]: { maxAllowed: "Confidential" },
      [DLP_ACTION.EXPORT_PDF]: { maxAllowed: "Confidential" },
      [DLP_ACTION.EXPORT_XLSX]: { maxAllowed: "Confidential" },
      [DLP_ACTION.CLIPBOARD_COPY]: { maxAllowed: "Confidential" },
      [DLP_ACTION.EXTERNAL_CONNECTOR]: { maxAllowed: "Internal" },
      [DLP_ACTION.AI_CLOUD_PROCESSING]: {
        maxAllowed: "Confidential",
        // Explicitly required to ever send Restricted content to a cloud model.
        allowRestrictedContent: false,
        // Redact cells above `maxAllowed` rather than blocking the entire request.
        redactDisallowed: true,
      },
    },
  };
}

function validateLevel(level) {
  if (level === null) return;
  if (!CLASSIFICATION_LEVELS.includes(level)) {
    throw new Error(`Invalid classification level in policy: ${level}`);
  }
}

export function validatePolicy(policy) {
  if (!policy || typeof policy !== "object") throw new Error("Policy must be an object");
  if (!Number.isInteger(policy.version)) throw new Error("Policy.version must be an integer");
  if (typeof policy.allowDocumentOverrides !== "boolean") {
    throw new Error("Policy.allowDocumentOverrides must be a boolean");
  }
  if (!policy.rules || typeof policy.rules !== "object") throw new Error("Policy.rules must be an object");
  for (const [action, rule] of Object.entries(policy.rules)) {
    if (!rule || typeof rule !== "object") throw new Error(`Rule for ${action} must be an object`);
    validateLevel(rule.maxAllowed);
    if (action === DLP_ACTION.AI_CLOUD_PROCESSING) {
      if (typeof rule.allowRestrictedContent !== "boolean") {
        throw new Error("AI rule allowRestrictedContent must be boolean");
      }
      if (typeof rule.redactDisallowed !== "boolean") {
        throw new Error("AI rule redactDisallowed must be boolean");
      }
    }
  }
}

function minLevel(a, b) {
  if (a === null || b === null) return null;
  // Smaller rank = less restrictive. We want the *most restrictive* effective maxAllowed,
  // i.e., the minimum of the two maxAllowed thresholds.
  return classificationRank(a) <= classificationRank(b) ? a : b;
}

/**
 * Merge org policy with a per-document override.
 *
 * The override is allowed to be more restrictive, but not more permissive. This prevents
 * document-level settings from weakening organization controls.
 *
 * @param {{orgPolicy: any, documentPolicy?: any}} params
 */
export function mergePolicies({ orgPolicy, documentPolicy }) {
  validatePolicy(orgPolicy);
  if (!documentPolicy) return { policy: orgPolicy, source: POLICY_SOURCE.ORG };
  validatePolicy(documentPolicy);
  if (!orgPolicy.allowDocumentOverrides) return { policy: orgPolicy, source: POLICY_SOURCE.ORG };

  const merged = {
    version: Math.max(orgPolicy.version, documentPolicy.version),
    allowDocumentOverrides: orgPolicy.allowDocumentOverrides,
    rules: { ...orgPolicy.rules },
  };

  for (const [action, overrideRule] of Object.entries(documentPolicy.rules || {})) {
    const baseRule = merged.rules[action];
    if (!baseRule) continue;
    merged.rules[action] = { ...baseRule, ...overrideRule };
    merged.rules[action].maxAllowed = minLevel(baseRule.maxAllowed ?? null, overrideRule.maxAllowed ?? null);

    // For AI, do not allow a doc override to enable sending Restricted data if org disallows it.
    if (action === DLP_ACTION.AI_CLOUD_PROCESSING) {
      const overrideAllowRestricted =
        Object.prototype.hasOwnProperty.call(overrideRule, "allowRestrictedContent") && typeof overrideRule.allowRestrictedContent === "boolean"
          ? overrideRule.allowRestrictedContent
          : undefined;
      const overrideRedactDisallowed =
        Object.prototype.hasOwnProperty.call(overrideRule, "redactDisallowed") && typeof overrideRule.redactDisallowed === "boolean"
          ? overrideRule.redactDisallowed
          : undefined;

      // Document overrides can only tighten org policy:
      // - They may disable allowRestrictedContent / redactDisallowed.
      // - They may not enable them if the org policy has them disabled.
      merged.rules[action].allowRestrictedContent =
        Boolean(baseRule.allowRestrictedContent) && overrideAllowRestricted !== false;
      merged.rules[action].redactDisallowed = Boolean(baseRule.redactDisallowed) && overrideRedactDisallowed !== false;
    }
  }

  validatePolicy(merged);
  return { policy: merged, source: POLICY_SOURCE.EFFECTIVE };
}
