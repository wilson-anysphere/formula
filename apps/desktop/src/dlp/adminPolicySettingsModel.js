import { createDefaultOrgPolicy, validatePolicy } from "../../../../packages/security/dlp/src/policy.js";

/**
 * Minimal policy model used by an "Org Settings" UI screen.
 *
 * The real app would have a richer form schema; this module provides the core defaults
 * and validation that the UI can call.
 */
export function getDefaultPolicyForUi() {
  return createDefaultOrgPolicy();
}

export function validatePolicyForUi(policy) {
  validatePolicy(policy);
  return true;
}
