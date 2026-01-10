import test from "node:test";
import assert from "node:assert/strict";

import { DLP_ACTION } from "../../packages/security/dlp/src/actions.js";
import { CLASSIFICATION_LEVEL } from "../../packages/security/dlp/src/classification.js";
import { createDefaultOrgPolicy, mergePolicies } from "../../packages/security/dlp/src/policy.js";
import { evaluatePolicy, DLP_DECISION } from "../../packages/security/dlp/src/policyEngine.js";

test("policy evaluation: clipboard copy allowed at or below maxAllowed", () => {
  const policy = createDefaultOrgPolicy();
  const decision = evaluatePolicy({
    action: DLP_ACTION.CLIPBOARD_COPY,
    classification: { level: CLASSIFICATION_LEVEL.CONFIDENTIAL },
    policy,
  });

  assert.equal(decision.decision, DLP_DECISION.ALLOW);
});

test("policy evaluation: clipboard copy blocked above maxAllowed", () => {
  const policy = createDefaultOrgPolicy();
  const decision = evaluatePolicy({
    action: DLP_ACTION.CLIPBOARD_COPY,
    classification: { level: CLASSIFICATION_LEVEL.RESTRICTED },
    policy,
  });

  assert.equal(decision.decision, DLP_DECISION.BLOCK);
});

test("policy merge: document overrides cannot weaken org policy", () => {
  const orgPolicy = createDefaultOrgPolicy();
  orgPolicy.rules[DLP_ACTION.CLIPBOARD_COPY] = { maxAllowed: CLASSIFICATION_LEVEL.INTERNAL };

  const documentPolicy = createDefaultOrgPolicy();
  // Attempt to allow more sensitive data than org permits.
  documentPolicy.rules[DLP_ACTION.CLIPBOARD_COPY] = { maxAllowed: CLASSIFICATION_LEVEL.RESTRICTED };

  const { policy: effective } = mergePolicies({ orgPolicy, documentPolicy });
  assert.equal(effective.rules[DLP_ACTION.CLIPBOARD_COPY].maxAllowed, CLASSIFICATION_LEVEL.INTERNAL);
});

test("policy merge: document overrides can tighten org policy", () => {
  const orgPolicy = createDefaultOrgPolicy();
  orgPolicy.rules[DLP_ACTION.CLIPBOARD_COPY] = { maxAllowed: CLASSIFICATION_LEVEL.CONFIDENTIAL };

  const documentPolicy = createDefaultOrgPolicy();
  documentPolicy.rules[DLP_ACTION.CLIPBOARD_COPY] = { maxAllowed: CLASSIFICATION_LEVEL.INTERNAL };

  const { policy: effective } = mergePolicies({ orgPolicy, documentPolicy });
  assert.equal(effective.rules[DLP_ACTION.CLIPBOARD_COPY].maxAllowed, CLASSIFICATION_LEVEL.INTERNAL);
});

test("AI policy: redact disallowed content instead of blocking by default", () => {
  const policy = createDefaultOrgPolicy();
  const decision = evaluatePolicy({
    action: DLP_ACTION.AI_CLOUD_PROCESSING,
    classification: { level: CLASSIFICATION_LEVEL.RESTRICTED },
    policy,
  });
  assert.equal(decision.decision, DLP_DECISION.REDACT);
});

test("AI policy: explicitly including restricted content is blocked unless allowed", () => {
  const policy = createDefaultOrgPolicy();
  const decision = evaluatePolicy({
    action: DLP_ACTION.AI_CLOUD_PROCESSING,
    classification: { level: CLASSIFICATION_LEVEL.RESTRICTED },
    policy,
    options: { includeRestrictedContent: true },
  });
  assert.equal(decision.decision, DLP_DECISION.BLOCK);
});

