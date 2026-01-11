import {
  CLASSIFICATION_LEVEL,
  classificationRank,
  DLP_ACTION,
  DLP_DECISION,
  DLP_REASON_CODE,
  evaluatePolicy,
  maxClassification,
  normalizeClassification,
  validateDlpPolicy,
  type Classification,
  type DlpPolicy,
  type PolicyEvaluationResult,
} from "./dlp";

import {
  getAggregateClassificationForRange,
  getEffectiveClassificationForSelector,
  normalizeSelectorColumns,
  type DbClient,
} from "./classificationResolver";

const DEFAULT_CLASSIFICATION: Classification = {
  level: CLASSIFICATION_LEVEL.PUBLIC,
  labels: [],
};

type EffectivePolicyResult =
  | { type: "unconfigured" }
  | { type: "invalid" }
  | { type: "configured"; policy: DlpPolicy };

function minMaxAllowed(a: any, b: any): any {
  if (a == null || b == null) return null;
  return classificationRank(a) <= classificationRank(b) ? a : b;
}

function mergePolicies(params: { orgPolicy: DlpPolicy; documentPolicy: DlpPolicy }): DlpPolicy {
  const { orgPolicy, documentPolicy } = params;

  const merged: DlpPolicy = {
    version: Math.max(orgPolicy.version, documentPolicy.version),
    allowDocumentOverrides: orgPolicy.allowDocumentOverrides,
    rules: { ...orgPolicy.rules },
  };

  for (const [action, overrideRuleRaw] of Object.entries(documentPolicy.rules || {})) {
    const baseRuleRaw = merged.rules[action];
    if (!baseRuleRaw) continue;

    const baseRule = baseRuleRaw as any;
    const overrideRule = overrideRuleRaw as any;

    const nextRule: any = { ...baseRule, ...overrideRule };
    nextRule.maxAllowed = minMaxAllowed(baseRule.maxAllowed ?? null, overrideRule.maxAllowed ?? null);

    if (action === DLP_ACTION.AI_CLOUD_PROCESSING) {
      const overrideAllowRestricted =
        Object.prototype.hasOwnProperty.call(overrideRule, "allowRestrictedContent") &&
        typeof overrideRule.allowRestrictedContent === "boolean"
          ? overrideRule.allowRestrictedContent
          : undefined;

      const overrideRedactDisallowed =
        Object.prototype.hasOwnProperty.call(overrideRule, "redactDisallowed") &&
        typeof overrideRule.redactDisallowed === "boolean"
          ? overrideRule.redactDisallowed
          : undefined;

      // Document overrides can only tighten org policy:
      // - They may disable allowRestrictedContent / redactDisallowed.
      // - They may not enable them if the org policy has them disabled.
      nextRule.allowRestrictedContent = Boolean(baseRule.allowRestrictedContent) && overrideAllowRestricted !== false;
      nextRule.redactDisallowed = Boolean(baseRule.redactDisallowed) && overrideRedactDisallowed !== false;
    }

    (merged.rules as any)[action] = nextRule;
  }

  validateDlpPolicy(merged);
  return merged;
}

export async function getClassificationForSelectorKey(
  db: DbClient,
  docId: string,
  selectorKey: string
): Promise<Classification> {
  const res = await db.query(
    `
      SELECT classification
      FROM document_classifications
      WHERE document_id = $1 AND selector_key = $2
      LIMIT 1
    `,
    [docId, selectorKey]
  );

  if (res.rowCount !== 1) return DEFAULT_CLASSIFICATION;
  return normalizeClassification(res.rows[0]!.classification);
}

async function getEffectiveClassificationForSelectorOrRange(db: DbClient, docId: string, selector: unknown): Promise<Classification> {
  const normalized = normalizeSelectorColumns(selector);
  if (normalized.scope === "range") {
    return getAggregateClassificationForRange(
      db,
      docId,
      normalized.sheetId!,
      normalized.startRow!,
      normalized.startCol!,
      normalized.endRow!,
      normalized.endCol!
    );
  }
  if (normalized.scope === "cell") {
    // DLP enforcement should be conservative: a cell classification should not be able to
    // weaken a broader range/column/sheet/document restriction. Reuse the aggregate range
    // classifier for a single cell to ensure we union labels and take the max level across
    // all overlapping selectors.
    return getAggregateClassificationForRange(
      db,
      docId,
      normalized.sheetId!,
      normalized.row!,
      normalized.col!,
      normalized.row!,
      normalized.col!
    );
  }

  const resolved = await getEffectiveClassificationForSelector(db, docId, selector);
  return resolved.classification;
}

export async function getEffectiveDocumentClassification(db: DbClient, docId: string): Promise<Classification> {
  const res = await db.query(
    `
      SELECT classification
      FROM document_classifications
      WHERE document_id = $1
    `,
    [docId]
  );

  if (res.rowCount === 0) return DEFAULT_CLASSIFICATION;

  let effective: Classification = DEFAULT_CLASSIFICATION;
  for (const row of res.rows as Array<{ classification: unknown }>) {
    effective = maxClassification(effective, row.classification);
  }
  return normalizeClassification(effective);
}

async function resolveEffectivePolicy(db: DbClient, orgId: string, docId: string): Promise<EffectivePolicyResult> {
  const orgPolicyRes = await db.query("SELECT policy FROM org_dlp_policies WHERE org_id = $1", [orgId]);
  if (orgPolicyRes.rowCount !== 1) {
    // We intentionally treat missing org policy as "no DLP configured" (allow-all)
    // so orgs can adopt the feature incrementally.
    return { type: "unconfigured" };
  }

  const orgPolicyRaw = orgPolicyRes.rows[0]!.policy as unknown;
  try {
    validateDlpPolicy(orgPolicyRaw);
  } catch {
    return { type: "invalid" };
  }

  const orgPolicy = orgPolicyRaw as DlpPolicy;
  if (!orgPolicy.allowDocumentOverrides) {
    return { type: "configured", policy: orgPolicy };
  }

  const docPolicyRes = await db.query("SELECT policy FROM document_dlp_policies WHERE document_id = $1", [docId]);
  if (docPolicyRes.rowCount !== 1) return { type: "configured", policy: orgPolicy };

  const docPolicyRaw = docPolicyRes.rows[0]!.policy as unknown;
  try {
    validateDlpPolicy(docPolicyRaw);
  } catch {
    return { type: "invalid" };
  }

  return { type: "configured", policy: mergePolicies({ orgPolicy, documentPolicy: docPolicyRaw as DlpPolicy }) };
}

export async function evaluateDocumentDlpPolicy(
  db: DbClient,
  params: {
    orgId: string;
    docId: string;
    action: string;
    options?: { includeRestrictedContent?: boolean };
    selector?: unknown;
  }
): Promise<PolicyEvaluationResult> {
  const classification =
    params.selector !== undefined
      ? await getEffectiveClassificationForSelectorOrRange(db, params.docId, params.selector)
      : await getEffectiveDocumentClassification(db, params.docId);

  const policyRes = await resolveEffectivePolicy(db, params.orgId, params.docId);

  if (policyRes.type === "invalid") {
    return {
      action: params.action,
      decision: DLP_DECISION.BLOCK,
      reasonCode: DLP_REASON_CODE.INVALID_POLICY,
      classification,
      maxAllowed: null,
    };
  }

  if (policyRes.type === "unconfigured") {
    return {
      action: params.action,
      decision: DLP_DECISION.ALLOW,
      classification,
      maxAllowed: CLASSIFICATION_LEVEL.RESTRICTED,
    };
  }

  return evaluatePolicy({
    action: params.action,
    classification,
    policy: policyRes.policy,
    options: params.options,
  });
}
