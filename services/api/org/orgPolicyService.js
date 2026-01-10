function isPlainObject(value) {
  return value !== null && typeof value === "object" && value.constructor === Object;
}

function deepMerge(base, patch) {
  if (!isPlainObject(base) || !isPlainObject(patch)) return patch;
  const out = { ...base };
  for (const [key, value] of Object.entries(patch)) {
    if (isPlainObject(value) && isPlainObject(base[key])) {
      out[key] = deepMerge(base[key], value);
    } else {
      out[key] = value;
    }
  }
  return out;
}

function normalizeDataResidency(dataResidency) {
  if (!dataResidency || typeof dataResidency !== "object") {
    throw new TypeError("dataResidency must be an object");
  }

  const region = dataResidency.region;
  if (region === "us" || region === "eu" || region === "apac") {
    return {
      ...dataResidency,
      region,
      allowedRegions: [region],
      primaryStorageRegion: region,
      backupStorageRegion: region,
      aiProcessingRegion: region,
      allowCrossRegionProcessing: false
    };
  }

  if (region === "custom") {
    const allowedRegions = Array.isArray(dataResidency.allowedRegions)
      ? dataResidency.allowedRegions
      : [];
    const fallback = allowedRegions[0] ?? dataResidency.primaryStorageRegion ?? "us";
    return {
      ...dataResidency,
      allowedRegions,
      primaryStorageRegion: dataResidency.primaryStorageRegion ?? fallback,
      backupStorageRegion: dataResidency.backupStorageRegion ?? fallback,
      aiProcessingRegion: dataResidency.aiProcessingRegion ?? fallback
    };
  }

  throw new Error(`Unsupported residency region: ${region}`);
}

function normalizeOrgPolicies(policies) {
  return {
    ...policies,
    dataResidency: normalizeDataResidency(policies.dataResidency)
  };
}

function defaultOrgPolicies() {
  return normalizeOrgPolicies({
    encryption: {
      cloudEncryptionAtRest: true,
      kms: { provider: "local" },
      keyRotationDays: 90,
      transit: {
        minTlsVersion: "TLSv1.3",
        certificatePinning: { enabled: false, pins: [] }
      }
    },
    dataResidency: {
      region: "us",
      allowedRegions: ["us"],
      primaryStorageRegion: "us",
      backupStorageRegion: "us",
      aiProcessingRegion: "us",
      allowCrossRegionProcessing: false
    },
    retention: {
      versionRetentionDays: 90,
      auditLogRetentionDays: 365,
      deletedDocumentRetentionDays: 30,
      legalHoldOverridesRetention: true
    }
  });
}

class InMemoryOrgPolicyStore {
  constructor() {
    this._byOrgId = new Map();
  }

  async get(orgId) {
    const existing = this._byOrgId.get(orgId);
    return existing ? structuredClone(existing) : null;
  }

  async set(orgId, value) {
    this._byOrgId.set(orgId, structuredClone(value));
  }
}

class OrgPolicyService {
  constructor({ store, auditLogger }) {
    this.store = store;
    this.auditLogger = auditLogger;
  }

  async get(orgId) {
    const existing = await this.store.get(orgId);
    if (existing) return existing;
    const defaults = defaultOrgPolicies();
    await this.store.set(orgId, defaults);
    return defaults;
  }

  async update(orgId, patch, { actor } = {}) {
    const before = await this.get(orgId);
    const after = normalizeOrgPolicies(deepMerge(before, patch));
    await this.store.set(orgId, after);

    if (this.auditLogger) {
      const changed = [];
      if (JSON.stringify(before.encryption) !== JSON.stringify(after.encryption)) {
        changed.push("encryption");
      }
      if (JSON.stringify(before.dataResidency) !== JSON.stringify(after.dataResidency)) {
        changed.push("dataResidency");
      }
      if (JSON.stringify(before.retention) !== JSON.stringify(after.retention)) {
        changed.push("retention");
      }

      for (const section of changed) {
        await this.auditLogger.log({
          eventType: `org.policy.${section}.updated`,
          orgId,
          actor: actor ?? null,
          details: {
            before: before[section],
            after: after[section]
          }
        });
      }
    }

    return after;
  }
}

module.exports = { defaultOrgPolicies, InMemoryOrgPolicyStore, OrgPolicyService };
