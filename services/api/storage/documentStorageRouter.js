const { resolveAiProcessingRegion, resolvePrimaryStorageRegion } = require("../policies/dataResidency.js");

async function loadEnvelopeCrypto() {
  return import("../../../packages/security/crypto/envelope.js");
}

class DocumentStorageRouter {
  constructor({ orgPolicyService, regionalObjectStore, kmsProvider }) {
    this.orgPolicyService = orgPolicyService;
    this.regionalObjectStore = regionalObjectStore;
    this.kmsProvider = kmsProvider;
  }

  async resolveStorageRegion(orgId) {
    const policy = await this.orgPolicyService.get(orgId);
    return resolvePrimaryStorageRegion(policy.dataResidency);
  }

  async resolveAiRegion(orgId) {
    const policy = await this.orgPolicyService.get(orgId);
    return resolveAiProcessingRegion(policy.dataResidency);
  }

  async putDocument({ orgId, docId, plaintext }) {
    if (!Buffer.isBuffer(plaintext)) {
      throw new TypeError("plaintext must be a Buffer");
    }

    const region = await this.resolveStorageRegion(orgId);

    const policy = await this.orgPolicyService.get(orgId);
    if (!policy.encryption.cloudEncryptionAtRest) {
      throw new Error("Cloud encryption-at-rest cannot be disabled for enterprise orgs");
    }

    const { encryptEnvelope } = await loadEnvelopeCrypto();
    const encryptionContext = { orgId, docId, purpose: "document" };
    const encrypted = await encryptEnvelope({
      plaintext,
      kmsProvider: this.kmsProvider,
      encryptionContext
    });

    const key = `documents/${orgId}/${docId}`;
    await this.regionalObjectStore.putObject(region, key, Buffer.from(JSON.stringify(encrypted), "utf8"));
    return { region, key };
  }

  async getDocument({ orgId, docId }) {
    const region = await this.resolveStorageRegion(orgId);
    const key = `documents/${orgId}/${docId}`;
    const stored = await this.regionalObjectStore.getObject(region, key);
    if (!stored) return null;

    const encrypted = JSON.parse(stored.toString("utf8"));
    const encryptionContext = { orgId, docId, purpose: "document" };

    const { decryptEnvelope } = await loadEnvelopeCrypto();
    return await decryptEnvelope({
      encryptedEnvelope: encrypted,
      kmsProvider: this.kmsProvider,
      encryptionContext
    });
  }
}

module.exports = { DocumentStorageRouter };
