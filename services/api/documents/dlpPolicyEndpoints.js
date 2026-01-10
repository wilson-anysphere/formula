import { validatePolicy } from "../../../packages/security/dlp/src/policy.js";

/**
 * Stores per-document DLP policy overrides. These are merged with the organization
 * policy at evaluation time (see `mergePolicies` in the security package).
 */
export class InMemoryDocumentPolicyStore {
  constructor() {
    this.byDocumentId = new Map();
  }

  get(documentId) {
    return this.byDocumentId.get(documentId) || null;
  }

  set(documentId, policy) {
    this.byDocumentId.set(documentId, policy);
  }
}

export function createDocumentDlpPolicyEndpoints({ store }) {
  if (!store) throw new Error("createDocumentDlpPolicyEndpoints requires store");

  return {
    /**
     * GET /documents/:documentId/dlp-policy
     */
    async get({ documentId }) {
      const policy = store.get(documentId);
      if (!policy) return { status: 404, body: { error: "not_found" } };
      return { status: 200, body: policy };
    },

    /**
     * PUT /documents/:documentId/dlp-policy
     */
    async put({ documentId, policy }) {
      try {
        validatePolicy(policy);
      } catch (error) {
        return { status: 400, body: { error: "invalid_policy", message: error.message } };
      }
      store.set(documentId, policy);
      return { status: 200, body: policy };
    },
  };
}

