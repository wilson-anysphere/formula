import { validatePolicy } from "../../../packages/security/dlp/src/policy.js";

/**
 * In-memory store used for local development and unit tests. Real deployments would
 * persist to a database (e.g., Postgres) keyed by organization id.
 */
export class InMemoryOrgPolicyStore {
  constructor() {
    this.byOrgId = new Map();
  }

  get(orgId) {
    return this.byOrgId.get(orgId) || null;
  }

  set(orgId, policy) {
    this.byOrgId.set(orgId, policy);
  }
}

/**
 * "Endpoints" are expressed as plain async functions to keep this repository
 * framework-agnostic. They can be adapted to Express/Fastify/etc by a thin wrapper.
 */
export function createOrgDlpPolicyEndpoints({ store }) {
  if (!store) throw new Error("createOrgDlpPolicyEndpoints requires store");

  return {
    /**
     * GET /orgs/:orgId/dlp-policy
     */
    async getOrgPolicy({ orgId }) {
      const policy = store.get(orgId);
      if (!policy) return { status: 404, body: { error: "not_found" } };
      return { status: 200, body: policy };
    },

    /**
     * PUT /orgs/:orgId/dlp-policy
     */
    async putOrgPolicy({ orgId, policy }) {
      try {
        validatePolicy(policy);
      } catch (error) {
        return { status: 400, body: { error: "invalid_policy", message: error.message } };
      }
      store.set(orgId, policy);
      return { status: 200, body: policy };
    },
  };
}

