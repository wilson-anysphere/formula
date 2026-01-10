import { validatePolicy } from "./policy.js";

/**
 * @typedef {{getItem(key:string): (string|null), setItem(key:string,value:string): void, removeItem(key:string): void}} StorageLike
 */

const ORG_POLICY_PREFIX = "dlp:orgPolicy:";
const DOC_POLICY_PREFIX = "dlp:docPolicy:";

export class LocalPolicyStore {
  /**
   * @param {{storage: StorageLike}} params
   */
  constructor({ storage }) {
    if (!storage) throw new Error("LocalPolicyStore requires storage");
    this.storage = storage;
  }

  getOrgPolicy(orgId) {
    const raw = this.storage.getItem(`${ORG_POLICY_PREFIX}${orgId}`);
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch {
      return null;
    }
  }

  setOrgPolicy(orgId, policy) {
    validatePolicy(policy);
    this.storage.setItem(`${ORG_POLICY_PREFIX}${orgId}`, JSON.stringify(policy));
  }

  getDocumentPolicy(documentId) {
    const raw = this.storage.getItem(`${DOC_POLICY_PREFIX}${documentId}`);
    if (!raw) return null;
    try {
      return JSON.parse(raw);
    } catch {
      return null;
    }
  }

  setDocumentPolicy(documentId, policy) {
    validatePolicy(policy);
    this.storage.setItem(`${DOC_POLICY_PREFIX}${documentId}`, JSON.stringify(policy));
  }

  clearDocumentPolicy(documentId) {
    this.storage.removeItem(`${DOC_POLICY_PREFIX}${documentId}`);
  }
}

export class CloudOrgPolicyStore {
  /**
   * @param {{fetchImpl?: typeof fetch, baseUrl: string, authToken?: string}} params
   */
  constructor({ fetchImpl = fetch, baseUrl, authToken }) {
    if (!baseUrl) throw new Error("CloudOrgPolicyStore requires baseUrl");
    this.fetchImpl = fetchImpl;
    this.baseUrl = baseUrl.replace(/\/+$/, "");
    this.authToken = authToken;
  }

  headers() {
    const headers = { "content-type": "application/json" };
    if (this.authToken) headers.authorization = `Bearer ${this.authToken}`;
    return headers;
  }

  async get(orgId) {
    const res = await this.fetchImpl(`${this.baseUrl}/orgs/${encodeURIComponent(orgId)}/dlp-policy`, {
      method: "GET",
      headers: this.headers(),
    });
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`Failed to fetch org policy: ${res.status}`);
    const data = await res.json();
    return data?.policy ?? data;
  }

  async set(orgId, policy) {
    validatePolicy(policy);
    const res = await this.fetchImpl(`${this.baseUrl}/orgs/${encodeURIComponent(orgId)}/dlp-policy`, {
      method: "PUT",
      headers: this.headers(),
      body: JSON.stringify({ policy }),
    });
    if (!res.ok) throw new Error(`Failed to update org policy: ${res.status}`);
    const data = await res.json().catch(() => null);
    return data?.policy ?? policy;
  }
}

export class CloudDocumentPolicyStore {
  /**
   * @param {{fetchImpl?: typeof fetch, baseUrl: string, authToken?: string}} params
   */
  constructor({ fetchImpl = fetch, baseUrl, authToken }) {
    if (!baseUrl) throw new Error("CloudDocumentPolicyStore requires baseUrl");
    this.fetchImpl = fetchImpl;
    this.baseUrl = baseUrl.replace(/\/+$/, "");
    this.authToken = authToken;
  }

  headers() {
    const headers = { "content-type": "application/json" };
    if (this.authToken) headers.authorization = `Bearer ${this.authToken}`;
    return headers;
  }

  async get(documentId) {
    const res = await this.fetchImpl(`${this.baseUrl}/docs/${encodeURIComponent(documentId)}/dlp-policy`, {
      method: "GET",
      headers: this.headers(),
    });
    if (res.status === 404) return null;
    if (!res.ok) throw new Error(`Failed to fetch document policy: ${res.status}`);
    const data = await res.json();
    return data?.policy ?? data;
  }

  async set(documentId, policy) {
    validatePolicy(policy);
    const res = await this.fetchImpl(`${this.baseUrl}/docs/${encodeURIComponent(documentId)}/dlp-policy`, {
      method: "PUT",
      headers: this.headers(),
      body: JSON.stringify({ policy }),
    });
    if (!res.ok) throw new Error(`Failed to update document policy: ${res.status}`);
    const data = await res.json().catch(() => null);
    return data?.policy ?? policy;
  }
}

