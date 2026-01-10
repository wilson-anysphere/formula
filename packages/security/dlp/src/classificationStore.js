import { normalizeClassification } from "./classification.js";
import { selectorKey } from "./selectors.js";

/**
 * @typedef {{getItem(key:string): (string|null), setItem(key:string,value:string): void, removeItem(key:string): void}} StorageLike
 */

export function createMemoryStorage() {
  const map = new Map();
  return {
    getItem(key) {
      return map.has(key) ? map.get(key) : null;
    },
    setItem(key, value) {
      map.set(key, String(value));
    },
    removeItem(key) {
      map.delete(key);
    },
  };
}

const STORAGE_PREFIX = "dlp:classifications:";

function storageKeyForDocument(documentId) {
  return `${STORAGE_PREFIX}${documentId}`;
}

/**
 * Stores classification metadata in a StorageLike implementation (browser localStorage
 * in production; an in-memory store in tests).
 */
export class LocalClassificationStore {
  /**
   * @param {{storage: StorageLike}} params
   */
  constructor({ storage }) {
    if (!storage) throw new Error("LocalClassificationStore requires storage");
    this.storage = storage;
  }

  /**
   * @param {string} documentId
   * @returns {Array<{selector:any, classification:any, updatedAt:string}>}
   */
  list(documentId) {
    const raw = this.storage.getItem(storageKeyForDocument(documentId));
    if (!raw) return [];
    try {
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return [];
      return parsed;
    } catch {
      return [];
    }
  }

  /**
   * @param {string} documentId
   * @param {any} selector
   * @param {any} classification
   */
  upsert(documentId, selector, classification) {
    const normalized = normalizeClassification(classification);
    const list = this.list(documentId);
    const key = selectorKey(selector);
    const now = new Date().toISOString();
    const next = list.filter((r) => selectorKey(r.selector) !== key);
    next.push({ selector, classification: normalized, updatedAt: now });
    this.storage.setItem(storageKeyForDocument(documentId), JSON.stringify(next));
  }

  /**
   * @param {string} documentId
   * @param {any} selector
   */
  remove(documentId, selector) {
    const list = this.list(documentId);
    const key = selectorKey(selector);
    const next = list.filter((r) => selectorKey(r.selector) !== key);
    this.storage.setItem(storageKeyForDocument(documentId), JSON.stringify(next));
  }
}

/**
 * Very small abstraction for a "cloud" persistence layer. In production this would
 * call `services/api` endpoints; here we keep the transport injectable to allow unit
 * testing without requiring a server framework.
 */
export class CloudClassificationStore {
  /**
   * @param {{fetchImpl?: typeof fetch, baseUrl: string, authToken?: string}} params
   */
  constructor({ fetchImpl = fetch, baseUrl, authToken }) {
    if (!baseUrl) throw new Error("CloudClassificationStore requires baseUrl");
    this.fetchImpl = fetchImpl;
    this.baseUrl = baseUrl.replace(/\/+$/, "");
    this.authToken = authToken;
  }

  headers() {
    const headers = { "content-type": "application/json" };
    if (this.authToken) headers.authorization = `Bearer ${this.authToken}`;
    return headers;
  }

  /**
   * @param {string} documentId
   */
  async list(documentId) {
    const res = await this.fetchImpl(`${this.baseUrl}/documents/${encodeURIComponent(documentId)}/classifications`, {
      method: "GET",
      headers: this.headers(),
    });
    if (!res.ok) throw new Error(`Failed to list classifications: ${res.status}`);
    const data = await res.json();
    if (!Array.isArray(data)) return [];
    return data;
  }

  /**
   * @param {string} documentId
   * @param {any} selector
   * @param {any} classification
   */
  async upsert(documentId, selector, classification) {
    const payload = { selector, classification: normalizeClassification(classification) };
    const res = await this.fetchImpl(`${this.baseUrl}/documents/${encodeURIComponent(documentId)}/classifications`, {
      method: "PUT",
      headers: this.headers(),
      body: JSON.stringify(payload),
    });
    if (!res.ok) throw new Error(`Failed to upsert classification: ${res.status}`);
  }

  /**
   * @param {string} documentId
   * @param {any} selector
   */
  async remove(documentId, selector) {
    const res = await this.fetchImpl(
      `${this.baseUrl}/documents/${encodeURIComponent(documentId)}/classifications/${encodeURIComponent(selectorKey(selector))}`,
      {
        method: "DELETE",
        headers: this.headers(),
      },
    );
    if (!res.ok) throw new Error(`Failed to delete classification: ${res.status}`);
  }
}

/**
 * Convenience store that writes to local storage and the cloud backend.
 *
 * If the cloud write fails, the local write remains. Callers can decide whether to
 * surface the error or retry in the background.
 */
export class HybridClassificationStore {
  /**
   * @param {{local: LocalClassificationStore, cloud: CloudClassificationStore}} params
   */
  constructor({ local, cloud }) {
    if (!local || !cloud) throw new Error("HybridClassificationStore requires local + cloud");
    this.local = local;
    this.cloud = cloud;
  }

  list(documentId) {
    return this.local.list(documentId);
  }

  async syncFromCloud(documentId) {
    const records = await this.cloud.list(documentId);
    // Overwrite local copy to match cloud state.
    for (const record of records) {
      this.local.upsert(documentId, record.selector, record.classification);
    }
    return this.local.list(documentId);
  }

  async upsert(documentId, selector, classification) {
    this.local.upsert(documentId, selector, classification);
    await this.cloud.upsert(documentId, selector, classification);
  }

  async remove(documentId, selector) {
    this.local.remove(documentId, selector);
    await this.cloud.remove(documentId, selector);
  }
}

