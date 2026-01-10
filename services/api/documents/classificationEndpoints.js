import { selectorKey } from "../../../packages/security/dlp/src/selectors.js";
import { normalizeClassification } from "../../../packages/security/dlp/src/classification.js";

export class InMemoryDocumentClassificationStore {
  constructor() {
    this.byDocumentId = new Map();
  }

  list(documentId) {
    return this.byDocumentId.get(documentId) || [];
  }

  upsert(documentId, selector, classification) {
    const list = this.list(documentId);
    const key = selectorKey(selector);
    const next = list.filter((r) => selectorKey(r.selector) !== key);
    next.push({ selector, classification: normalizeClassification(classification), updatedAt: new Date().toISOString() });
    this.byDocumentId.set(documentId, next);
  }

  remove(documentId, selector) {
    const list = this.list(documentId);
    const key = selectorKey(selector);
    const next = list.filter((r) => selectorKey(r.selector) !== key);
    this.byDocumentId.set(documentId, next);
  }
}

export function createDocumentClassificationEndpoints({ store }) {
  if (!store) throw new Error("createDocumentClassificationEndpoints requires store");

  return {
    /**
     * GET /documents/:documentId/classifications
     */
    async list({ documentId }) {
      return { status: 200, body: store.list(documentId) };
    },

    /**
     * PUT /documents/:documentId/classifications
     */
    async upsert({ documentId, selector, classification }) {
      store.upsert(documentId, selector, classification);
      return { status: 204, body: null };
    },

    /**
     * DELETE /documents/:documentId/classifications/:selectorKey
     */
    async remove({ documentId, selectorKey: key }) {
      const existing = store.list(documentId);
      const record = existing.find((r) => selectorKey(r.selector) === key);
      if (!record) return { status: 404, body: { error: "not_found" } };
      store.remove(documentId, record.selector);
      return { status: 204, body: null };
    },
  };
}

