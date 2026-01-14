import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { InMemoryVectorStore } from "../../ai-rag/src/store/inMemoryVectorStore.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";

function makePolicy({ maxAllowed = "Internal", redactDisallowed = true } = {}) {
  return {
    version: 1,
    allowDocumentOverrides: true,
    rules: {
      [DLP_ACTION.AI_CLOUD_PROCESSING]: {
        maxAllowed,
        allowRestrictedContent: false,
        redactDisallowed,
      },
    },
  };
}

function instrumentRecordList(records) {
  let passes = 0;
  let elementGets = 0;
  let propGets = 0;
  /** @type {WeakMap<object, any>} */
  const objectProxyCache = new WeakMap();

  /**
   * Wrap a plain object in a Proxy that counts property reads (recursively).
   *
   * This catches regressions where callers clone the record array once and then scan
   * the clone per chunk, bypassing array-level scan counters.
   *
   * @param {any} value
   * @returns {any}
   */
  function wrapObject(value) {
    if (!value || typeof value !== "object") return value;
    if (Array.isArray(value)) return value;
    // Avoid proxying built-ins with internal slots (Map/Set/Date) since methods can throw
    // when `this` is a Proxy.
    if (value instanceof Map || value instanceof Set || value instanceof Date) return value;
    const cached = objectProxyCache.get(value);
    if (cached) return cached;
    const proxy = new Proxy(value, {
      get(target, prop, receiver) {
        propGets += 1;
        return wrapObject(Reflect.get(target, prop, receiver));
      },
    });
    objectProxyCache.set(value, proxy);
    return proxy;
  }

  const wrappedRecords = (records ?? []).map((r) => wrapObject(r));
  const proxy = new Proxy(wrappedRecords, {
    get(target, prop, receiver) {
      if (prop === Symbol.iterator) {
        return function () {
          passes += 1;
          // Bind the iterator to the proxy so element reads go through this trap.
          return Array.prototype[Symbol.iterator].call(receiver);
        };
      }
      if (typeof prop === "string" && /^[0-9]+$/.test(prop)) {
        elementGets += 1;
      }
      return Reflect.get(target, prop, receiver);
    },
  });

  return { proxy, getPasses: () => passes, getElementGets: () => elementGets, getPropGets: () => propGets };
}

class ZeroEmbedder {
  constructor(dimension) {
    this.dimension = dimension;
  }
  async embedTexts(texts) {
    return texts.map(() => new Float32Array(this.dimension));
  }
}

test("ContextManager.buildWorkbookContext: avoids scanning classification records per indexed chunk during indexing transform", async () => {
  const workbookId = "wb-dlp-index-perf-indexing";
  const sheetName = "Sheet1";

  const dimension = 8;
  const vectorStore = new InMemoryVectorStore({ dimension });
  const embedder = new ZeroEmbedder(dimension);

  // Create many disconnected non-empty regions so chunkWorkbook emits many dataRegion chunks.
  // Each region contains at least two adjacent cells (chunkWorkbook intentionally drops
  // single-cell regions) and is separated by at least one empty row so regions don't merge.
  const rows = 120;
  const cells = Array.from({ length: rows }, () => []);
  for (let i = 0; i < 50; i++) {
    const row = i * 2;
    cells[row][0] = { v: `v${i}` };
    cells[row][1] = { v: `w${i}` };
  }

  const workbook = {
    id: workbookId,
    sheets: [{ name: sheetName, cells }],
    tables: [],
    namedRanges: [],
  };

  // Provide a modest set of structured records. The expected fast path builds a document index once,
  // then uses it for every chunk in the `indexWorkbook` transform. Any per-chunk scan regression
  // would multiply record reads by ~50.
  const classificationRecords = Array.from({ length: 10 }, (_, idx) => ({
    selector: { scope: "document", documentId: workbookId, idx },
    classification: { level: "Public", labels: [] },
  }));
  const { proxy: recordsProxy, getPasses, getElementGets, getPropGets } = instrumentRecordList(classificationRecords);

  const cm = new ContextManager({
    tokenBudgetTokens: 500,
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "hello",
    topK: 0,
    includePromptContext: false,
    dlp: {
      documentId: workbookId,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: recordsProxy,
    },
  });

  assert.equal(out.retrieved.length, 0);
  assert.ok(out.indexStats);
  assert.ok(out.indexStats.totalChunks >= 30, `expected many chunks, got ${out.indexStats.totalChunks}`);

  // Expect only a handful of linear scans (overall classification + document index build).
  // Any per-chunk structured scan regression would exceed these by orders of magnitude.
  assert.ok(getPasses() < 30, `expected < 30 record iteration passes, got ${getPasses()}`);
  assert.ok(getElementGets() < 400, `expected < 400 record element reads, got ${getElementGets()}`);
  assert.ok(getPropGets() < 2_000, `expected < 2000 record property reads, got ${getPropGets()}`);
});
