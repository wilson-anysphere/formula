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
  const proxy = new Proxy(records, {
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
  return { proxy, getPasses: () => passes, getElementGets: () => elementGets };
}

class ZeroEmbedder {
  constructor(dimension) {
    this.dimension = dimension;
  }
  async embedTexts(texts) {
    return texts.map(() => new Float32Array(this.dimension));
  }
}

test("ContextManager.buildWorkbookContext: avoids scanning classification records per retrieved chunk (uses document index)", async () => {
  const workbookId = "wb-dlp-index-perf";
  const sheetName = "Sheet1";

  const dimension = 8;
  const vectorStore = new InMemoryVectorStore({ dimension });
  const embedder = new ZeroEmbedder(dimension);

  // Pre-populate the vector store with enough unique, non-overlapping chunks so retrieval
  // returns many hits (amplifying any per-hit structured DLP scans).
  const records = [];
  for (let i = 0; i < 60; i++) {
    records.push({
      id: `chunk-${String(i).padStart(3, "0")}`,
      vector: new Float32Array(dimension),
      metadata: {
        workbookId,
        sheetName,
        rect: { r0: i, c0: 0, r1: i, c1: 0 },
        kind: "chunk",
        title: `Row ${i + 1}`,
        text: `Row ${i + 1}`,
      },
    });
  }
  await vectorStore.upsert(records);

  // Minimal structured DLP record set. If `buildWorkbookContext` regresses to calling
  // `effectiveRangeClassification(..., classificationRecords)` for each hit, we'd see one
  // pass over this list per retrieved chunk.
  const classificationRecords = [
    {
      selector: { scope: "document", documentId: workbookId },
      classification: { level: "Public", labels: [] },
    },
  ];
  const { proxy: recordsProxy, getPasses, getElementGets } = instrumentRecordList(classificationRecords);

  const cm = new ContextManager({
    tokenBudgetTokens: 500,
    workbookRag: { vectorStore, embedder, topK: 40 },
  });

  const out = await cm.buildWorkbookContext({
    workbook: { id: workbookId, sheets: [{ name: sheetName }] },
    query: "hello",
    topK: 40,
    // Avoid extra prompt formatting logic that can invoke additional structured checks;
    // this test focuses on per-hit classification.
    includePromptContext: false,
    skipIndexing: true,
    skipIndexingWithDlp: true,
    dlp: {
      documentId: workbookId,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: recordsProxy,
    },
  });

  assert.equal(out.retrieved.length, 40);

  // With the document index fast path, we expect only a small number of linear passes over
  // the classification record list (for overall classification + index build). If the code
  // falls back to scanning `classificationRecords` for each hit, this would jump above 40.
  assert.ok(getPasses() < 20, `expected < 20 record iteration passes, got ${getPasses()}`);
  // Defense-in-depth: catch per-hit scans even if implemented without Symbol.iterator.
  assert.ok(getElementGets() < 20, `expected < 20 record element reads, got ${getElementGets()}`);
});
