import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { InMemoryVectorStore, RagIndex } from "../src/rag.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";

class CapturingEmbedder {
  constructor() {
    /** @type {string[]} */
    this.seen = [];
  }

  /**
   * @param {string} text
   */
  async embed(text) {
    this.seen.push(String(text));
    return [1];
  }
}

test("buildContext: redacts sensitive query before sheet-level RAG embedding when DLP is enabled", async () => {
  const embedder = new CapturingEmbedder();
  const ragIndex = new RagIndex({ embedder, store: new InMemoryVectorStore() });

  // Use a no-op redactor to ensure ContextManager still enforces heuristic safety under DLP.
  const cm = new ContextManager({ tokenBudgetTokens: 1_000, ragIndex, redactor: (text) => text });

  await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [["Header"], ["Value"]],
    },
    query: "Find alice@example.com",
    dlp: {
      documentId: "doc-1",
      sheetId: "Sheet1",
      policy: {
        version: 1,
        allowDocumentOverrides: true,
        rules: {
          [DLP_ACTION.AI_CLOUD_PROCESSING]: {
            maxAllowed: "Internal",
            allowRestrictedContent: false,
            redactDisallowed: true,
          },
        },
      },
      classificationRecords: [],
    },
  });

  // The query embedding happens after indexing, so the last embedded string should be the query.
  assert.equal(embedder.seen.at(-1), "[REDACTED]");
  assert.doesNotMatch(embedder.seen.join("\n"), /alice@example\.com/);
});

