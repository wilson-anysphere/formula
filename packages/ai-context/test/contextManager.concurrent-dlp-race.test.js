import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { RagIndex } from "../src/rag.js";

class GateEmbedder {
  constructor() {
    /** @type {string[]} */
    this.seen = [];
    /** @type {() => void} */
    this.release = () => {};
    this.gate = new Promise((resolve) => {
      this.release = resolve;
    });
    /** @type {() => void} */
    this._queryStartedResolve = () => {};
    this.queryStarted = new Promise((resolve) => {
      this._queryStartedResolve = resolve;
    });
  }

  /**
   * @param {string} text
   */
  async embed(text) {
    this.seen.push(text);
    if (text === "DELAYED_QUERY") {
      this._queryStartedResolve();
      await this.gate;
    }
    // Deterministic embedding; dimension is irrelevant for these tests.
    return [1];
  }
}

test("buildContext: concurrent re-indexing cannot swap the store between DLP redaction and retrieval", async () => {
  const embedder = new GateEmbedder();
  const ragIndex = new RagIndex({ embedder });

  const cm = new ContextManager({
    tokenBudgetTokens: 500,
    ragIndex,
    // No-op redactor: ensure the test catches leaks that aren't covered by heuristic redaction.
    redactor: (t) => t,
  });

  const sheet = {
    name: "Sheet1",
    values: [["TopSecret"]],
  };

  const documentId = "doc-1";
  const policy = {
    version: 1,
    allowDocumentOverrides: false,
    rules: {
      "ai.cloudProcessing": {
        maxAllowed: "Public",
        allowRestrictedContent: false,
        redactDisallowed: true,
      },
    },
  };

  const dlp = {
    documentId,
    sheetId: "Sheet1",
    policy,
    classificationRecords: [
      {
        selector: {
          scope: "range",
          documentId,
          sheetId: "Sheet1",
          range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
        },
        classification: { level: "Restricted", labels: [] },
      },
    ],
  };

  const p1 = cm.buildContext({
    sheet,
    query: "DELAYED_QUERY",
    dlp,
  });

  // Wait until the first call reaches query embedding. In the buggy implementation, this
  // happens *after* indexing releases the per-sheet lock, which allows concurrent calls to
  // re-index the sheet with different (unredacted) content.
  await embedder.queryStarted;

  const p2 = cm.buildContext({
    sheet,
    query: "OTHER_QUERY",
  });

  // Let p2 reach the lock/index path before allowing p1 to continue.
  await new Promise((resolve) => setTimeout(resolve, 0));
  embedder.release();

  const out1 = await p1;
  await p2;

  assert.doesNotMatch(out1.promptContext, /TopSecret/);
  assert.doesNotMatch(JSON.stringify(out1.retrieved), /TopSecret/);
});

