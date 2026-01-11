import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { HashEmbedder, InMemoryVectorStore } from "../../ai-rag/src/index.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";
import { DlpViolationError } from "../../security/dlp/src/errors.js";

class CapturingEmbedder {
  /**
   * @param {{dimension: number}} opts
   */
  constructor(opts) {
    this.dimension = opts.dimension;
    /** @type {string[]} */
    this.seen = [];
  }

  /**
   * @param {string[]} texts
   */
  async embedTexts(texts) {
    this.seen.push(...texts);
    return texts.map(() => new Float32Array(this.dimension));
  }
}

function makeSensitiveWorkbook() {
  return {
    id: "wb-dlp",
    sheets: [
      {
        name: "Contacts",
        cells: [
          [{ v: "Name" }, { v: "Email" }, { v: "SSN" }],
          [{ v: "Alice" }, { v: "alice@example.com" }, { v: "123-45-6789" }],
        ],
      },
    ],
    tables: [
      {
        name: "ContactsTable",
        sheetName: "Contacts",
        rect: { r0: 0, c0: 0, r1: 1, c1: 2 },
      },
    ],
  };
}

function makePolicy({ maxAllowed = "Public", redactDisallowed }) {
  return {
    version: 1,
    allowDocumentOverrides: false,
    rules: {
      [DLP_ACTION.AI_CLOUD_PROCESSING]: {
        maxAllowed,
        allowRestrictedContent: false,
        redactDisallowed,
      },
    },
  };
}

test("buildWorkbookContext: blocks sensitive workbook chunks when policy blocks", async () => {
  const workbook = makeSensitiveWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 500,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const auditEvents = [];

  await assert.rejects(
    () =>
      cm.buildWorkbookContext({
        workbook,
        query: "alice@example.com",
        dlp: {
          documentId: workbook.id,
          policy: makePolicy({ redactDisallowed: false }),
          auditLogger: { log: (e) => auditEvents.push(e) },
        },
      }),
    (err) => {
      assert.ok(err instanceof DlpViolationError);
      return true;
    }
  );

  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0].type, "ai.workbook_context");
  assert.equal(auditEvents[0].documentId, workbook.id);
  assert.equal(auditEvents[0].decision.decision, "block");
});

test("buildWorkbookContext: redacts sensitive workbook chunks when policy allows with redaction", async () => {
  const workbook = makeSensitiveWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const auditEvents = [];

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "123-45-6789",
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ redactDisallowed: true }),
      auditLogger: { log: (e) => auditEvents.push(e) },
    },
  });

  assert.match(out.promptContext, /## dlp/i);
  assert.match(out.promptContext, /\[REDACTED_(EMAIL|SSN)\]/);
  assert.ok(out.retrieved.length > 0);
  assert.equal(out.retrieved[0].metadata.text, undefined);
  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0].type, "ai.workbook_context");
  assert.equal(auditEvents[0].documentId, workbook.id);
  assert.equal(auditEvents[0].decision.decision, "redact");
  assert.equal(auditEvents[0].redactedChunkCount, 1);
});

test("buildWorkbookContext: does not send raw sensitive workbook text to embedder when blocked", async () => {
  const workbook = makeSensitiveWorkbook();
  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 500,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  await assert.rejects(() =>
    cm.buildWorkbookContext({
      workbook,
      query: "contacts",
      dlp: {
        documentId: workbook.id,
        policy: makePolicy({ redactDisallowed: false }),
      },
    })
  );

  const joined = embedder.seen.join("\n");
  assert.doesNotMatch(joined, /\bAlice\b/);
  assert.doesNotMatch(joined, /alice@example\.com/);
  assert.doesNotMatch(joined, /123-45-6789/);
});

test("buildWorkbookContext: does not send raw sensitive workbook text to embedder when redacted", async () => {
  const workbook = makeSensitiveWorkbook();
  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  await cm.buildWorkbookContext({
    workbook,
    query: "contacts",
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ redactDisallowed: true }),
    },
  });

  const joined = embedder.seen.join("\n");
  assert.doesNotMatch(joined, /alice@example\.com/);
  assert.doesNotMatch(joined, /123-45-6789/);
  assert.match(joined, /\[REDACTED_(EMAIL|SSN)\]/);
});

test("buildWorkbookContext: redacts sensitive query before embedding when DLP is enabled", async () => {
  const workbook = makeSensitiveWorkbook();
  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  await cm.buildWorkbookContext({
    workbook,
    query: "Find alice@example.com",
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ redactDisallowed: true }),
    },
  });

  // embedder.seen includes: [chunk_text, query_text]
  assert.equal(embedder.seen.length, 2);
  assert.doesNotMatch(embedder.seen[1], /alice@example\.com/);
  assert.match(embedder.seen[1], /\[REDACTED_EMAIL\]/);
});

test("buildWorkbookContext: structured Restricted classifications fully redact retrieved chunks", async () => {
  const workbook = makeSensitiveWorkbook();
  // Add a value that isn't handled by the regex redactor but should still be suppressed by
  // explicit DLP classification.
  workbook.sheets[0].cells[1][0].v = "TopSecret";

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const auditEvents = [];

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "TopSecret",
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Confidential", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Contacts",
            range: { start: { row: 0, col: 0 }, end: { row: 1, col: 2 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
      auditLogger: { log: (e) => auditEvents.push(e) },
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.match(out.retrieved[0].text, /\[REDACTED\]/);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0].decision.decision, "redact");
});

test("buildWorkbookContext: structured Restricted classifications can block when policy requires", async () => {
  const workbook = makeSensitiveWorkbook();
  workbook.sheets[0].cells[1][0].v = "TopSecret";

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const auditEvents = [];

  await assert.rejects(
    () =>
      cm.buildWorkbookContext({
        workbook,
        query: "TopSecret",
        dlp: {
          documentId: workbook.id,
          policy: makePolicy({ maxAllowed: "Confidential", redactDisallowed: false }),
          classificationRecords: [
            {
              selector: {
                scope: "range",
                documentId: workbook.id,
                sheetId: "Contacts",
                range: { start: { row: 0, col: 0 }, end: { row: 1, col: 2 } },
              },
              classification: { level: "Restricted", labels: [] },
            },
          ],
          auditLogger: { log: (e) => auditEvents.push(e) },
        },
      }),
    (err) => {
      assert.ok(err instanceof DlpViolationError);
      return true;
    }
  );

  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0].type, "ai.workbook_context");
  assert.equal(auditEvents[0].decision.decision, "block");
});

test("buildWorkbookContext: classificationStore records are enforced for non-regex restricted content", async () => {
  const workbook = makeSensitiveWorkbook();
  workbook.sheets[0].cells[1][0].v = "TopSecret";

  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const classificationStore = {
    list(documentId) {
      assert.equal(documentId, workbook.id);
      return [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Contacts",
            range: { start: { row: 0, col: 0 }, end: { row: 1, col: 2 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ];
    },
  };

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "TopSecret",
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Confidential", redactDisallowed: true }),
      classificationStore,
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.match(out.retrieved[0].text, /\[REDACTED\]/);
  assert.doesNotMatch(out.promptContext, /TopSecret/);

  assert.ok(embedder.seen.length >= 1);
  // First embedder call is for the chunk text; the query may still contain the search term.
  assert.doesNotMatch(embedder.seen[0], /TopSecret/);
});
