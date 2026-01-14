import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { HashEmbedder } from "../../ai-rag/src/embedding/hashEmbedder.js";
import { indexWorkbook } from "../../ai-rag/src/pipeline/indexWorkbook.js";
import { InMemoryVectorStore } from "../../ai-rag/src/store/inMemoryVectorStore.js";
import { workbookFromSpreadsheetApi } from "../../ai-rag/src/workbook/fromSpreadsheetApi.js";
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

function makeTopSecretWorkbook() {
  return {
    id: "wb-dlp-structured",
    sheets: [
      {
        name: "Secrets",
        cells: [[{ v: "Label" }], [{ v: "TopSecret" }]],
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

test("buildWorkbookContext: workbook_schema does not contain raw sensitive strings when DLP redaction is enabled", async () => {
  const workbook = makeSensitiveWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "contacts",
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ redactDisallowed: true }),
    },
  });

  assert.match(out.promptContext, /## workbook_schema/i);
  assert.match(out.promptContext, /ContactsTable/);
  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
  assert.doesNotMatch(out.promptContext, /123-45-6789/);
});

test("buildWorkbookContext: workbook_schema redacts sensitive header strings when DLP redaction is enabled", async () => {
  const workbook = {
    id: "wb-dlp-schema-header",
    sheets: [
      {
        name: "Contacts",
        cells: [
          // Force header detection by ensuring the next row contains a numeric value.
          [{ v: "Name" }, { v: "alice@example.com" }, { v: "123-45-6789" }, { v: "Amount" }],
          [{ v: "Alice" }, { v: "foo" }, { v: "bar" }, { v: 100 }],
        ],
      },
    ],
    tables: [{ name: "HeaderSensitiveTable", sheetName: "Contacts", rect: { r0: 0, c0: 0, r1: 1, c1: 3 } }],
  };

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "contacts",
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ redactDisallowed: true }),
    },
  });

  const schemaSection =
    out.promptContext.match(/## workbook_schema\n([\s\S]*?)(?:\n\n## [^\n]+\n|$)/i)?.[1] ?? "";

  assert.match(out.promptContext, /## workbook_schema/i);
  assert.doesNotMatch(schemaSection, /alice@example\.com/);
  assert.doesNotMatch(schemaSection, /123-45-6789/);
  assert.match(schemaSection, /\[REDACTED_EMAIL\]/);
  assert.match(schemaSection, /\[REDACTED_SSN\]/);
});

test("buildWorkbookContext: attachments are redacted under DLP REDACT even when redactor is a no-op", async () => {
  const workbook = makeSensitiveWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    workbookRag: { vectorStore, embedder, topK: 3 },
    redactor: (text) => text,
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "contacts",
    attachments: [{ type: "chart", reference: "Chart1", data: { note: "987-65-4321" } }],
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ redactDisallowed: true }),
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /987-65-4321/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: workbook_schema redacts sensitive header strings even with a no-op redactor (DLP REDACT + empty retrieval)", async () => {
  const workbook = {
    id: "wb-dlp-schema-header-noop-redactor",
    sheets: [
      {
        name: "Contacts",
        cells: [
          // Force header detection by ensuring the next row contains a numeric value.
          [{ v: "Name" }, { v: "alice@example.com" }, { v: "123-45-6789" }, { v: "Amount" }],
          [{ v: "Alice" }, { v: "foo" }, { v: "bar" }, { v: 100 }],
        ],
      },
    ],
    tables: [{ name: "HeaderSensitiveTable", sheetName: "Contacts", rect: { r0: 0, c0: 0, r1: 1, c1: 3 } }],
  };

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  // Use a no-op redactor to ensure ContextManager still enforces heuristic safety under DLP.
  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t,
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "contacts",
    topK: 0, // force empty retrieval so schema is the only cell-derived section
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ redactDisallowed: true }),
    },
  });

  assert.equal(out.retrieved.length, 0);
  const schemaSection =
    out.promptContext.match(/## workbook_schema\n([\s\S]*?)(?:\n\n## [^\n]+\n|$)/i)?.[1] ?? "";

  assert.match(out.promptContext, /## workbook_schema/i);
  assert.doesNotMatch(schemaSection, /alice@example\.com/);
  assert.doesNotMatch(schemaSection, /123-45-6789/);
  assert.match(schemaSection, /\[REDACTED\]/);
});

test("buildWorkbookContext: workbook_schema fallback redacts sensitive columns even with a no-op redactor (DLP REDACT)", async () => {
  const workbookId = "wb-api-schema-header-noop-redactor";

  const spreadsheet = {
    listSheets() {
      return ["Contacts"];
    },
    listNonEmptyCells(sheet) {
      assert.equal(sheet, "Contacts");
      return [
        // Header row contains heuristic-sensitive strings (email + SSN).
        { address: { sheet: "Contacts", row: 1, col: 1 }, cell: { value: "alice@example.com" } },
        { address: { sheet: "Contacts", row: 1, col: 2 }, cell: { value: "123-45-6789" } },
        { address: { sheet: "Contacts", row: 1, col: 3 }, cell: { value: "Amount" } },
        // Data row ensures header detection + type inference.
        { address: { sheet: "Contacts", row: 2, col: 1 }, cell: { value: "Alice" } },
        { address: { sheet: "Contacts", row: 2, col: 2 }, cell: { value: "foo" } },
        { address: { sheet: "Contacts", row: 2, col: 3 }, cell: { value: 100 } },
      ];
    },
  };

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  // Pre-index without DLP so the persisted chunk metadata contains the raw header values.
  const workbook = workbookFromSpreadsheetApi({ spreadsheet, workbookId, coordinateBase: "one" });
  await indexWorkbook({ workbook, vectorStore, embedder });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t,
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  const out = await cm.buildWorkbookContextFromSpreadsheetApi({
    spreadsheet,
    workbookId,
    query: "ignore",
    topK: 0, // force no retrieval (so schema comes from vectorStore.list fallback)
    skipIndexing: true,
    skipIndexingWithDlp: true,
    dlp: {
      documentId: workbookId,
      policy: makePolicy({ redactDisallowed: true }),
    },
  });

  assert.equal(out.retrieved.length, 0);
  assert.match(out.promptContext, /## workbook_schema/i);
  const schemaSection =
    out.promptContext.match(/## workbook_schema\n([\s\S]*?)(?:\n\n## [^\n]+\n|$)/i)?.[1] ?? "";
  assert.doesNotMatch(schemaSection, /alice@example\.com/);
  assert.doesNotMatch(schemaSection, /123-45-6789/);
  assert.match(schemaSection, /\[REDACTED\]/);
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

test("buildWorkbookContext: redacts sensitive query before embedding even when includeRestrictedContent=true but policy blocks", async () => {
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
      query: "alice@example.com",
      dlp: {
        documentId: workbook.id,
        policy: makePolicy({ redactDisallowed: false }),
        includeRestrictedContent: true,
      },
    })
  );

  // embedder.seen includes [chunk_text, query_text]; both should be safe.
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

test("buildWorkbookContext: supports snake_case DLP options (structured block + audit documentId)", async () => {
  const workbook = makeTopSecretWorkbook();
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
          document_id: workbook.id,
          policy: makePolicy({ maxAllowed: "Confidential", redactDisallowed: false }),
          classification_records: [
            {
              selector: {
                scope: "range",
                documentId: workbook.id,
                sheetId: "Secrets",
                range: { start: { row: 0, col: 0 }, end: { row: 1, col: 0 } },
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
  assert.equal(auditEvents[0].documentId, workbook.id);
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

test("buildWorkbookContextFromSpreadsheetApi: blocks sensitive workbook chunks when policy blocks", async () => {
  const spreadsheet = {
    listSheets() {
      return ["Contacts"];
    },
    listNonEmptyCells(sheet) {
      assert.equal(sheet, "Contacts");
      return [
        { address: { sheet: "Contacts", row: 1, col: 1 }, cell: { value: "Name" } },
        { address: { sheet: "Contacts", row: 1, col: 2 }, cell: { value: "Email" } },
        { address: { sheet: "Contacts", row: 1, col: 3 }, cell: { value: "SSN" } },
        { address: { sheet: "Contacts", row: 2, col: 1 }, cell: { value: "Alice" } },
        { address: { sheet: "Contacts", row: 2, col: 2 }, cell: { value: "alice@example.com" } },
        { address: { sheet: "Contacts", row: 2, col: 3 }, cell: { value: "123-45-6789" } },
      ];
    },
  };

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 500,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  await assert.rejects(
    () =>
      cm.buildWorkbookContextFromSpreadsheetApi({
        spreadsheet,
        workbookId: "wb-api-dlp",
        query: "alice@example.com",
        dlp: {
          documentId: "wb-api-dlp",
          policy: makePolicy({ redactDisallowed: false }),
        },
      }),
    (err) => {
      assert.ok(err instanceof DlpViolationError);
      return true;
    }
  );
});

test("buildWorkbookContextFromSpreadsheetApi: redacts sensitive workbook chunks when policy allows with redaction", async () => {
  const spreadsheet = {
    listSheets() {
      return ["Contacts"];
    },
    listNonEmptyCells(sheet) {
      assert.equal(sheet, "Contacts");
      return [
        { address: { sheet: "Contacts", row: 1, col: 1 }, cell: { value: "Name" } },
        { address: { sheet: "Contacts", row: 1, col: 2 }, cell: { value: "Email" } },
        { address: { sheet: "Contacts", row: 1, col: 3 }, cell: { value: "SSN" } },
        { address: { sheet: "Contacts", row: 2, col: 1 }, cell: { value: "Alice" } },
        { address: { sheet: "Contacts", row: 2, col: 2 }, cell: { value: "alice@example.com" } },
        { address: { sheet: "Contacts", row: 2, col: 3 }, cell: { value: "123-45-6789" } },
      ];
    },
  };

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const auditEvents = [];

  const out = await cm.buildWorkbookContextFromSpreadsheetApi({
    spreadsheet,
    workbookId: "wb-api-dlp-redact",
    query: "123-45-6789",
    dlp: {
      documentId: "wb-api-dlp-redact",
      policy: makePolicy({ redactDisallowed: true }),
      auditLogger: { log: (e) => auditEvents.push(e) },
    },
  });

  assert.match(out.promptContext, /## dlp/i);
  assert.match(out.promptContext, /\[REDACTED_(EMAIL|SSN)\]/);
  assert.match(out.promptContext, /## workbook_schema/i);
  const schemaSection =
    out.promptContext.match(/## workbook_schema\n([\s\S]*?)(?:\n\n## [^\n]+\n|$)/i)?.[1] ?? "";
  assert.doesNotMatch(schemaSection, /alice@example\.com/);
  assert.doesNotMatch(schemaSection, /123-45-6789/);
  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0].type, "ai.workbook_context");
  assert.equal(auditEvents[0].decision.decision, "redact");
});

test("buildWorkbookContext: workbook_schema fallback respects structured DLP even when skipIndexingWithDlp=true", async () => {
  const workbookId = "wb-api-schema-structured";

  const spreadsheet = {
    listSheets() {
      return ["Secrets"];
    },
    listNonEmptyCells(sheet) {
      assert.equal(sheet, "Secrets");
      return [
        // Header row contains non-regex "TopSecret" (should be suppressed by structured DLP).
        { address: { sheet: "Secrets", row: 1, col: 1 }, cell: { value: "TopSecret" } },
        { address: { sheet: "Secrets", row: 1, col: 2 }, cell: { value: "Value" } },
        { address: { sheet: "Secrets", row: 2, col: 1 }, cell: { value: "A" } },
        { address: { sheet: "Secrets", row: 2, col: 2 }, cell: { value: 1 } },
      ];
    },
  };

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  // Pre-index without DLP so the persisted chunk metadata contains the raw header.
  const workbook = workbookFromSpreadsheetApi({ spreadsheet, workbookId, coordinateBase: "one" });
  await indexWorkbook({ workbook, vectorStore, embedder });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const out = await cm.buildWorkbookContextFromSpreadsheetApi({
    spreadsheet,
    workbookId,
    query: "ignore",
    topK: 0, // force no retrieval (so schema comes from vectorStore.list fallback)
    skipIndexing: true,
    skipIndexingWithDlp: true,
    dlp: {
      documentId: workbookId,
      policy: makePolicy({ maxAllowed: "Confidential", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbookId,
            sheetId: "Secrets",
            range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.equal(out.retrieved.length, 0);
  assert.match(out.promptContext, /## workbook_schema/i);
  const schemaSection =
    out.promptContext.match(/## workbook_schema\n([\s\S]*?)(?:\n\n## [^\n]+\n|$)/i)?.[1] ?? "";
  assert.doesNotMatch(schemaSection, /TopSecret/);
  assert.match(schemaSection, /\[REDACTED\]/);
});

test("buildWorkbookContextFromSpreadsheetApi: supports snake_case DLP options (structured block + audit documentId)", async () => {
  const spreadsheet = {
    listSheets() {
      return ["Secrets"];
    },
    listNonEmptyCells(sheet) {
      assert.equal(sheet, "Secrets");
      return [
        { address: { sheet: "Secrets", row: 1, col: 1 }, cell: { value: "Label" } },
        { address: { sheet: "Secrets", row: 2, col: 1 }, cell: { value: "TopSecret" } },
      ];
    },
  };

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const auditEvents = [];

  await assert.rejects(
    () =>
      cm.buildWorkbookContextFromSpreadsheetApi({
        spreadsheet,
        workbookId: "wb-api-dlp-structured",
        query: "TopSecret",
        dlp: {
          document_id: "wb-api-dlp-structured",
          policy: makePolicy({ maxAllowed: "Confidential", redactDisallowed: false }),
          classification_records: [
            {
              selector: {
                scope: "range",
                documentId: "wb-api-dlp-structured",
                sheetId: "Secrets",
                range: { start: { row: 0, col: 0 }, end: { row: 1, col: 0 } },
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
  assert.equal(auditEvents[0].documentId, "wb-api-dlp-structured");
  assert.equal(auditEvents[0].decision.decision, "block");
});

test("buildWorkbookContext: does not rely on persisted dlpHeuristic metadata for block decisions", async () => {
  const workbook = makeSensitiveWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  function firstLine(text) {
    const s = String(text ?? "");
    const idx = s.indexOf("\n");
    return idx === -1 ? s : s.slice(0, idx);
  }

  // Simulate a legacy / third-party index that stored only a redacted placeholder and
  // did not persist `metadata.dlpHeuristic`. `buildWorkbookContext` should still block
  // based on the current workbook content + policy.
  await indexWorkbook({
    workbook,
    vectorStore,
    embedder,
    transform: (record) => {
      const placeholder = `${firstLine(record.text)}\n[REDACTED]`;
      return { text: placeholder, metadata: { ...(record.metadata ?? {}), text: placeholder } };
    },
  });

  const cm = new ContextManager({
    tokenBudgetTokens: 500,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  await assert.rejects(
    () =>
      cm.buildWorkbookContext({
        workbook,
        query: "ContactsTable",
        dlp: {
          documentId: workbook.id,
          policy: makePolicy({ redactDisallowed: false }),
        },
      }),
    (err) => {
      assert.ok(err instanceof DlpViolationError);
      return true;
    }
  );
});
