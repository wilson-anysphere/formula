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
  // Internal-only metadata used for structured DLP matching should never be returned to callers.
  assert.equal(out.retrieved[0].metadata.dlpSheetId, undefined);
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

test("buildWorkbookContext: structured DLP REDACT also drops attachment payload fields for non-heuristic strings (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-attachments-structured",
    sheets: [
      {
        name: "Sheet1",
        cells: [[{ v: "Hello" }]],
      },
    ],
  };

  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    workbookRag: { vectorStore, embedder, topK: 0 },
    redactor: (text) => text,
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "ignore",
    topK: 0,
    attachments: [{ type: "chart", reference: "Chart1", note: "TopSecret" }],
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Public", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
});

test("buildWorkbookContext: structured DLP REDACT also redacts non-heuristic attachment references (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-attachments-structured-reference",
    sheets: [
      {
        name: "Sheet1",
        cells: [[{ v: "Hello" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    workbookRag: { vectorStore, embedder, topK: 0 },
    redactor: (text) => text,
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "ignore",
    topK: 0,
    attachments: [{ type: "chart", reference: "TopSecret" }],
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Public", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: structured DLP REDACT also redacts non-string attachment references (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-attachments-structured-reference-object",
    sheets: [
      {
        name: "Sheet1",
        cells: [[{ v: "Hello" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    workbookRag: { vectorStore, embedder, topK: 0 },
    redactor: (text) => text,
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "ignore",
    topK: 0,
    // `reference` should be a string, but if a host passes a structured object, it must not leak.
    attachments: [{ type: "chart", reference: { token: "TopSecret" } }],
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Public", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: structured DLP REDACT also redacts non-object attachment entries (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-attachments-structured-reference-non-object",
    sheets: [
      {
        name: "Sheet1",
        cells: [[{ v: "Hello" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    workbookRag: { vectorStore, embedder, topK: 0 },
    redactor: (text) => text,
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "ignore",
    topK: 0,
    attachments: ["TopSecret"],
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Public", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: structured DLP REDACT also redacts non-array attachments payloads (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-attachments-structured-non-array",
    sheets: [
      {
        name: "Sheet1",
        cells: [[{ v: "Hello" }, { v: "World" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    workbookRag: { vectorStore, embedder, topK: 0 },
    redactor: (text) => text,
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "ignore",
    topK: 0,
    skipIndexing: true,
    skipIndexingWithDlp: true,
    // `attachments` should be an array, but callers may accidentally pass a scalar. Under structured
    // DLP redaction, treat it as prompt-unsafe to avoid leaking non-heuristic strings.
    attachments: "TopSecret",
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Public", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: structured DLP REDACT strips extra fields from rect metadata (no-op redactor)", async () => {
  const rectSecret = "RectSecretToken";
  const workbook = {
    id: "wb-dlp-rect-extra-fields",
    sheets: [
      {
        name: "Sheet1",
        cells: [[{ v: "Hello" }, { v: "World" }]],
      },
      {
        name: "OtherSheet",
        cells: [[{ v: "Ignore" }, { v: "Me" }]],
      },
    ],
    tables: [
      {
        name: "Table1",
        sheetName: "Sheet1",
        rect: { r0: 0, c0: 0, r1: 0, c1: 1, note: rectSecret },
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    workbookRag: { vectorStore, embedder, topK: 1 },
    redactor: (text) => text,
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "hello",
    topK: 1,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Public", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "OtherSheet",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.doesNotMatch(out.promptContext, new RegExp(rectSecret));
  assert.doesNotMatch(JSON.stringify(out.retrieved), new RegExp(rectSecret));
});

test("buildWorkbookContext: structured DLP REDACT drops unknown chunk metadata fields that can contain non-heuristic strings (no-op redactor)", async () => {
  const metaSecret = "MetaSecretToken";
  const workbook = {
    id: "wb-dlp-extra-metadata",
    sheets: [
      {
        name: "PublicSheet",
        // Needs at least 2 connected non-empty cells for ai-rag region detection to produce a chunk.
        cells: [[{ v: "Hello" }, { v: "public" }]],
      },
      {
        name: "SecretSheet",
        cells: [[{ v: "Ignore" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  // Pre-index the workbook, but persist an extra metadata field that contains a non-heuristic secret.
  // Under structured DLP redaction, ContextManager must not leak this field back to callers even if
  // its configured redactor is a no-op.
  await indexWorkbook({
    workbook,
    vectorStore,
    embedder,
    transform: (record) => ({
      metadata: { ...(record.metadata ?? {}), extraMeta: metaSecret },
    }),
  });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 1 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "public",
    topK: 1,
    skipIndexing: true,
    skipIndexingWithDlp: true,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      // Restrict a different sheet so the retrieved chunk is still allowed, but structuredOverallDecision is REDACT.
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "SecretSheet",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.doesNotMatch(out.promptContext, new RegExp(metaSecret));
  assert.doesNotMatch(JSON.stringify(out.retrieved), new RegExp(metaSecret));
  assert.equal(out.retrieved[0]?.metadata?.extraMeta, undefined);
});

test("buildWorkbookContext: structured DLP REDACT does not leak non-heuristic chunk kind strings from vector-store metadata (no-op redactor)", async () => {
  const kindSecret = "TopSecretKind";
  const workbook = {
    id: "wb-dlp-kind-secret",
    sheets: [
      {
        name: "PublicSheet",
        // Needs at least 2 connected non-empty cells for ai-rag region detection to produce a chunk.
        cells: [[{ v: "Hello" }, { v: "public" }]],
      },
      {
        name: "SecretSheet",
        cells: [[{ v: "Ignore" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  // Pre-index with a malicious/legacy metadata shape: override `metadata.kind` with an arbitrary string.
  // Under structured DLP redaction, ContextManager should not leak this identifier into prompts/outputs.
  await indexWorkbook({
    workbook,
    vectorStore,
    embedder,
    transform: (record) => ({
      metadata: { ...(record.metadata ?? {}), kind: kindSecret },
    }),
  });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 1 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "public",
    topK: 1,
    skipIndexing: true,
    skipIndexingWithDlp: true,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      // Restrict a different sheet so the retrieved chunk is still allowed, but structuredOverallDecision is REDACT.
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "SecretSheet",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.doesNotMatch(out.promptContext, new RegExp(kindSecret));
  assert.doesNotMatch(JSON.stringify(out.retrieved), new RegExp(kindSecret));
});

test("buildWorkbookContext: structured DLP REDACT also redacts non-heuristic titles for unknown chunk kinds (no-op redactor)", async () => {
  const titleSecret = "TopSecretTitle";
  const workbook = {
    id: "wb-dlp-unknown-kind-title",
    sheets: [
      {
        name: "PublicSheet",
        cells: [[{ v: "Hello" }, { v: "public" }]],
      },
      {
        name: "SecretSheet",
        cells: [[{ v: "Ignore" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  // Persist a legacy/third-party chunk metadata shape where `kind` is an unknown string.
  // Under structured DLP redaction, treat `title` as disallowed metadata too so non-heuristic
  // secrets cannot leak even if the configured redactor is a no-op.
  await indexWorkbook({
    workbook,
    vectorStore,
    embedder,
    transform: (record) => ({
      metadata: { ...(record.metadata ?? {}), kind: "custom", title: titleSecret },
    }),
  });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 1 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "public",
    topK: 1,
    skipIndexing: true,
    skipIndexingWithDlp: true,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "SecretSheet",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.doesNotMatch(out.promptContext, new RegExp(titleSecret));
  assert.doesNotMatch(JSON.stringify(out.retrieved), new RegExp(titleSecret));
  assert.equal(out.retrieved[0]?.metadata?.title, "[REDACTED]");
});

test("buildWorkbookContext: structured DLP REDACT does not return non-string title objects that can contain non-heuristic secrets (no-op redactor)", async () => {
  const titleSecret = "TopSecretTitleObject";
  const workbook = {
    id: "wb-dlp-title-object",
    sheets: [
      {
        name: "PublicSheet",
        cells: [[{ v: "Hello" }, { v: "public" }]],
      },
      {
        name: "SecretSheet",
        cells: [[{ v: "Ignore" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  await indexWorkbook({
    workbook,
    vectorStore,
    embedder,
    transform: (record) => ({
      metadata: { ...(record.metadata ?? {}), title: { token: titleSecret } },
    }),
  });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 1 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "public",
    topK: 1,
    skipIndexing: true,
    skipIndexingWithDlp: true,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "SecretSheet",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.doesNotMatch(JSON.stringify(out.retrieved), new RegExp(titleSecret));
  assert.equal(out.retrieved[0]?.metadata?.title, "[REDACTED]");
});

test("buildWorkbookContext: structured DLP REDACT does not leak non-string sheetName objects via custom toString (no-op redactor)", async () => {
  const sheetSecret = "TopSecretSheet";
  const workbook = {
    id: "wb-dlp-sheetname-tostring",
    sheets: [
      {
        name: "PublicSheet",
        // Needs at least 2 connected non-empty cells for ai-rag region detection to produce a chunk.
        cells: [[{ v: "Hello" }, { v: "public" }]],
      },
      {
        name: "SecretSheet",
        cells: [[{ v: "Ignore" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  // Persist a non-string sheetName with a custom toString(). This should never leak into
  // prompt context under structured DLP redaction, even if the configured redactor is a no-op.
  const sheetNameObj = Object.create(null);
  Object.defineProperty(sheetNameObj, "toString", { value: () => sheetSecret, enumerable: false });

  await indexWorkbook({
    workbook,
    vectorStore,
    embedder,
    transform: (record) => ({
      metadata: { ...(record.metadata ?? {}), sheetName: sheetNameObj },
    }),
  });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 1 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "public",
    topK: 1,
    skipIndexing: true,
    skipIndexingWithDlp: true,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      // Restrict a different sheet so the retrieved chunk is still allowed, but structuredOverallDecision is REDACT.
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "SecretSheet",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.doesNotMatch(out.promptContext, new RegExp(sheetSecret));
  assert.doesNotMatch(JSON.stringify(out.retrieved), new RegExp(sheetSecret));
});

test("buildWorkbookContext: structured DLP REDACT conservatively redacts chunk text when structured metadata is missing (no-op redactor)", async () => {
  const cellSecret = "TopSecret";
  const workbook = {
    id: "wb-dlp-missing-rect-metadata",
    sheets: [
      {
        name: "Sheet1",
        // Needs at least 2 connected non-empty cells for ai-rag region detection to produce a chunk.
        cells: [[{ v: cellSecret }, { v: "Hello" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  // Simulate a legacy/third-party vector store that omits `metadata.rect`, which is required
  // to apply structured range classifications to retrieved chunks.
  await indexWorkbook({
    workbook,
    vectorStore,
    embedder,
    transform: (record) => {
      const meta = record.metadata && typeof record.metadata === "object" ? record.metadata : {};
      const { rect: _rect, ...rest } = meta;
      return { metadata: rest };
    },
  });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 1 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "hello",
    topK: 1,
    skipIndexing: true,
    skipIndexingWithDlp: true,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.match(out.retrieved[0].text, /\[REDACTED\]/);
  assert.doesNotMatch(out.promptContext, new RegExp(cellSecret));
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

test("buildWorkbookContext: attachment-only sensitive patterns can trigger DLP REDACT even when workbook content is public", async () => {
  const workbook = {
    id: "wb-dlp-attachments-only",
    sheets: [
      {
        name: "PublicSheet",
        cells: [[{ v: "Hello" }], [{ v: "World" }]],
      },
    ],
  };
  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    workbookRag: { vectorStore, embedder, topK: 3 },
    // No-op redactor to ensure deep structured DLP redaction does not rely on regex helpers.
    redactor: (text) => text,
  });

  const auditEvents = [];

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "hello",
    attachments: [{ type: "chart", reference: "Chart1", data: { note: "987-65-4321" } }],
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ redactDisallowed: true }),
      auditLogger: { log: (e) => auditEvents.push(e) },
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /987-65-4321/);
  assert.match(out.promptContext, /\[REDACTED\]/);
  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0]?.decision?.decision, "redact");
});

test("buildWorkbookContext: workbook_summary redacts heuristic-sensitive sheet names even with a no-op redactor", async () => {
  const workbook = {
    id: "wb-dlp-summary-noop-redactor",
    sheets: [
      {
        // Heuristic-sensitive sheet name.
        name: "alice@example.com",
        cells: [[{ v: "Hello" }], [{ v: "World" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    // No-op redactor to ensure we still don't leak heuristic-sensitive strings under DLP.
    redactor: (t) => t,
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "hello",
    topK: 0, // force empty retrieval
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
    },
  });

  const summarySection =
    out.promptContext.match(/## workbook_summary\n([\s\S]*?)(?:\n\n## [^\n]+\n|$)/i)?.[1] ?? "";

  assert.match(out.promptContext, /## workbook_summary/i);
  assert.doesNotMatch(summarySection, /alice@example\.com/);
  assert.match(summarySection, /\[REDACTED\]/);
});

test("buildWorkbookContext: redacts heuristic-sensitive sheet names (schema + summary) even with a no-op redactor (DLP REDACT)", async () => {
  const workbook = {
    id: "wb-dlp-sheetname-email",
    sheets: [
      {
        name: "alice@example.com",
        cells: [[{ v: "Header" }], [{ v: "Value" }]],
      },
    ],
    tables: [
      {
        name: "ExampleTable",
        sheetName: "alice@example.com",
        rect: { r0: 0, c0: 0, r1: 1, c1: 0 },
      },
    ],
  };

  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1600,
    redactor: (t) => t,
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "ignore",
    topK: 0,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ redactDisallowed: true }),
    },
  });

  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
  assert.match(out.promptContext, /\[REDACTED\]/);

  const joined = embedder.seen.join("\n");
  assert.doesNotMatch(joined, /alice@example\.com/);
});

test("buildWorkbookContext: blocks when sheet names are heuristic-sensitive and policy blocks (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-sheetname-email-block",
    sheets: [
      {
        name: "alice@example.com",
        cells: [[{ v: "Header" }], [{ v: "Value" }]],
      },
    ],
    tables: [
      {
        name: "ExampleTable",
        sheetName: "alice@example.com",
        rect: { r0: 0, c0: 0, r1: 1, c1: 0 },
      },
    ],
  };

  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1600,
    redactor: (t) => t,
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  await assert.rejects(
    () =>
      cm.buildWorkbookContext({
        workbook,
        query: "ignore",
        topK: 0,
        dlp: {
          documentId: workbook.id,
          policy: makePolicy({ redactDisallowed: false }),
        },
      }),
    (err) => {
      assert.ok(err instanceof DlpViolationError);
      return true;
    },
  );

  const joined = embedder.seen.join("\n");
  assert.doesNotMatch(joined, /alice@example\.com/);
});

test("buildWorkbookContext: promptContext does not leak heuristic-sensitive sheet names in schema/retrieved sections (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-retrieved-sheetname-noop-redactor",
    sheets: [
      {
        name: "alice@example.com",
        cells: [[{ v: "Hello" }], [{ v: "World" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1500,
    redactor: (t) => t,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "hello",
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
  assert.doesNotMatch(out.retrieved[0]?.text ?? "", /alice@example\.com/);
  // Structured return fields should not leak the encoded chunk id either (ids include URL-encoded sheet names).
  assert.doesNotMatch(JSON.stringify(out.retrieved), /alice%40example\.com/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: does not redact heuristic-sensitive schema tokens when policy allows Restricted", async () => {
  const workbook = {
    id: "wb-dlp-allow-restricted",
    sheets: [
      {
        name: "Contacts",
        cells: [
          // Header row contains heuristic-sensitive strings; should be allowed when maxAllowed=Restricted.
          [{ v: "alice@example.com" }, { v: "123-45-6789" }, { v: "Amount" }],
          [{ v: "Alice" }, { v: "foo" }, { v: 100 }],
        ],
      },
    ],
    tables: [{ name: "SensitiveHeaders", sheetName: "Contacts", rect: { r0: 0, c0: 0, r1: 1, c1: 2 } }],
  };

  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1600,
    redactor: (t) => t,
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "ignore",
    topK: 0,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Restricted", redactDisallowed: true }),
    },
  });

  assert.match(out.promptContext, /## workbook_schema/i);
  const schemaSection =
    out.promptContext.match(/## workbook_schema\n([\s\S]*?)(?:\n\n## [^\n]+\n|$)/i)?.[1] ?? "";
  assert.match(schemaSection, /alice@example\.com/);
  assert.match(schemaSection, /123-45-6789/);
  assert.doesNotMatch(schemaSection, /\[REDACTED\]/);

  const joined = embedder.seen.join("\n");
  assert.match(joined, /alice@example\.com/);
  assert.match(joined, /123-45-6789/);
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

test("buildWorkbookContext: structured DLP also redacts non-heuristic sheet names in queries before embedding (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-structured-query-sheetname",
    sheets: [
      {
        name: "TopSecret",
        // Single-cell region -> no workbook chunks (chunkWorkbook drops 1-cell regions).
        cells: [[{ v: "Hello" }]],
      },
    ],
  };

  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 1 },
  });

  await cm.buildWorkbookContext({
    workbook,
    query: "Ask about TopSecret",
    topK: 1,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "TopSecret",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  // Only the query is embedded (no chunks were indexed).
  assert.equal(embedder.seen.length, 1);
  assert.doesNotMatch(embedder.seen[0], /TopSecret/);
  assert.match(embedder.seen[0], /\[REDACTED\]/);
});

test("buildWorkbookContext: structured DLP also redacts non-heuristic table names in queries before embedding (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-structured-query-tablename",
    sheets: [
      {
        name: "Sheet1",
        // Two disconnected 2-cell regions so indexing yields chunks.
        cells: [[{ v: "SecretA" }, { v: "SecretB" }, null, { v: "Hello" }, { v: "World" }]],
      },
    ],
    tables: [{ name: "TopSecretTable", sheetName: "Sheet1", rect: { r0: 0, c0: 3, r1: 0, c1: 4 } }],
  };

  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 1 },
  });

  await cm.buildWorkbookContext({
    workbook,
    query: "TopSecretTable",
    topK: 1,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  const joined = embedder.seen.join("\n");
  assert.doesNotMatch(joined, /TopSecretTable/);
});

test("buildWorkbookContext: structured DLP REDACT also redacts non-heuristic workbook ids in retrieved metadata/ids (no-op redactor)", async () => {
  const workbook = {
    id: "TopSecretWorkbook",
    sheets: [
      {
        name: "Sheet1",
        cells: [[{ v: "SecretA" }, { v: "SecretB" }]],
      },
      {
        name: "PublicSheet",
        cells: [[{ v: "Hello" }, { v: "World" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1600,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 1 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "Hello",
    topK: 1,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.doesNotMatch(out.promptContext, /TopSecretWorkbook/);
  assert.doesNotMatch(JSON.stringify(out.retrieved), /TopSecretWorkbook/);
  assert.match(out.promptContext, /\[REDACTED\]/);
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

test("buildWorkbookContext: includeRestrictedContent can block even when maxAllowed=Restricted (no-op redactor does not leak metadata)", async () => {
  const workbook = {
    id: "wb-dlp-include-restricted-block",
    sheets: [
      {
        // Heuristic-sensitive sheet name.
        name: "alice@example.com",
        cells: [[{ v: "X" }]],
      },
    ],
    tables: [{ name: "T", sheetName: "alice@example.com", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } }],
  };

  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 500,
    redactor: (t) => t, // no-op redactor: rely on ContextManager defense-in-depth
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  await assert.rejects(
    () =>
      cm.buildWorkbookContext({
        workbook,
        query: "ignore",
        topK: 0,
        dlp: {
          documentId: workbook.id,
          // maxAllowed=Restricted (threshold allows), but includeRestrictedContent=true requires allowRestrictedContent=true.
          policy: makePolicy({ maxAllowed: "Restricted", redactDisallowed: false }),
          includeRestrictedContent: true,
        },
      }),
    (err) => {
      assert.ok(err instanceof DlpViolationError);
      return true;
    },
  );

  const joined = embedder.seen.join("\n");
  assert.doesNotMatch(joined, /alice@example\.com/);
  assert.match(joined, /\[REDACTED\]/);
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

test("buildWorkbookContext: structured DLP redaction does not leak non-heuristic table/namedRange names (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-structured-metadata-tokens",
    sheets: [
      {
        name: "Sheet1",
        cells: [[{ v: "Hello" }]],
      },
    ],
    tables: [{ name: "TopSecret", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } }],
    namedRanges: [{ name: "TopSecretRange", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } }],
  };

  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t,
    workbookRag: { vectorStore, embedder, topK: 10 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "hello",
    topK: 10,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Public", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.ok(out.retrieved.length > 0);

  const joined = embedder.seen.join("\n");
  assert.doesNotMatch(joined, /TopSecret/);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.doesNotMatch(JSON.stringify(out.retrieved), /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: structured DLP REDACT also redacts non-heuristic sheet names inside range attachment references (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-structured-attachment-sheetname",
    sheets: [
      {
        name: "TopSecret",
        cells: [[{ v: "Hello" }]],
      },
    ],
    // Include at least one schema object so workbook_schema is populated without relying on vectorStore.list fallback.
    tables: [{ name: "T", sheetName: "TopSecret", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } }],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "hello",
    topK: 0,
    attachments: [{ type: "range", reference: "TopSecret!A1:A1" }],
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: { scope: "sheet", documentId: workbook.id, sheetId: "TopSecret" },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: range-level structured DLP REDACT also redacts non-heuristic sheet names (summary + attachments, no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-structured-range-sheetname",
    sheets: [
      {
        name: "TopSecret",
        cells: [[{ v: "Hello" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "hello",
    topK: 0,
    attachments: [{ type: "range", reference: "TopSecret!A1:A1" }],
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "TopSecret",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: structured DLP REDACT does not leak non-heuristic sheet names via retrieved/schema when an allowed chunk is retrieved (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-structured-sheetname-allowed-chunk",
    sheets: [
      {
        name: "TopSecret",
        // Two disconnected 2-cell regions (A1:B1 and D1:E1) so chunking yields two chunks.
        // The A1:B1 region is classified as Restricted; D1:E1 remains allowed and should still be retrievable.
        cells: [[{ v: "SecretA" }, { v: "SecretB" }, null, { v: "Hello" }, { v: "World" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1600,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 1 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "Hello",
    topK: 1,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "TopSecret",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.doesNotMatch(JSON.stringify(out.retrieved), /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: structured DLP REDACT does not send disallowed sheet-name metadata to embedder for allowed chunks (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-structured-sheetname-embedder",
    sheets: [
      {
        name: "TopSecret",
        cells: [[{ v: "SecretA" }, { v: "SecretB" }, null, { v: "Hello" }, { v: "World" }]],
      },
    ],
  };

  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  await cm.buildWorkbookContext({
    workbook,
    query: "ignore",
    topK: 0, // skip retrieval; we only care about what was sent to the embedder during indexing
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "TopSecret",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  const joined = embedder.seen.join("\n");
  assert.doesNotMatch(joined, /TopSecret/);
});

test("buildWorkbookContext: structured DLP REDACT also redacts non-heuristic table names when a disallowed range exists elsewhere on the sheet (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-structured-table-name-allowed",
    sheets: [
      {
        name: "Sheet1",
        cells: [
          // A1:B1 is classified as Restricted.
          [{ v: "SecretA" }, { v: "SecretB" }, null, { v: "Key" }, { v: "Value" }],
          [null, null, null, { v: "Hello" }, { v: "World" }],
        ],
      },
    ],
    // Table range is allowed (D1:E2), but the table name is a non-heuristic sensitive token.
    tables: [{ name: "TopSecretTable", sheetName: "Sheet1", rect: { r0: 0, c0: 3, r1: 1, c1: 4 } }],
  };

  const embedder = new CapturingEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1600,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "ignore",
    topK: 0,
    attachments: [{ type: "table", reference: "TopSecretTable" }],
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  const joined = embedder.seen.join("\n");
  assert.doesNotMatch(joined, /TopSecretTable/);
  assert.doesNotMatch(out.promptContext, /TopSecretTable/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: structured DLP REDACT also redacts non-heuristic sheet names in allowed range attachment references (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-structured-attachment-allowed-range",
    sheets: [
      {
        name: "TopSecret",
        // Two disconnected 2-cell regions (A1:B1 and D1:E1).
        cells: [[{ v: "SecretA" }, { v: "SecretB" }, null, { v: "Hello" }, { v: "World" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1600,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "ignore",
    topK: 0,
    attachments: [{ type: "range", reference: "TopSecret!D1:E1" }],
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "TopSecret",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 1 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.match(out.promptContext, /## attachments/i);
  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: structured DLP REDACT also redacts non-heuristic table attachment references (no-op redactor)", async () => {
  const workbook = {
    id: "wb-dlp-structured-table-attachment",
    sheets: [
      {
        name: "Sheet1",
        cells: [[{ v: "Hello" }]],
      },
    ],
    tables: [{ name: "TopSecret", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } }],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "hello",
    topK: 0,
    attachments: [{ type: "table", reference: "TopSecret" }],
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.doesNotMatch(out.promptContext, /TopSecret/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: structured DLP REDACT also redacts non-heuristic workbook ids (no-op redactor)", async () => {
  const workbook = {
    id: "TopSecretWorkbook",
    sheets: [
      {
        name: "Sheet1",
        cells: [[{ v: "Hello" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 0 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "hello",
    topK: 0,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "Sheet1",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.doesNotMatch(out.promptContext, /TopSecretWorkbook/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildWorkbookContext: structured DLP REDACT does not leak workbook ids via missing metadata.title (no-op redactor)", async () => {
  const workbook = {
    id: "TopSecretWorkbook",
    sheets: [
      {
        name: "PublicSheet",
        // Needs at least 2 connected non-empty cells for ai-rag region detection to produce a chunk.
        cells: [[{ v: "Hello" }, { v: "public" }]],
      },
      {
        name: "SecretSheet",
        cells: [[{ v: "Ignore" }]],
      },
    ],
  };

  const embedder = new HashEmbedder({ dimension: 64 });
  const vectorStore = new InMemoryVectorStore({ dimension: 64 });

  // Pre-index the workbook but simulate a legacy/third-party store that did not persist
  // `metadata.title`. In that case, ContextManager should not fall back to `hit.id` under
  // structured DLP redaction because the id embeds user-controlled workbook identifiers.
  await indexWorkbook({
    workbook,
    vectorStore,
    embedder,
    transform: (record) => {
      const meta = record.metadata && typeof record.metadata === "object" ? record.metadata : {};
      // Remove title entirely.
      const { title: _title, ...rest } = meta;
      return { metadata: rest };
    },
  });

  const cm = new ContextManager({
    tokenBudgetTokens: 1200,
    redactor: (t) => t, // no-op redactor
    workbookRag: { vectorStore, embedder, topK: 1 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "public",
    topK: 1,
    skipIndexing: true,
    skipIndexingWithDlp: true,
    dlp: {
      documentId: workbook.id,
      policy: makePolicy({ maxAllowed: "Internal", redactDisallowed: true }),
      // Restrict a different sheet so the retrieved chunk is still allowed, but structuredOverallDecision is REDACT.
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: workbook.id,
            sheetId: "SecretSheet",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.ok(out.retrieved.length > 0);
  assert.match(out.promptContext, /\[REDACTED\]/);
  assert.doesNotMatch(out.promptContext, /TopSecretWorkbook/);
  assert.doesNotMatch(JSON.stringify(out.retrieved), /TopSecretWorkbook/);
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
