import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { InMemoryVectorStore, RagIndex } from "../src/rag.js";
import { DLP_ACTION } from "../../security/dlp/src/actions.js";
import { DlpViolationError } from "../../security/dlp/src/errors.js";

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

test("buildContext: heuristic DLP detects rich object cell values and prevents them from leaking to embeddings (no-op redactor)", async () => {
  const embedder = new CapturingEmbedder();
  const ragIndex = new RagIndex({ embedder, store: new InMemoryVectorStore() });

  const cm = new ContextManager({ tokenBudgetTokens: 1_000, ragIndex, redactor: (text) => text });

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [
        // Rich text cells (DocumentController-style) use `{ text, runs }`.
        [{ text: "Email", runs: [] }, { text: "alice@example.com", runs: [] }],
        [{ text: "Other", runs: [] }, { text: "Hello", runs: [] }],
      ],
    },
    query: "anything",
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

  assert.doesNotMatch(embedder.seen.join("\n"), /alice@example\.com/);
  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildContext: heuristic DLP detects class instance cell values via toString and prevents them from leaking (no-op redactor)", async () => {
  const embedder = new CapturingEmbedder();
  const ragIndex = new RagIndex({ embedder, store: new InMemoryVectorStore() });

  const cm = new ContextManager({ tokenBudgetTokens: 1_000, ragIndex, redactor: (text) => text });

  class SecretValue {
    toString() {
      return "alice@example.com";
    }
  }

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [[new SecretValue(), "Hello"]],
    },
    query: "anything",
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

  assert.doesNotMatch(embedder.seen.join("\n"), /alice@example\.com/);
  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
  assert.match(out.promptContext, /\[REDACTED\]/);
});

test("buildContext: heuristic DLP detects plain-object toString overrides and prevents them from leaking (no-op redactor)", async () => {
  const embedder = new CapturingEmbedder();
  const ragIndex = new RagIndex({ embedder, store: new InMemoryVectorStore() });

  const cm = new ContextManager({ tokenBudgetTokens: 1_000, ragIndex, redactor: (text) => text });

  const sneaky = {
    // Override the stringification path that `valuesRangeToTsv` uses for some object values.
    // This should never leak under DLP REDACT.
    toString() {
      return "alice@example.com";
    },
  };

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      // Needs at least 2 connected non-empty cells for region detection.
      values: [[sneaky, "Hello"]],
    },
    query: "anything",
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

  assert.doesNotMatch(embedder.seen.join("\n"), /alice@example\.com/);
  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
  assert.match(out.promptContext, /## dlp/i);
});

test("buildContext: heuristic DLP detects Symbol.toPrimitive cell values and prevents them from leaking (no-op redactor)", async () => {
  const embedder = new CapturingEmbedder();
  const ragIndex = new RagIndex({ embedder, store: new InMemoryVectorStore() });

  const cm = new ContextManager({ tokenBudgetTokens: 1_000, ragIndex, redactor: (text) => text });

  const sneaky = {};
  sneaky[Symbol.toPrimitive] = () => "alice@example.com";

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [[sneaky, "Hello"]],
    },
    query: "anything",
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

  assert.doesNotMatch(embedder.seen.join("\n"), /alice@example\.com/);
  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
  assert.match(out.promptContext, /## dlp/i);
});

test("buildContext: heuristic DLP does not leak Date.toISOString overrides (no-op redactor)", async () => {
  const embedder = new CapturingEmbedder();
  const ragIndex = new RagIndex({ embedder, store: new InMemoryVectorStore() });

  const cm = new ContextManager({ tokenBudgetTokens: 1_000, ragIndex, redactor: (text) => text });

  const sneakyDate = new Date("2024-01-01T00:00:00.000Z");
  // Some hosts (or malicious input) can attach custom methods to otherwise-safe values.
  // TSV + JSON formatting should never call these overrides in a way that leaks the secret.
  sneakyDate.toISOString = () => "alice@example.com";

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      // Needs at least 2 connected non-empty cells for region detection.
      values: [[sneakyDate, "Hello"]],
    },
    query: "anything",
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

  assert.doesNotMatch(embedder.seen.join("\n"), /alice@example\.com/);
  assert.doesNotMatch(out.promptContext, /alice@example\.com/);
});

test("buildContext: structured DLP also redacts non-heuristic sheet-name tokens in queries before embedding (no-op redactor)", async () => {
  const embedder = new CapturingEmbedder();
  const ragIndex = new RagIndex({ embedder, store: new InMemoryVectorStore() });

  const cm = new ContextManager({ tokenBudgetTokens: 1_000, ragIndex, redactor: (text) => text });

  await cm.buildContext({
    sheet: {
      name: "TopSecretSheet",
      values: [["Hello"]],
    },
    query: "Search TopSecretSheet",
    dlp: {
      documentId: "doc-1",
      sheetId: "TopSecretSheet",
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
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: "doc-1",
            sheetId: "TopSecretSheet",
            range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.doesNotMatch(embedder.seen.at(-1), /TopSecretSheet/);
  assert.match(embedder.seen.at(-1), /\[REDACTED\]/);
});

test("buildContext: structured DLP also redacts non-heuristic table/namedRange tokens in queries before embedding (no-op redactor)", async () => {
  const embedder = new CapturingEmbedder();
  const ragIndex = new RagIndex({ embedder, store: new InMemoryVectorStore() });

  const cm = new ContextManager({ tokenBudgetTokens: 1_000, ragIndex, redactor: (text) => text });

  await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [
        ["Hello", "World"],
        ["a", "b"],
      ],
      tables: [{ name: "TopSecretTable", range: "Sheet1!A1:A1" }],
      namedRanges: [{ name: "TopSecretRange", range: "Sheet1!A1:A1" }],
    },
    query: "Describe TopSecretTable and TopSecretRange",
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
      classificationRecords: [
        {
          selector: {
            scope: "range",
            documentId: "doc-1",
            sheetId: "Sheet1",
            // Structured redaction is triggered by a disallowed cell elsewhere in the window;
            // table/namedRange identifiers should still be treated as disallowed metadata tokens.
            range: { start: { row: 1, col: 1 }, end: { row: 1, col: 1 } }, // B2
          },
          classification: { level: "Restricted", labels: [] },
        },
      ],
    },
  });

  assert.doesNotMatch(embedder.seen.at(-1), /TopSecretTable/);
  assert.doesNotMatch(embedder.seen.at(-1), /TopSecretRange/);
  assert.match(embedder.seen.at(-1), /\[REDACTED\]/);
});

test("buildContext: heuristic Restricted findings can block when policy requires", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000, redactor: (text) => text });

  /** @type {any[]} */
  const auditEvents = [];
  const auditLogger = { log: (event) => auditEvents.push(event) };

  await assert.rejects(
    cm.buildContext({
      sheet: {
        name: "Sheet1",
        values: [["Email"], ["alice@example.com"]],
      },
      query: "anything",
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
              redactDisallowed: false,
            },
          },
        },
        classificationRecords: [],
        auditLogger,
      },
    }),
    (err) => err instanceof DlpViolationError,
  );

  assert.equal(auditEvents.length, 1);
  assert.equal(auditEvents[0]?.type, "ai.context");
  assert.equal(cm.ragIndex.store.size, 0);
});
