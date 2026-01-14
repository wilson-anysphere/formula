import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";

test("ContextManager promptContext: is deterministic and uses compact stable JSON (no pretty indentation)", async () => {
  const cm = new ContextManager({
    // Avoid trimming so output shape/size checks are stable.
    tokenBudgetTokens: 1_000_000,
    // Disable redaction for this test so we can compare prompt strings directly.
    redactor: (text) => text,
  });

  const header = Array.from({ length: 8 }, (_v, idx) => `Col${idx + 1}`);
  const rows = Array.from({ length: 14 }, (_v, rIdx) => Array.from({ length: 8 }, (_v2, cIdx) => `R${rIdx + 1}C${cIdx + 1}`));

  const sheet = { name: "Sheet1", values: [header, ...rows] };
  const attachments = [{ type: "range", reference: "Sheet1!A1:H15" }];

  const out1 = await cm.buildContext({ sheet, query: "col1", attachments });
  const out2 = await cm.buildContext({ sheet, query: "col1", attachments });

  assert.equal(out1.promptContext, out2.promptContext);
  assert.ok(out1.promptContext.includes("## schema_summary"));
  assert.ok(out1.promptContext.includes("sheet=[Sheet1]"));

  // The compact format should not include pretty-print indentation.
  assert.ok(!out1.promptContext.includes('\n  "'));

  const attachmentData = extractPromptSection(out1.promptContext, "attachment_data");
  const pretty = buildPrettyPromptContext({
    schema: out1.schema,
    attachmentData,
    attachments,
    sampledRows: out1.sampledRows,
    retrieved: out1.retrieved,
  });
  assert.ok(out1.promptContext.length < pretty.length);
  assert.ok(pretty.length - out1.promptContext.length > 100);
});

test("ContextManager promptContext: does not inline raw range/table attachment data into promptContext", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    // Disable redaction so we can assert directly on the prompt output.
    redactor: (text) => text,
  });

  const sheet = { name: "Sheet1", values: [[1]] };
  const secret = "TOP SECRET";
  const out = await cm.buildContext({
    sheet,
    query: "hi",
    attachments: [{ type: "table", reference: "SalesTable", data: { snapshot: secret } }],
  });

  assert.ok(out.promptContext.includes("## attachments"));
  assert.ok(out.promptContext.includes("SalesTable"));
  assert.ok(!out.promptContext.includes(secret));
});

test("ContextManager promptContext: does not leak schema column sampleValues into the prompt schema JSON", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (text) => text,
  });

  const values = [["ID", "Secret"]];
  // Keep the secret value outside the first few rows so the retrieved TSV preview
  // (bounded by maxChunkRows) does not contain it. `extractSheetSchema` will still
  // capture it in `TableSchema.columns[*].sampleValues`.
  for (let i = 0; i < 50; i++) {
    values.push([i + 1, i === 49 ? "TOP_SECRET" : 0]);
  }

  const sheet = { name: "Sheet1", values };
  const out = await cm.buildContext({
    sheet,
    query: "id",
    sampleRows: 0,
    limits: { maxChunkRows: 5 },
  });

  assert.ok(out.schema.tables[0]?.columns?.[1]?.sampleValues?.includes("TOP_SECRET"));
  assert.ok(!out.promptContext.includes("TOP_SECRET"));
});

test("ContextManager samplingStrategy: supports samplingStrategy=head", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000_000, redactor: (text) => text });
  const values = Array.from({ length: 10 }, (_v, i) => [`R${i}`]);
  const sheet = { name: "Sheet1", values };

  const out = await cm.buildContext({ sheet, query: "R", sampleRows: 3, samplingStrategy: "head" });
  assert.deepStrictEqual(out.sampledRows, values.slice(0, 3));
});

test("ContextManager samplingStrategy: supports samplingStrategy=tail", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000_000, redactor: (text) => text });
  const values = Array.from({ length: 10 }, (_v, i) => [`R${i}`]);
  const sheet = { name: "Sheet1", values };

  const out = await cm.buildContext({ sheet, query: "R", sampleRows: 3, samplingStrategy: "tail" });
  assert.deepStrictEqual(out.sampledRows, values.slice(-3));
});

test("ContextManager samplingStrategy: supports samplingStrategy=systematic", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000_000, redactor: (text) => text });
  const values = Array.from({ length: 10 }, (_v, i) => [`R${i}`]);
  const sheet = { name: "Sheet1", values };

  // `ContextManager` uses a fixed seed=1 for deterministic sampling.
  const out = await cm.buildContext({ sheet, query: "R", sampleRows: 4, samplingStrategy: "systematic" });
  assert.deepStrictEqual(out.sampledRows, [values[1], values[4], values[6], values[9]]);
});

function buildPrettyPromptContext(params) {
  const sections = [
    {
      key: "attachment_data",
      priority: 4.5,
      text: params.attachmentData ? params.attachmentData : "",
    },
    {
      key: "schema",
      priority: 3,
      text: `Sheet schema (schema-first):\n${JSON.stringify(params.schema, null, 2)}`,
    },
    {
      key: "attachments",
      priority: 2,
      text: params.attachments.length ? `User-provided attachments:\n${JSON.stringify(params.attachments, null, 2)}` : "",
    },
    {
      key: "samples",
      priority: 1,
      text: params.sampledRows.length ? `Sample rows:\n${params.sampledRows.map((r) => JSON.stringify(r)).join("\n")}` : "",
    },
    {
      key: "retrieved",
      priority: 4,
      text: params.retrieved.length ? `Retrieved context:\n${JSON.stringify(params.retrieved, null, 2)}` : "",
    },
  ].filter((s) => s.text);

  sections.sort((a, b) => b.priority - a.priority);
  return sections.map((s) => `## ${s.key}\n${s.text}`).join("\n\n");
}

function extractPromptSection(promptContext, key) {
  const marker = `## ${key}\n`;
  const start = promptContext.indexOf(marker);
  if (start === -1) return "";
  const rest = promptContext.slice(start + marker.length);
  const next = rest.indexOf("\n\n## ");
  return next === -1 ? rest : rest.slice(0, next);
}

