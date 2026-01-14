import { describe, expect, it } from "vitest";

import { ContextManager } from "./contextManager.js";

describe("ContextManager promptContext", () => {
  it("is deterministic and uses compact stable JSON (no pretty indentation)", async () => {
    const cm = new ContextManager({
      // Avoid trimming so output shape/size checks are stable.
      tokenBudgetTokens: 1_000_000,
      // Disable redaction for this test so we can compare prompt strings directly.
      redactor: (text: string) => text,
    });

    const header = Array.from({ length: 8 }, (_v, idx) => `Col${idx + 1}`);
    const rows = Array.from({ length: 14 }, (_v, rIdx) =>
      Array.from({ length: 8 }, (_v2, cIdx) => `R${rIdx + 1}C${cIdx + 1}`),
    );

    const sheet = { name: "Sheet1", values: [header, ...rows] };
    const attachments = [{ type: "range" as const, reference: "Sheet1!A1:H15" }];

    const out1 = await cm.buildContext({ sheet, query: "col1", attachments });
    const out2 = await cm.buildContext({ sheet, query: "col1", attachments });

    expect(out1.promptContext).toEqual(out2.promptContext);
    expect(out1.promptContext).toContain("## schema_summary");
    expect(out1.promptContext).toContain("sheet=[Sheet1]");

    // The compact format should not include pretty-print indentation.
    expect(out1.promptContext).not.toContain('\n  "');

    const attachmentData = extractPromptSection(out1.promptContext, "attachment_data");
    const pretty = buildPrettyPromptContext({
      schema: out1.schema,
      attachmentData,
      attachments,
      sampledRows: out1.sampledRows,
      retrieved: out1.retrieved,
    });
    expect(out1.promptContext.length).toBeLessThan(pretty.length);
    expect(pretty.length - out1.promptContext.length).toBeGreaterThan(100);
  });

  it("does not inline raw range/table attachment data into promptContext", async () => {
    const cm = new ContextManager({
      tokenBudgetTokens: 1_000_000,
      // Disable redaction so we can assert directly on the prompt output.
      redactor: (text: string) => text,
    });

    const sheet = { name: "Sheet1", values: [[1]] };
    const secret = "TOP SECRET";
    const out = await cm.buildContext({
      sheet,
      query: "hi",
      attachments: [{ type: "table" as const, reference: "SalesTable", data: { snapshot: secret } }],
    });

    expect(out.promptContext).toContain("## attachments");
    expect(out.promptContext).toContain("SalesTable");
    expect(out.promptContext).not.toContain(secret);
  });

  it("does not leak schema column sampleValues into the prompt schema JSON", async () => {
    const cm = new ContextManager({
      tokenBudgetTokens: 1_000_000,
      redactor: (text: string) => text,
    });

    const values: any[][] = [["ID", "Secret"]];
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

    expect(out.schema.tables[0]?.columns?.[1]?.sampleValues).toContain("TOP_SECRET");
    expect(out.promptContext).not.toContain("TOP_SECRET");
  });
});

describe("ContextManager samplingStrategy", () => {
  it("supports samplingStrategy=head", async () => {
    const cm = new ContextManager({ tokenBudgetTokens: 1_000_000, redactor: (text: string) => text });
    const values = Array.from({ length: 10 }, (_v, i) => [`R${i}`]);
    const sheet = { name: "Sheet1", values };

    const out = await cm.buildContext({ sheet, query: "R", sampleRows: 3, samplingStrategy: "head" });
    expect(out.sampledRows).toEqual(values.slice(0, 3));
  });

  it("supports samplingStrategy=tail", async () => {
    const cm = new ContextManager({ tokenBudgetTokens: 1_000_000, redactor: (text: string) => text });
    const values = Array.from({ length: 10 }, (_v, i) => [`R${i}`]);
    const sheet = { name: "Sheet1", values };

    const out = await cm.buildContext({ sheet, query: "R", sampleRows: 3, samplingStrategy: "tail" });
    expect(out.sampledRows).toEqual(values.slice(-3));
  });

  it("supports samplingStrategy=systematic", async () => {
    const cm = new ContextManager({ tokenBudgetTokens: 1_000_000, redactor: (text: string) => text });
    const values = Array.from({ length: 10 }, (_v, i) => [`R${i}`]);
    const sheet = { name: "Sheet1", values };

    // `ContextManager` uses a fixed seed=1 for deterministic sampling.
    const out = await cm.buildContext({ sheet, query: "R", sampleRows: 4, samplingStrategy: "systematic" });
    expect(out.sampledRows).toEqual([values[1], values[4], values[6], values[9]]);
  });
});

function buildPrettyPromptContext(params: {
  schema: unknown;
  attachmentData: string;
  attachments: unknown[];
  sampledRows: unknown[][];
  retrieved: unknown[];
}): string {
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
      text: params.sampledRows.length
        ? `Sample rows:\n${params.sampledRows.map((r) => JSON.stringify(r)).join("\n")}`
        : "",
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

function extractPromptSection(promptContext: string, key: string): string {
  const marker = `## ${key}\n`;
  const start = promptContext.indexOf(marker);
  if (start === -1) return "";
  const rest = promptContext.slice(start + marker.length);
  const next = rest.indexOf("\n\n## ");
  return next === -1 ? rest : rest.slice(0, next);
}
