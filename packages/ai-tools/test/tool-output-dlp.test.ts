import { describe, expect, it } from "vitest";

import { InMemoryWorkbook } from "../src/spreadsheet/in-memory-workbook.js";
import { parseA1Cell } from "../src/spreadsheet/a1.js";
import { SpreadsheetLLMToolExecutor } from "../src/llm/integration.js";

import { DLP_ACTION } from "../../security/dlp/src/actions.js";

function makePolicy({ maxAllowed = "Public", redactDisallowed }: { maxAllowed?: string; redactDisallowed: boolean }) {
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

describe("tool output DLP enforcement", () => {
  it("does not alter tool results when DLP is not configured", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Public" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: null, formula: '="TopSecret"' });

    const executor = new SpreadsheetLLMToolExecutor(workbook);
    const result = await executor.execute({
      name: "read_range",
      arguments: { range: "Sheet1!A1:B1", include_formulas: true },
    });

    expect(result.ok).toBe(true);
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");
    expect(result.data?.formulas?.[0]?.[1]).toBe('="TopSecret"');
  });

  it("redacts restricted cells in read_range (values + formulas)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Public" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: null, formula: '="TopSecret"' });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: 1 });
    workbook.setCell(parseA1Cell("Sheet1!B2"), { value: 2 });

    const executor = new SpreadsheetLLMToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: makePolicy({ redactDisallowed: true }),
        classification_records: [
          {
            selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row: 0, col: 1 },
            classification: { level: "Restricted", labels: [] },
          },
        ],
      },
    });

    const result = await executor.execute({
      name: "read_range",
      arguments: { range: "Sheet1!A1:B2", include_formulas: true },
    });

    expect(result.ok).toBe(true);
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    // Mixed allowed + restricted cells: only B1 should be redacted.
    expect(result.data?.values).toEqual([
      ["Public", "[REDACTED]"],
      [1, 2],
    ]);
    expect(result.data?.formulas).toEqual([
      [null, "[REDACTED]"],
      [null, null],
    ]);

    const serialized = JSON.stringify(result);
    expect(serialized).not.toContain("TopSecret");
  });

  it("redacts heuristically sensitive values in read_range even without classification records", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "user@example.com" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: "Public" });

    const executor = new SpreadsheetLLMToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: makePolicy({ redactDisallowed: true }),
        classification_records: [],
      },
    });

    const result = await executor.execute({
      name: "read_range",
      arguments: { range: "Sheet1!A1:B1" },
    });

    expect(result.ok).toBe(true);
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([["[REDACTED]", "Public"]]);
    expect(JSON.stringify(result)).not.toContain("user@example.com");
  });

  it("redacts heuristically sensitive formulas in read_range even without classification records", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: null, formula: '="user@example.com"' });

    const executor = new SpreadsheetLLMToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: makePolicy({ redactDisallowed: true }),
        classification_records: [],
      },
    });

    const result = await executor.execute({
      name: "read_range",
      arguments: { range: "Sheet1!A1:A1", include_formulas: true },
    });

    expect(result.ok).toBe(true);
    if (!result.ok || result.tool !== "read_range") throw new Error("Unexpected tool result");

    expect(result.data?.values).toEqual([["[REDACTED]"]]);
    expect(result.data?.formulas).toEqual([["[REDACTED]"]]);
    expect(JSON.stringify(result)).not.toContain("user@example.com");
  });

  it("blocks tool results when policy disallows cloud processing (BLOCK)", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: "Public" });
    workbook.setCell(parseA1Cell("Sheet1!B1"), { value: "TopSecret" });

    const executor = new SpreadsheetLLMToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: makePolicy({ redactDisallowed: false }),
        classification_records: [
          {
            selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row: 0, col: 1 },
            classification: { level: "Restricted", labels: [] },
          },
        ],
      },
    });

    const result = await executor.execute({
      name: "read_range",
      arguments: { range: "Sheet1!A1:B1", include_formulas: true },
    });

    expect(result.ok).toBe(false);
    expect(result.error?.code).toBe("permission_denied");

    // Defense-in-depth: the blocked result must not contain raw restricted values.
    expect(JSON.stringify(result)).not.toContain("TopSecret");
  });

  it("redacts derived tool outputs (compute_statistics) when selection is REDACT", async () => {
    const workbook = new InMemoryWorkbook(["Sheet1"]);
    workbook.setCell(parseA1Cell("Sheet1!A1"), { value: 10 });
    workbook.setCell(parseA1Cell("Sheet1!A2"), { value: 9999 });
    workbook.setCell(parseA1Cell("Sheet1!A3"), { value: 30 });

    const executor = new SpreadsheetLLMToolExecutor(workbook, {
      dlp: {
        document_id: "doc-1",
        policy: makePolicy({ redactDisallowed: true }),
        classification_records: [
          {
            selector: { scope: "cell", documentId: "doc-1", sheetId: "Sheet1", row: 1, col: 0 },
            classification: { level: "Restricted", labels: [] },
          },
        ],
      },
    });

    const result = await executor.execute({
      name: "compute_statistics",
      arguments: { range: "Sheet1!A1:A3", measures: ["mean", "min", "max"] },
    });

    expect(result.ok).toBe(true);
    if (!result.ok || result.tool !== "compute_statistics") throw new Error("Unexpected tool result");
    expect(result.data?.statistics).toEqual({ mean: 20, min: 10, max: 30 });
    // Use a word-boundary regex so floating-point timing fields that contain "...999999..."
    // don't create false positives.
    expect(JSON.stringify(result)).not.toMatch(/\b9999\b/);
  });
});
