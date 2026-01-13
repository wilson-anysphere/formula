import { describe, expect, it } from "vitest";

import { serializeToolResultForModel } from "./toolResultSerialization.js";

describe("serializeToolResultForModel", () => {
  it("read_range includes formulas preview + truncation metadata within budget", () => {
    const rows = 30;
    const cols = 15;
    const values = Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => r * cols + c));
    const formulas = Array.from({ length: rows }, (_, r) =>
      Array.from({ length: cols }, (_, c) => (r === 0 ? `=A${c + 1}` : null))
    );

    const toolCall = {
      id: "call-1",
      name: "read_range",
      arguments: { range: "Sheet1!A1:O30", include_formulas: true }
    };

    const result = {
      tool: "read_range",
      ok: true,
      timing: { started_at_ms: 0, duration_ms: 1 },
      data: { range: "Sheet1!A1:O30", values, formulas }
    };

    const maxChars = 5_000;
    const serialized = serializeToolResultForModel({ toolCall: toolCall as any, result, maxChars });
    expect(serialized.length).toBeLessThanOrEqual(maxChars);

    const payload = JSON.parse(serialized);
    expect(payload.tool).toBe("read_range");
    expect(payload.ok).toBe(true);
    expect(payload.data.shape).toEqual({ rows, cols });

    const previewRows = payload.data.values.length;
    const previewCols = payload.data.values[0]?.length ?? 0;
    expect(previewRows).toBeGreaterThan(0);
    expect(previewCols).toBeGreaterThan(0);

    expect(payload.data.formulas).toBeDefined();
    expect(payload.data.formulas.length).toBe(previewRows);
    expect(payload.data.formulas[0]?.length ?? 0).toBe(previewCols);

    expect(payload.data.truncated).toBe(true);
    expect(payload.data.truncated_rows).toBe(Math.max(0, rows - previewRows));
    expect(payload.data.truncated_cols).toBe(Math.max(0, cols - previewCols));
  });

  it("filter_range truncates matching_rows while preserving count within budget", () => {
    const matchingRows = Array.from({ length: 2_000 }, (_, i) => i + 1);

    const toolCall = {
      id: "call-2",
      name: "filter_range",
      arguments: { range: "Sheet1!A1:D2000" }
    };

    const result = {
      tool: "filter_range",
      ok: true,
      timing: { started_at_ms: 0, duration_ms: 1 },
      data: {
        range: "Sheet1!A1:D2000",
        matching_rows: matchingRows,
        count: matchingRows.length
      }
    };

    const maxChars = 1_000;
    const serialized = serializeToolResultForModel({ toolCall: toolCall as any, result, maxChars });
    expect(serialized.length).toBeLessThanOrEqual(maxChars);

    const payload = JSON.parse(serialized);
    expect(payload.tool).toBe("filter_range");
    expect(payload.ok).toBe(true);
    expect(payload.data.range).toBe("Sheet1!A1:D2000");
    expect(payload.data.count).toBe(2_000);
    expect(payload.data.matching_rows.length).toBeLessThan(payload.data.count);
    expect(payload.data.truncated).toBe(true);
  });

  it("detect_anomalies truncates large anomaly lists and reports method + counts within budget", () => {
    const anomalies = Array.from({ length: 1_000 }, (_, i) => ({
      cell: `Sheet1!A${i + 1}`,
      value: i,
      score: 0.9
    }));

    const toolCall = {
      id: "call-3",
      name: "detect_anomalies",
      arguments: { range: "Sheet1!A1:A5000", method: "zscore" }
    };

    const result = {
      tool: "detect_anomalies",
      ok: true,
      timing: { started_at_ms: 0, duration_ms: 1 },
      data: {
        // Intentionally omit range to ensure we fall back to toolCall.arguments.range.
        method: "zscore",
        anomalies,
        truncated: true,
        total_anomalies: 5_000
      }
    };

    const maxChars = 1_500;
    const serialized = serializeToolResultForModel({ toolCall: toolCall as any, result, maxChars });
    expect(serialized.length).toBeLessThanOrEqual(maxChars);

    const payload = JSON.parse(serialized);
    expect(payload.tool).toBe("detect_anomalies");
    expect(payload.ok).toBe(true);
    expect(payload.data.range).toBe("Sheet1!A1:A5000");
    expect(payload.data.method).toBe("zscore");
    expect(payload.data.total_anomalies).toBe(5_000);
    expect(payload.data.anomalies.length).toBeGreaterThan(0);
    expect(payload.data.anomalies.length).toBeLessThan(1_000);
    expect(payload.data.anomalies[0]).toEqual(expect.objectContaining({ cell: "Sheet1!A1", value: 0, score: 0.9 }));
    expect(payload.data.truncated).toBe(true);
  });

  it("generic serializer handles circular references + huge strings without throwing", () => {
    const huge = "x".repeat(100_000);
    const circular: any = { tool: "some_tool", ok: true, data: { huge } };
    circular.data.self = circular;

    const toolCall = { id: "call-4", name: "some_tool", arguments: {} };
    const maxChars = 800;

    const serialized = serializeToolResultForModel({ toolCall: toolCall as any, result: circular, maxChars });
    expect(serialized.length).toBeLessThanOrEqual(maxChars);

    expect(() => JSON.parse(serialized)).not.toThrow();
  });
});

