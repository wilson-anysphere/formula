import { describe, expect, it } from "vitest";

import { serializeToolResultForModel } from "./toolResultSerialization.js";

describe("serializeToolResultForModel", () => {
  it("trims tool names in tool execution envelopes", () => {
    const toolCall = {
      id: "call-trim",
      name: "read_range",
      arguments: { range: "Sheet1!A1:A1" }
    };

    const result = {
      tool: "  read_range  ",
      ok: true,
      data: { values: [[1]] }
    };

    const serialized = serializeToolResultForModel({ toolCall: toolCall as any, result, maxChars: 5_000 });
    const payload = JSON.parse(serialized);
    expect(payload.tool).toBe("read_range");
  });

  it("uses specialized serializers even when toolCall.name includes whitespace", () => {
    const toolCall = {
      id: "call-trim-2",
      name: "  read_range  ",
      arguments: { range: "Sheet1!A1:A1" }
    };

    const result = {
      tool: "read_range",
      ok: true,
      data: { values: [[1]] }
    };

    const serialized = serializeToolResultForModel({ toolCall: toolCall as any, result, maxChars: 5_000 });
    const payload = JSON.parse(serialized);

    expect(payload.tool).toBe("read_range");
    // `read_range` serializer always includes a `shape` summary (generic serializer does not).
    expect(payload.data.shape).toEqual({ rows: 1, cols: 1 });
  });

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
    expect(payload.timing).toEqual({ started_at_ms: 0, duration_ms: 1 });
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

  it("read_range falls back to toolCall.arguments.range when result omits range", () => {
    const toolCall = {
      id: "call-1b",
      name: "read_range",
      arguments: { range: "Sheet1!A1:B2", include_formulas: true }
    };

    const result = {
      tool: "read_range",
      ok: true,
      timing: { started_at_ms: 0, duration_ms: 1 },
      data: {
        values: [
          [1, 2],
          [3, 4]
        ],
        formulas: [
          ["=1", "=2"],
          [null, null]
        ]
      }
    };

    const serialized = serializeToolResultForModel({ toolCall: toolCall as any, result, maxChars: 5_000 });
    const payload = JSON.parse(serialized);
    expect(payload.tool).toBe("read_range");
    expect(payload.data.range).toBe("Sheet1!A1:B2");
    expect(payload.data.formulas).toBeDefined();
  });

  it("read_range can fall back to a minimal summary when the tool envelope is too large", () => {
    const toolCall = {
      id: "call-1c",
      name: "read_range",
      arguments: { range: "Sheet1!A1:B2", include_formulas: true }
    };

    const result = {
      tool: "read_range",
      ok: true,
      timing: { started_at_ms: 0, duration_ms: 1 },
      warnings: ["x".repeat(5_000)],
      data: {
        values: [
          [1, 2],
          [3, 4]
        ],
        formulas: [
          ["=1", "=2"],
          [null, null]
        ]
      }
    };

    const maxChars = 350;
    const serialized = serializeToolResultForModel({ toolCall: toolCall as any, result, maxChars });
    expect(serialized.length).toBeLessThanOrEqual(maxChars);

    const payload = JSON.parse(serialized);
    expect(payload.tool).toBe("read_range");
    expect(payload.ok).toBe(true);
    // Fallback drops huge envelope fields like warnings.
    expect(payload.warnings).toBeUndefined();
    expect(payload.timing).toEqual({ started_at_ms: 0, duration_ms: 1 });
    expect(payload.data).toEqual({
      range: "Sheet1!A1:B2",
      shape: { rows: 2, cols: 2 },
      truncated: true
    });
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
    expect(payload.timing).toEqual({ started_at_ms: 0, duration_ms: 1 });
    expect(payload.data.range).toBe("Sheet1!A1:D2000");
    expect(payload.data.count).toBe(2_000);
    expect(payload.data.matching_rows.length).toBeLessThan(payload.data.count);
    expect(payload.data.truncated).toBe(true);
  });

  it("filter_range falls back to toolCall.arguments.range when result omits range", () => {
    const toolCall = {
      id: "call-2c",
      name: "filter_range",
      arguments: { range: "Sheet1!A1:D10" }
    };

    const result = {
      tool: "filter_range",
      ok: true,
      timing: { started_at_ms: 0, duration_ms: 1 },
      data: {
        matching_rows: [2, 4, 6],
        count: 3
      }
    };

    const serialized = serializeToolResultForModel({ toolCall: toolCall as any, result, maxChars: 1_000 });
    const payload = JSON.parse(serialized);
    expect(payload.tool).toBe("filter_range");
    expect(payload.data.range).toBe("Sheet1!A1:D10");
    expect(payload.data.count).toBe(3);
  });

  it("filter_range can fall back to a minimal summary when the tool envelope is too large", () => {
    const toolCall = {
      id: "call-2d",
      name: "filter_range",
      arguments: { range: "Sheet1!A1:D10" }
    };

    const result = {
      tool: "filter_range",
      ok: true,
      timing: { started_at_ms: 0, duration_ms: 1 },
      warnings: ["x".repeat(5_000)],
      data: {
        matching_rows: [2, 4, 6],
        count: 3
      }
    };

    const maxChars = 350;
    const serialized = serializeToolResultForModel({ toolCall: toolCall as any, result, maxChars });
    expect(serialized.length).toBeLessThanOrEqual(maxChars);

    const payload = JSON.parse(serialized);
    expect(payload.tool).toBe("filter_range");
    expect(payload.ok).toBe(true);
    expect(payload.warnings).toBeUndefined();
    expect(payload.timing).toEqual({ started_at_ms: 0, duration_ms: 1 });
    expect(payload.data).toEqual({
      range: "Sheet1!A1:D10",
      count: 3,
      truncated: true
    });
  });

  it("filter_range preserves tool-provided truncated flag even when matching_rows fits preview limit", () => {
    const toolCall = {
      id: "call-2b",
      name: "filter_range",
      arguments: { range: "Sheet1!A1:D50" }
    };

    const result = {
      tool: "filter_range",
      ok: true,
      timing: { started_at_ms: 0, duration_ms: 1 },
      data: {
        range: "Sheet1!A1:D50",
        matching_rows: Array.from({ length: 50 }, (_, i) => i + 1),
        count: 50,
        truncated: true
      }
    };

    const maxChars = 5_000;
    const serialized = serializeToolResultForModel({ toolCall: toolCall as any, result, maxChars });
    expect(serialized.length).toBeLessThanOrEqual(maxChars);

    const payload = JSON.parse(serialized);
    expect(payload.tool).toBe("filter_range");
    expect(payload.ok).toBe(true);
    expect(payload.data.matching_rows).toHaveLength(50);
    expect(payload.data.count).toBe(50);
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
    expect(payload.timing).toEqual({ started_at_ms: 0, duration_ms: 1 });
    expect(payload.data.range).toBe("Sheet1!A1:A5000");
    expect(payload.data.method).toBe("zscore");
    expect(payload.data.total_anomalies).toBe(5_000);
    expect(payload.data.anomalies.length).toBeGreaterThan(0);
    expect(payload.data.anomalies.length).toBeLessThan(1_000);
    expect(payload.data.anomalies[0]).toEqual(expect.objectContaining({ cell: "Sheet1!A1", value: 0, score: 0.9 }));
    expect(payload.data.truncated).toBe(true);
  });

  it("detect_anomalies preserves tool-provided truncated flag even when preview includes all returned anomalies", () => {
    const toolCall = {
      id: "call-3b",
      name: "detect_anomalies",
      arguments: { range: "Sheet1!A1:A10", method: "zscore" }
    };

    const result = {
      tool: "detect_anomalies",
      ok: true,
      timing: { started_at_ms: 0, duration_ms: 1 },
      data: {
        range: "Sheet1!A1:A10",
        method: "zscore",
        anomalies: [{ cell: "Sheet1!A1", value: 1, score: 4 }],
        truncated: true
      }
    };

    const maxChars = 2_000;
    const serialized = serializeToolResultForModel({ toolCall: toolCall as any, result, maxChars });
    expect(serialized.length).toBeLessThanOrEqual(maxChars);

    const payload = JSON.parse(serialized);
    expect(payload.tool).toBe("detect_anomalies");
    expect(payload.data.anomalies).toHaveLength(1);
    expect(payload.data.truncated).toBe(true);
  });

  it("detect_anomalies falls back to a minimal summary when the envelope would exceed maxChars", () => {
    const toolCall = {
      id: "call-3c",
      name: "detect_anomalies",
      arguments: { range: "Sheet1!A1:A10", method: "zscore" }
    };

    const result = {
      tool: "detect_anomalies",
      ok: true,
      timing: { started_at_ms: 0, duration_ms: 1 },
      warnings: ["x".repeat(5_000)],
      data: {
        range: "Sheet1!A1:A10",
        method: "zscore",
        anomalies: [{ cell: "Sheet1!A1", value: 1, score: 4 }],
        truncated: true,
        total_anomalies: 123
      }
    };

    const maxChars = 400;
    const serialized = serializeToolResultForModel({ toolCall: toolCall as any, result, maxChars });
    expect(serialized.length).toBeLessThanOrEqual(maxChars);

    const payload = JSON.parse(serialized);
    expect(payload.tool).toBe("detect_anomalies");
    expect(payload.ok).toBe(true);
    // The fallback path drops huge envelope fields like warnings to fit the budget.
    expect(payload.warnings).toBeUndefined();
    expect(payload.timing).toEqual({ started_at_ms: 0, duration_ms: 1 });
    expect(payload.data).toEqual({
      range: "Sheet1!A1:A10",
      method: "zscore",
      total_anomalies: 123,
      truncated: true
    });
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
