import { describe, expect, it } from "vitest";

import { extractVerifiableClaims } from "../src/llm/verification.js";

describe("extractVerifiableClaims", () => {
  it("extracts a range statistic claim with an explicit reference", () => {
    const claims = extractVerifiableClaims({
      assistantText: "The average of Sheet1!A1:A3 is 2.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mean",
        reference: "Sheet1!A1:A3",
        expected: 2,
        source: "average of Sheet1!A1:A3 is 2"
      }
    ]);
  });

  it("extracts a range statistic claim when the sheet name contains dots", () => {
    const claims = extractVerifiableClaims({
      assistantText: "The average of Sheet.Name!A1:A3 is 2.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mean",
        reference: "Sheet.Name!A1:A3",
        expected: 2,
        source: "average of Sheet.Name!A1:A3 is 2"
      }
    ]);
  });

  it("extracts claims for hyphenated, unquoted sheet names", () => {
    const claims = extractVerifiableClaims({
      assistantText: "The average of Q1-2025!a1:a3 is 2.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mean",
        reference: "Q1-2025!A1:A3",
        expected: 2,
        source: "average of Q1-2025!a1:a3 is 2"
      }
    ]);
  });

  it("attaches the user question reference when the assistant omits the range", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Average is 2.",
      userText: "What is the average of A1:A3?"
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mean",
        reference: "A1:A3",
        expected: 2,
        source: "Average is 2"
      }
    ]);
  });

  it("infers dot-containing sheet references from the user question when the assistant omits the range", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Average is 2.",
      userText: "What is the average of Sheet.Name!A1:A3?"
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mean",
        reference: "Sheet.Name!A1:A3",
        expected: 2,
        source: "Average is 2"
      }
    ]);
  });

  it("attaches the pivot source_range when the assistant omits the range", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Average is 2.",
      userText: "",
      toolCalls: [
        {
          name: "create_pivot_table",
          parameters: {
            source_range: "Sheet1!A1:A3",
            destination: "Sheet1!C1",
            rows: ["X"],
            values: [{ field: "Y", aggregation: "sum" }]
          }
        }
      ]
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mean",
        reference: "Sheet1!A1:A3",
        expected: 2,
        source: "Average is 2"
      }
    ]);
  });

  it("attaches the pivot source_range when tool calls use `arguments` (LLM ToolCall shape)", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Average is 2.",
      userText: "",
      toolCalls: [
        {
          name: "create_pivot_table",
          arguments: {
            source_range: "Sheet1!A1:A3",
            destination: "Sheet1!C1",
            rows: ["X"],
            values: [{ field: "Y", aggregation: "sum" }]
          }
        }
      ]
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mean",
        reference: "Sheet1!A1:A3",
        expected: 2,
        source: "Average is 2"
      }
    ]);
  });

  it("attaches the pivot source_range when tool calls use JSON string arguments", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Average is 2.",
      userText: "",
      toolCalls: [
        {
          name: "create_pivot_table",
          arguments: JSON.stringify({
            source_range: "Sheet1!A1:A3",
            destination: "Sheet1!C1",
            rows: ["X"],
            values: [{ field: "Y", aggregation: "sum" }]
          })
        }
      ]
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mean",
        reference: "Sheet1!A1:A3",
        expected: 2,
        source: "Average is 2"
      }
    ]);
  });

  it("attaches the pivot sourceRange (camelCase) when the assistant omits the range", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Average is 2.",
      userText: "",
      toolCalls: [
        {
          name: "create_pivot_table",
          parameters: {
            sourceRange: "Sheet1!A1:A3",
            destination: "Sheet1!C1",
            rows: ["X"],
            values: [{ field: "Y", aggregation: "sum" }]
          }
        }
      ]
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mean",
        reference: "Sheet1!A1:A3",
        expected: 2,
        source: "Average is 2"
      }
    ]);
  });

  it("attaches the chart data_range when the assistant omits the range", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Average is 2.",
      userText: "",
      toolCalls: [
        {
          name: "create_chart",
          parameters: {
            chart_type: "bar",
            data_range: "Sheet1!B1:B3",
            position: "Sheet1!D1"
          }
        }
      ]
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mean",
        reference: "Sheet1!B1:B3",
        expected: 2,
        source: "Average is 2"
      }
    ]);
  });

  it("attaches the chart data_range when tool calls use `arguments` (LLM ToolCall shape)", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Average is 2.",
      userText: "",
      toolCalls: [
        {
          name: "create_chart",
          arguments: {
            chart_type: "bar",
            data_range: "Sheet1!B1:B3",
            position: "Sheet1!D1"
          }
        }
      ]
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mean",
        reference: "Sheet1!B1:B3",
        expected: 2,
        source: "Average is 2"
      }
    ]);
  });

  it("attaches the chart data_range when tool calls use JSON string arguments", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Average is 2.",
      userText: "",
      toolCalls: [
        {
          name: "create_chart",
          arguments: JSON.stringify({
            chart_type: "bar",
            data_range: "Sheet1!B1:B3",
            position: "Sheet1!D1"
          })
        }
      ]
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mean",
        reference: "Sheet1!B1:B3",
        expected: 2,
        source: "Average is 2"
      }
    ]);
  });

  it("attaches the chart dataRange (camelCase) when the assistant omits the range", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Average is 2.",
      userText: "",
      toolCalls: [
        {
          name: "create_chart",
          parameters: {
            chart_type: "bar",
            dataRange: "Sheet1!B1:B3",
            position: "Sheet1!D1"
          }
        }
      ]
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mean",
        reference: "Sheet1!B1:B3",
        expected: 2,
        source: "Average is 2"
      }
    ]);
  });

  it("extracts sum/total claims (comma-separated numbers)", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Total for range Sheet1!B1:B2 = 1,200.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "sum",
        reference: "Sheet1!B1:B2",
        expected: 1200,
        source: "Total for range Sheet1!B1:B2 = 1,200"
      }
    ]);
  });

  it("supports \"of the range\" phrasing", () => {
    const claims = extractVerifiableClaims({
      assistantText: "The median of the range Sheet1!A1:A2 is 0.5.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "median",
        reference: "Sheet1!A1:A2",
        expected: 0.5,
        source: "median of the range Sheet1!A1:A2 is 0.5"
      }
    ]);
  });

  it("supports \"in range\" phrasing", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Std dev in range Sheet1!B2:B100 = 12.3.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "stdev",
        reference: "Sheet1!B2:B100",
        expected: 12.3,
        source: "Std dev in range Sheet1!B2:B100 = 12.3"
      }
    ]);
  });

  it("extracts implicit count claims (there are ... in range)", () => {
    const claims = extractVerifiableClaims({
      assistantText: "There are 99 values in the range Sheet1!A1:A10.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "count",
        reference: "Sheet1!A1:A10",
        expected: 99,
        source: "There are 99 values in the range Sheet1!A1:A10"
      }
    ]);
  });

  it("extracts implicit count claims (range has ... values)", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Sheet1!A1:A10 has 99 values.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "count",
        reference: "Sheet1!A1:A10",
        expected: 99,
        source: "Sheet1!A1:A10 has 99 values"
      }
    ]);
  });

  it("extracts implicit count claims (number of values in range ...)", () => {
    const claims = extractVerifiableClaims({
      assistantText: "The number of values in Sheet1!A1:A10 is 99.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "count",
        reference: "Sheet1!A1:A10",
        expected: 99,
        source: "The number of values in Sheet1!A1:A10 is 99"
      }
    ]);
  });

  it("parses parenthesized negative numbers", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Total for range Sheet1!B1:B2 = (1,200).",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "sum",
        reference: "Sheet1!B1:B2",
        expected: -1200,
        source: "Total for range Sheet1!B1:B2 = (1,200)"
      }
    ]);
  });

  it("parses currency-prefixed numbers", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Total for range Sheet1!B1:B2 = $1,200.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "sum",
        reference: "Sheet1!B1:B2",
        expected: 1200,
        source: "Total for range Sheet1!B1:B2 = $1,200"
      }
    ]);
  });

  it("extracts cell value claims", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Sheet1!C5 is 10.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "cell_value",
        reference: "Sheet1!C5",
        expected: 10,
        source: "Sheet1!C5 is 10"
      }
    ]);
  });

  it("extracts stdev claims (std dev phrasing)", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Std dev for Sheet1!B2:B100 = 12.3.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "stdev",
        reference: "Sheet1!B2:B100",
        expected: 12.3,
        source: "Std dev for Sheet1!B2:B100 = 12.3"
      }
    ]);
  });

  it("extracts range-stat function call claims (formula-style)", () => {
    const claims = extractVerifiableClaims({
      assistantText: "SUM(Sheet1!A1:A3) = 6.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "sum",
        reference: "Sheet1!A1:A3",
        expected: 6,
        source: "SUM(Sheet1!A1:A3) = 6"
      }
    ]);
  });

  it("extracts mode claims", () => {
    const claims = extractVerifiableClaims({
      assistantText: "The mode of Sheet1!A1:A5 is 2.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "mode",
        reference: "Sheet1!A1:A5",
        expected: 2,
        source: "mode of Sheet1!A1:A5 is 2"
      }
    ]);
  });

  it("extracts correlation claims", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Correlation of Sheet1!A1:B3 is 1.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "correlation",
        reference: "Sheet1!A1:B3",
        expected: 1,
        source: "Correlation of Sheet1!A1:B3 is 1"
      }
    ]);
  });

  it("extracts stdev claims (std. dev punctuation)", () => {
    const claims = extractVerifiableClaims({
      assistantText: "Std. dev for Sheet1!B2:B100 = 12.3.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "stdev",
        reference: "Sheet1!B2:B100",
        expected: 12.3,
        source: "Std. dev for Sheet1!B2:B100 = 12.3"
      }
    ]);
  });

  it("parses leading-decimal numbers", () => {
    const claims = extractVerifiableClaims({
      assistantText: "The median of Sheet1!A1:A2 is .5.",
      userText: ""
    });

    expect(claims).toEqual([
      {
        kind: "range_stat",
        measure: "median",
        reference: "Sheet1!A1:A2",
        expected: 0.5,
        source: "median of Sheet1!A1:A2 is .5"
      }
    ]);
  });
});
