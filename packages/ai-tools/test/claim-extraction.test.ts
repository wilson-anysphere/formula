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
});

