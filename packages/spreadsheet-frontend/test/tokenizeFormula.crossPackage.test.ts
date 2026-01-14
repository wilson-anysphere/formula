import { describe, expect, it } from "vitest";

import { tokenizeFormula as sharedTokenizeFormula } from "../src/formula/tokenizeFormula";
// Import via the package export path to mirror how downstream consumers (including apps/desktop)
// resolve the shared tokenizer.
import { tokenizeFormula as consumerTokenizeFormula } from "@formula/spreadsheet-frontend/formula/tokenizeFormula";

describe("tokenizeFormula (cross-package)", () => {
  it("does not tokenize the tail of invalid unquoted sheet names with spaces", () => {
    // Regression: `My Sheet!A1` is invalid without quoting, but we should not end up
    // highlighting/extracting `Sheet!A1` as a sheet-qualified reference.
    const input = "=My Sheet!A1";

    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["A1"]);
    expect(consumerRefs).toEqual(sharedRefs);
  });

  it("matches between packages for non-BMP Unicode sheet names", () => {
    const input = "=𝔘!A1+𐐷!B2";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["𝔘!A1", "𐐷!B2"]);
    expect(consumerRefs).toEqual(sharedRefs);
  });

  it("matches between packages for external workbook refs with escaped closing brackets in the workbook name", () => {
    const input = "=SUM([Book]]Name.xlsx]Sheet1!A1, 1)";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["[Book]]Name.xlsx]Sheet1!A1"]);
    expect(consumerRefs).toEqual(sharedRefs);
  });

  it("matches between packages for external workbook refs whose workbook name contains '[' characters", () => {
    // Workbook names may contain `[` without any escaping. Workbook prefixes are not nested, so
    // the bracket span still ends at the first (non-escaped) `]`.
    const input = "=SUM([A1[Name.xlsx]Sheet1!A1, 1)";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["[A1[Name.xlsx]Sheet1!A1"]);
    expect(consumerRefs).toEqual(sharedRefs);
  });

  it("matches between packages for workbook-scoped external defined names (quoted name refs)", () => {
    const input = "='[Book.xlsx]MyName'+1";
    const sharedIdents = sharedTokenizeFormula(input)
      .filter((t) => t.type === "identifier")
      .map((t) => t.text);
    const consumerIdents = consumerTokenizeFormula(input)
      .filter((t) => t.type === "identifier")
      .map((t) => t.text);

    expect(sharedIdents).toEqual(["'[Book.xlsx]MyName'"]);
    expect(consumerIdents).toEqual(sharedIdents);
  });

  it("matches between packages for structured table specifiers and selectors", () => {
    const input =
      "=SUM(Table1[#All], Table1[#Headers], Table1[#Data], Table1[#Totals], Table1[[#Headers],[Amount]], Table1[[#Totals],[Amount]])";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual([
      "Table1[#All]",
      "Table1[#Headers]",
      "Table1[#Data]",
      "Table1[#Totals]",
      "Table1[[#Headers],[Amount]]",
      "Table1[[#Totals],[Amount]]"
    ]);
    expect(consumerRefs).toEqual(sharedRefs);
  });

  it("matches between packages for structured refs followed by operators", () => {
    const input = "=Table1[[#All],[Amount]]+1";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["Table1[[#All],[Amount]]"]);
    expect(consumerRefs).toEqual(sharedRefs);
  });

  it("matches between packages for structured refs with #This Row selectors", () => {
    const input = "=SUM(Table1[[#This Row],[Amount]])";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["Table1[[#This Row],[Amount]]"]);
    expect(consumerRefs).toEqual(sharedRefs);
  });

  it("matches between packages for multi-column structured references", () => {
    const input = "=SUM(Table1[[#All],[Col1],[Col2]])";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["Table1[[#All],[Col1],[Col2]]"]);
    expect(consumerRefs).toEqual(sharedRefs);
  });

  it("matches between packages for structured refs followed by strings containing `]]`", () => {
    const input = '=SUM(Table1[[#All],[Amount]] & "]]", 1)';
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["Table1[[#All],[Amount]]"]);
    expect(consumerRefs).toEqual(sharedRefs);
  });

  it("matches between packages for structured refs with escaped closing brackets in column names", () => {
    const input = "=COUNTA(Table1[[#Headers],[A]]B]])";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["Table1[[#Headers],[A]]B]]"]);
    expect(consumerRefs).toEqual(sharedRefs);
  });

  it("matches between packages for structured refs with escaped `]` followed by operator characters in column names", () => {
    const input = "=COUNTA(Table1[[#Headers],[A]]+B]])";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["Table1[[#Headers],[A]]+B]]"]);
    expect(consumerRefs).toEqual(sharedRefs);
  });

  it("matches between packages for @ and implicit this-row structured references", () => {
    const input = "=SUM(Table1[@Amount], Table1[@[Total Amount]], [@Amount], [@[Total Amount]])";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["Table1[@Amount]", "Table1[@[Total Amount]]", "[@Amount]", "[@[Total Amount]]"]);
    expect(consumerRefs).toEqual(sharedRefs);
  });

  it("matches between packages for implicit structured refs followed by strings containing `]]`", () => {
    const input = '=SUM([[#All],[Amount]] & "]]", 1)';
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["[[#All],[Amount]]"]);
    expect(consumerRefs).toEqual(sharedRefs);
  });

  it("matches between packages for nested implicit this-row refs followed by strings containing `]]`", () => {
    const input = '=SUM([@[Total Amount]] & "]]", 1)';
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const consumerRefs = consumerTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["[@[Total Amount]]"]);
    expect(consumerRefs).toEqual(sharedRefs);
  });
});
