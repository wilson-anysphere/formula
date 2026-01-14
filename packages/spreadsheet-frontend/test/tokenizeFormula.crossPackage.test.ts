import { describe, expect, it } from "vitest";

import { tokenizeFormula as sharedTokenizeFormula } from "../src/formula/tokenizeFormula";
import { tokenizeFormula as desktopTokenizeFormula } from "../../../apps/desktop/src/formula-bar/highlight/tokenizeFormula.js";

describe("tokenizeFormula (cross-package)", () => {
  it("does not tokenize the tail of invalid unquoted sheet names with spaces", () => {
    // Regression: `My Sheet!A1` is invalid without quoting, but we should not end up
    // highlighting/extracting `Sheet!A1` as a sheet-qualified reference.
    const input = "=My Sheet!A1";

    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const desktopRefs = desktopTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["A1"]);
    expect(desktopRefs).toEqual(sharedRefs);
  });

  it("matches between packages for non-BMP Unicode sheet names", () => {
    const input = "=𝔘!A1+𐐷!B2";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const desktopRefs = desktopTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["𝔘!A1", "𐐷!B2"]);
    expect(desktopRefs).toEqual(sharedRefs);
  });

  it("matches between packages for structured table specifiers and selectors", () => {
    const input =
      "=SUM(Table1[#All], Table1[#Headers], Table1[#Data], Table1[#Totals], Table1[[#Headers],[Amount]], Table1[[#Totals],[Amount]])";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const desktopRefs = desktopTokenizeFormula(input)
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
    expect(desktopRefs).toEqual(sharedRefs);
  });

  it("matches between packages for structured refs followed by operators", () => {
    const input = "=Table1[[#All],[Amount]]+1";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const desktopRefs = desktopTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["Table1[[#All],[Amount]]"]);
    expect(desktopRefs).toEqual(sharedRefs);
  });

  it("matches between packages for structured refs with escaped closing brackets in column names", () => {
    const input = "=COUNTA(Table1[[#Headers],[A]]B]])";
    const sharedRefs = sharedTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);
    const desktopRefs = desktopTokenizeFormula(input)
      .filter((t) => t.type === "reference")
      .map((t) => t.text);

    expect(sharedRefs).toEqual(["Table1[[#Headers],[A]]B]]"]);
    expect(desktopRefs).toEqual(sharedRefs);
  });
});
