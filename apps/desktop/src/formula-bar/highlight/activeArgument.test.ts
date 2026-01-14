import { describe, expect, it } from "vitest";

import { getActiveArgumentSpan } from "./activeArgument.js";

describe("getActiveArgumentSpan", () => {
  it("returns null when the cursor is not inside a function call", () => {
    expect(getActiveArgumentSpan("=A1+1", 2)).toBeNull();
    expect(getActiveArgumentSpan("SUM(A1)", 2)).toBeNull();
  });

  it("returns the innermost function + arg span (nested calls)", () => {
    const formula = '=IF(SUM(A1:A2) > 3, "yes", "no")';
    const insideSum = formula.indexOf("A1") + 1;
    expect(getActiveArgumentSpan(formula, insideSum)).toEqual({
      fnName: "SUM",
      argIndex: 0,
      argText: "A1:A2",
      span: { start: formula.indexOf("A1:A2"), end: formula.indexOf("A1:A2") + "A1:A2".length },
    });

    const insideIfSecondArg = formula.indexOf('"yes"') + 1;
    expect(getActiveArgumentSpan(formula, insideIfSecondArg)).toEqual({
      fnName: "IF",
      argIndex: 1,
      argText: '"yes"',
      span: { start: formula.indexOf('"yes"'), end: formula.indexOf('"yes"') + '"yes"'.length },
    });
  });

  it("treats whitespace between function name and '(' as part of the call (Excel-style)", () => {
    const formula = "=SUM ( A1 , B1 )";
    const insideA1 = formula.indexOf("A1") + 1;
    expect(getActiveArgumentSpan(formula, insideA1)).toMatchObject({
      fnName: "SUM",
      argIndex: 0,
      argText: "A1",
    });

    const insideB1 = formula.indexOf("B1") + 1;
    expect(getActiveArgumentSpan(formula, insideB1)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "B1",
    });
  });

  it("ignores commas inside string literals", () => {
    const formula = '=CONCAT("a,b", "c")';
    const insideFirstString = formula.indexOf("a,b") + 1;
    expect(getActiveArgumentSpan(formula, insideFirstString)).toMatchObject({
      fnName: "CONCAT",
      argIndex: 0,
      argText: '"a,b"',
    });
  });

  it("ignores commas inside nested parentheses (union operator)", () => {
    const formula = "=SUM((A1,B1), C1)";
    const insideC1 = formula.indexOf("C1") + 1;
    expect(getActiveArgumentSpan(formula, insideC1)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "C1",
    });
  });

  it("ignores commas inside square brackets (structured / external refs)", () => {
    const formula = "=SUM(Table1[[#All],[Col1],[Col2]], 1)";
    const insideSecondArg = formula.lastIndexOf("1") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "1",
    });
  });

  it("ignores commas inside external workbook prefixes even when the workbook name contains '[' characters", () => {
    // Workbook name contains a literal `[`, which does not introduce nesting in workbook prefixes.
    // The bracket skipper must still find the closing `]` so the comma after the reference is
    // treated as the argument separator.
    const formula = "=SUM([A1[Name.xlsx]Sheet1!A1, 1)";
    const insideSecondArg = formula.lastIndexOf("1") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "1",
    });
  });

  it("ignores commas inside external workbook name refs even when the workbook name contains '[' characters", () => {
    // Workbook-scoped external defined names do not use `!`, but we still need to detect the end
    // of the workbook prefix (which is non-nesting) so the comma after the name is treated as the
    // argument separator.
    const formula = "=SUM([A1[Name.xlsx]MyName, 1)";
    const insideSecondArg = formula.lastIndexOf("1") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "1",
    });
  });

  it("ignores commas inside structured refs with escaped closing brackets", () => {
    // Regression: `]]` escapes inside column names should not cause us to treat internal
    // structured-ref commas as argument separators.
    const formula = "=SUM(Table1[[#Headers],[A]]B],[Col2]], 1)";
    const insideSecondArg = formula.lastIndexOf("1") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "1",
    });
  });

  it("treats parentheses inside structured references as plain text", () => {
    const formula = "=SUM(Table1[Amount)], 1)";
    const insideSecondArg = formula.lastIndexOf("1") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "1",
    });
  });

  it("does not treat function-like text inside structured references as nested calls", () => {
    const formula = "=SUM(Table1[Amount(USD)], 1)";
    const insideSecondArg = formula.lastIndexOf("1") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "1",
    });
  });

  it("treats escaped closing brackets inside structured references as plain text", () => {
    const formula = "=SUM(Table1[Total]],USD], 1)";
    const insideRef = formula.indexOf("USD") + 1;
    expect(getActiveArgumentSpan(formula, insideRef)).toMatchObject({
      fnName: "SUM",
      argIndex: 0,
      argText: "Table1[Total]],USD]",
    });

    const insideSecondArg = formula.lastIndexOf("1") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "1",
    });
  });

  it("ignores commas and parentheses inside quoted sheet names", () => {
    const formula = "=SUM('Budget,2025)'!A1, 1)";
    const insideSecondArg = formula.lastIndexOf("1") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "1",
    });

    const insideSheetRef = formula.indexOf("A1") + 1;
    expect(getActiveArgumentSpan(formula, insideSheetRef)).toMatchObject({
      fnName: "SUM",
      argIndex: 0,
      argText: "'Budget,2025)'!A1",
    });
  });

  it("ignores semicolons inside quoted sheet names (locale arg separators)", () => {
    const formula = "=SUM('Budget;2025'!A1; 1)";
    const insideSecondArg = formula.lastIndexOf("1") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "1",
    });
  });

  it("does not let unbalanced parentheses inside array literals affect argument indexing", () => {
    const formula = "=SUM({1,(2,3}, 4)";
    const insideSecondArg = formula.indexOf("4") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "4",
    });
  });

  it("ignores commas inside curly braces (array literals)", () => {
    const formula = "=SUM({1,2,3}, 4)";
    const insideSecondArg = formula.indexOf("4") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "4",
    });
  });

  it("supports semicolons as argument separators", () => {
    const formula = "=IF(A1; B1; C1)";
    const insideSecondArg = formula.indexOf("B1") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg)).toMatchObject({
      fnName: "IF",
      argIndex: 1,
      argText: "B1",
    });
  });

  it("supports non-ASCII letters in function names (e.g. de-DE ZÄHLENWENN)", () => {
    const formula = '=ZÄHLENWENN(A1:A3; ">0")';
    const insideRange = formula.indexOf("A1") + 1;
    expect(getActiveArgumentSpan(formula, insideRange, { argSeparators: ";" })).toMatchObject({
      fnName: "ZÄHLENWENN",
      argIndex: 0,
      argText: "A1:A3",
    });

    const insideCriteria = formula.indexOf('">0"') + 2;
    expect(getActiveArgumentSpan(formula, insideCriteria, { argSeparators: ";" })).toMatchObject({
      fnName: "ZÄHLENWENN",
      argIndex: 1,
      argText: '">0"',
    });
  });

  it("ignores semicolons inside curly braces (array literals)", () => {
    const formula = "=SUM({1;2;3}; 4)";
    const insideSecondArg = formula.indexOf("4") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "4",
    });
  });

  it("returns an empty argText span when the argument is currently empty", () => {
    const formula = "=SUM(A1, )";
    const cursorInEmptyArg = formula.indexOf(")") - 1;
    expect(getActiveArgumentSpan(formula, cursorInEmptyArg)).toMatchObject({
      fnName: "SUM",
      argIndex: 1,
      argText: "",
    });
  });

  it("can be configured to only treat ';' as an argument separator (so decimal commas don't split)", () => {
    const formula = "=ROUND(1,2; 0)";
    const insideFirstArg = formula.indexOf("1,2") + 1;
    expect(getActiveArgumentSpan(formula, insideFirstArg, { argSeparators: ";" })).toMatchObject({
      fnName: "ROUND",
      argIndex: 0,
      argText: "1,2",
    });

    const insideSecondArg = formula.indexOf("0") + 1;
    expect(getActiveArgumentSpan(formula, insideSecondArg, { argSeparators: ";" })).toMatchObject({
      fnName: "ROUND",
      argIndex: 1,
      argText: "0",
    });
  });

  it("treats escaped closing brackets inside structured refs as plain text (does not end bracket context)", () => {
    // Regression: column names may contain escaped `]` (written as `]]`). When the escaped
    // `]` is followed by function-like text, we must not treat it as a nested call.
    const structuredRef = "Table1[[#All],[A]]SUM(1,2)]]";
    const formula = `=IF(${structuredRef}, 1, 2)`;
    const cursorInsideNestedFunctionText = formula.indexOf("SUM(1,2)") + "SUM(".length; // on the `1`

    expect(getActiveArgumentSpan(formula, cursorInsideNestedFunctionText)).toEqual({
      fnName: "IF",
      argIndex: 0,
      argText: structuredRef,
      span: { start: formula.indexOf(structuredRef), end: formula.indexOf(structuredRef) + structuredRef.length },
    });
  });
});
