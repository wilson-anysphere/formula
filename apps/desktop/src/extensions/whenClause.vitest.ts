import { describe, expect, it } from "vitest";

import { evaluateWhenClause, parseWhenClause } from "./whenClause.js";

describe("when-clause parsing + eval", () => {
  const lookup = (values: Record<string, any>) => (key: string) => values[key];

  it("treats missing/empty clauses as satisfied", () => {
    expect(evaluateWhenClause(null, lookup({}))).toBe(true);
    expect(evaluateWhenClause(undefined, lookup({}))).toBe(true);
    expect(evaluateWhenClause("   ", lookup({}))).toBe(true);
  });

  it("supports identifiers + boolean logic", () => {
    expect(evaluateWhenClause("cellHasValue", lookup({ cellHasValue: true }))).toBe(true);
    expect(evaluateWhenClause("cellHasValue", lookup({ cellHasValue: false }))).toBe(false);
    expect(evaluateWhenClause("cellHasValue && hasSelection", lookup({ cellHasValue: true, hasSelection: false }))).toBe(
      false,
    );
    expect(evaluateWhenClause("cellHasValue || hasSelection", lookup({ cellHasValue: false, hasSelection: true }))).toBe(
      true,
    );
    expect(evaluateWhenClause("!cellHasValue || hasSelection", lookup({ cellHasValue: true, hasSelection: false }))).toBe(
      false,
    );
  });

  it("honors precedence + parentheses", () => {
    const ctx = lookup({ a: false, b: true, c: false });
    // && binds tighter than ||
    expect(evaluateWhenClause("a || b && c", ctx)).toBe(false);
    expect(evaluateWhenClause("(a || b) && c", ctx)).toBe(false);
    expect(evaluateWhenClause("a || (b && !c)", ctx)).toBe(true);
  });

  it("supports == / != literals", () => {
    const ctx = lookup({ sheetName: "Sheet1", count: 2 });
    expect(evaluateWhenClause("sheetName == 'Sheet1'", ctx)).toBe(true);
    expect(evaluateWhenClause("sheetName != 'Sheet2'", ctx)).toBe(true);
    expect(evaluateWhenClause("count == 2", ctx)).toBe(true);
    expect(evaluateWhenClause("count != 3", ctx)).toBe(true);
  });

  it("fails closed on invalid expressions", () => {
    expect(evaluateWhenClause("cellHasValue &&", lookup({ cellHasValue: true }))).toBe(false);
    expect(() => parseWhenClause("cellHasValue &&")).toThrow();
  });
});

