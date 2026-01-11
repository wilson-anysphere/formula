import { describe, expect, it } from "vitest";

import { explainFormulaError } from "./errors.js";

describe("explainFormulaError", () => {
  it("returns explanations for AI-related error codes", () => {
    expect(explainFormulaError("#GETTING_DATA")?.title).toBe("Loading");
    expect(explainFormulaError("#DLP!")?.title).toBe("Blocked by data loss prevention");
    expect(explainFormulaError("#AI!")?.title).toBe("AI error");
  });

  it("returns explanations for newer Excel error codes", () => {
    expect(explainFormulaError("#CALC!")?.title).toBe("Calculation error");
    expect(explainFormulaError("#CONNECT!")?.title).toBe("Connection error");
    expect(explainFormulaError("#FIELD!")?.title).toBe("Invalid field");
    expect(explainFormulaError("#BLOCKED!")?.title).toBe("Blocked");
    expect(explainFormulaError("#UNKNOWN!")?.title).toBe("Unknown error");
    expect(explainFormulaError("#NULL!")?.title).toBe("Null intersection");
  });
});
