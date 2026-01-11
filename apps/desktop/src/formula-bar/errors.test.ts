import { describe, expect, it } from "vitest";

import { explainFormulaError } from "./errors.js";

describe("explainFormulaError", () => {
  it("returns explanations for AI-related error codes", () => {
    expect(explainFormulaError("#GETTING_DATA")?.title).toBe("Loading");
    expect(explainFormulaError("#DLP!")?.title).toBe("Blocked by data loss prevention");
    expect(explainFormulaError("#AI!")?.title).toBe("AI error");
  });
});

