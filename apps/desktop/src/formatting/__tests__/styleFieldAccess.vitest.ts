import { describe, expect, it } from "vitest";

import { getStyleNumberFormat } from "../styleFieldAccess.js";

describe("styleFieldAccess.getStyleNumberFormat", () => {
  it("returns null for empty/General formats", () => {
    expect(getStyleNumberFormat({})).toBeNull();
    expect(getStyleNumberFormat({ numberFormat: "" })).toBeNull();
    expect(getStyleNumberFormat({ numberFormat: "   " })).toBeNull();
    expect(getStyleNumberFormat({ numberFormat: "General" })).toBeNull();
    expect(getStyleNumberFormat({ numberFormat: " general " })).toBeNull();
    expect(getStyleNumberFormat({ number_format: "GENERAL" })).toBeNull();
  });

  it("returns the raw string for non-General formats", () => {
    expect(getStyleNumberFormat({ numberFormat: "0.00" })).toBe("0.00");
    // Preserve whitespace when it is not the special "General" format.
    expect(getStyleNumberFormat({ numberFormat: "0.00 " })).toBe("0.00 ");
    expect(getStyleNumberFormat({ number_format: "0%" })).toBe("0%");
  });
});

