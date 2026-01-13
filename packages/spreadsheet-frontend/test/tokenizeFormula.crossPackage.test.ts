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
});
