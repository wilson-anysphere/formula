import { describe, expect, it } from "vitest";
import { colToName, fromA1, toA1 } from "../src/index.js";

describe("A1 conversions", () => {
  it("roundtrips common addresses", () => {
    const samples = [
      { row0: 0, col0: 0, a1: "A1" },
      { row0: 0, col0: 25, a1: "Z1" },
      { row0: 0, col0: 26, a1: "AA1" },
      { row0: 9, col0: 0, a1: "A10" },
      { row0: 123, col0: 702, a1: "AAA124" }
    ];

    for (const sample of samples) {
      expect(toA1(sample.row0, sample.col0)).toBe(sample.a1);
      expect(fromA1(sample.a1)).toEqual({ row0: sample.row0, col0: sample.col0 });
    }
  });

  it("parses optional sheet prefixes", () => {
    expect(fromA1("Sheet1!B2")).toEqual({ row0: 1, col0: 1 });
    expect(fromA1("'Sheet Name'!$C$3")).toEqual({ row0: 2, col0: 2 });
  });

  it("converts column indices to labels", () => {
    expect(colToName(0)).toBe("A");
    expect(colToName(25)).toBe("Z");
    expect(colToName(26)).toBe("AA");
    expect(colToName(701)).toBe("ZZ");
    expect(colToName(702)).toBe("AAA");
  });
});
