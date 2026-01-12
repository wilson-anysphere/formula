import { describe, expect, it } from "vitest";
import { shiftA1References } from "../src/index.js";

describe("shiftA1References", () => {
  it("shifts simple references", () => {
    expect(shiftA1References("=A1", 1, 0)).toBe("=A2");
    expect(shiftA1References("=A1", 0, 1)).toBe("=B1");
    expect(shiftA1References("=A1", 2, 3)).toBe("=D3");
  });

  it("respects $-absolute columns/rows", () => {
    expect(shiftA1References("=$A1", 1, 1)).toBe("=$A2");
    expect(shiftA1References("=A$1", 1, 1)).toBe("=B$1");
    expect(shiftA1References("=$A$1", 10, 10)).toBe("=$A$1");
  });

  it("shifts both endpoints of simple ranges", () => {
    expect(shiftA1References("=SUM(A1:B2)", 1, 1)).toBe("=SUM(B2:C3)");
    expect(shiftA1References("=SUM($A1:B$2)", 1, 1)).toBe("=SUM($A2:C$2)");
  });

  it("shifts whole-row and whole-column references", () => {
    expect(shiftA1References("=SUM(A:A)", 0, 1)).toBe("=SUM(B:B)");
    expect(shiftA1References("=SUM(1:1)", 1, 0)).toBe("=SUM(2:2)");
    expect(shiftA1References("=SUM(Sheet1!A:B)", 0, 2)).toBe("=SUM(Sheet1!C:D)");
    expect(shiftA1References("=SUM('Sheet Name'!$1:2)", 3, 0)).toBe("=SUM('Sheet Name'!$1:5)");
  });

  it("drops the spill-range postfix when shifting creates a #REF!", () => {
    expect(shiftA1References("=A1#", 0, -1)).toBe("=#REF!");
  });

  it("drops sheet prefixes when shifting creates a #REF! (engine grammar does not accept Sheet1!#REF!)", () => {
    expect(shiftA1References("=Sheet1!A1", 0, -1)).toBe("=#REF!");
    expect(shiftA1References("='Sheet Name'!A1", 0, -1)).toBe("=#REF!");
  });

  it("shifts sheet-qualified references", () => {
    expect(shiftA1References("=Sheet1!A1+1", 2, 0)).toBe("=Sheet1!A3+1");
    expect(shiftA1References("='Sheet Name'!$A$1", 3, 2)).toBe("='Sheet Name'!$A$1");
    expect(shiftA1References("='Sheet'' Name'!A1", 1, 0)).toBe("='Sheet'' Name'!A2");
  });

  it("does not shift inside double-quoted strings", () => {
    expect(shiftA1References('="A1"&A1', 1, 0)).toBe('="A1"&A2');
  });

  it("avoids shifting function names that look like A1 refs", () => {
    expect(shiftA1References("=LOG10(A1)", 1, 0)).toBe("=LOG10(A2)");
  });
});
