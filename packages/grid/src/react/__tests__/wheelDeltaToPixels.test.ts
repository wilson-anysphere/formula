import { describe, expect, it } from "vitest";
import { wheelDeltaToPixels } from "../wheelDeltaToPixels";

describe("wheelDeltaToPixels", () => {
  it("returns pixels unchanged for DOM_DELTA_PIXEL", () => {
    expect(wheelDeltaToPixels(12, 0)).toBe(12);
    expect(wheelDeltaToPixels(-3.5, 0)).toBe(-3.5);
  });

  it("converts DOM_DELTA_LINE using lineHeight", () => {
    expect(wheelDeltaToPixels(3, 1, { lineHeight: 21 })).toBe(63);
    expect(wheelDeltaToPixels(-2, 1, { lineHeight: 10 })).toBe(-20);
  });

  it("converts DOM_DELTA_PAGE using pageSize", () => {
    expect(wheelDeltaToPixels(1, 2, { pageSize: 600 })).toBe(600);
    expect(wheelDeltaToPixels(-0.5, 2, { pageSize: 800 })).toBe(-400);
  });

  it("falls back for invalid lineHeight/pageSize options", () => {
    expect(wheelDeltaToPixels(1, 1, { lineHeight: 0 })).toBe(16);
    // eslint-disable-next-line unicorn/no-null
    expect(wheelDeltaToPixels(1, 1, { lineHeight: Number.NaN })).toBe(16);
    // eslint-disable-next-line unicorn/no-null
    expect(wheelDeltaToPixels(1, 2, { pageSize: Number.NaN })).toBe(800);
    expect(wheelDeltaToPixels(1, 2, { pageSize: 0 })).toBe(0);
  });

  it("treats non-finite values as 0", () => {
    expect(wheelDeltaToPixels(Number.NaN, 0)).toBe(0);
    expect(wheelDeltaToPixels(Number.POSITIVE_INFINITY, 1, { lineHeight: 10 })).toBe(0);
  });
});
