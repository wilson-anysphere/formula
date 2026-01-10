import { describe, expect, it } from "vitest";
import { VariableSizeAxis } from "../VariableSizeAxis";

describe("VariableSizeAxis", () => {
  it("computes positions with default sizing", () => {
    const axis = new VariableSizeAxis(10);
    expect(axis.positionOf(0)).toBe(0);
    expect(axis.positionOf(1)).toBe(10);
    expect(axis.positionOf(5)).toBe(50);
    expect(axis.totalSize(3)).toBe(30);
  });

  it("supports overrides", () => {
    const axis = new VariableSizeAxis(10);
    axis.setSize(1, 20);
    expect(axis.positionOf(0)).toBe(0);
    expect(axis.positionOf(1)).toBe(10);
    expect(axis.positionOf(2)).toBe(30);
    expect(axis.totalSize(3)).toBe(40);
    expect(axis.getSize(1)).toBe(20);
  });

  it("finds indices at positions (binary search)", () => {
    const axis = new VariableSizeAxis(10);
    axis.setSize(1, 20);

    expect(axis.indexAt(0)).toBe(0);
    expect(axis.indexAt(9)).toBe(0);
    expect(axis.indexAt(10)).toBe(1);
    expect(axis.indexAt(15)).toBe(1);
    expect(axis.indexAt(29)).toBe(1);
    expect(axis.indexAt(30)).toBe(2);
  });

  it("computes visible range", () => {
    const axis = new VariableSizeAxis(10);
    const range = axis.visibleRange(0, 25, { min: 0, maxExclusive: 100 });
    expect(range).toEqual({ start: 0, end: 3, offset: 0 });
  });
});

