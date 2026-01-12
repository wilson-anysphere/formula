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

  it("can replace overrides in bulk", () => {
    const axis = new VariableSizeAxis(10);
    axis.setOverrides(
      new Map([
        [1, 20],
        [3, 5]
      ])
    );

    expect(axis.getSize(1)).toBe(20);
    expect(axis.getSize(2)).toBe(10);
    expect(axis.getSize(3)).toBe(5);

    expect(axis.positionOf(0)).toBe(0);
    expect(axis.positionOf(1)).toBe(10);
    expect(axis.positionOf(2)).toBe(30);
    expect(axis.positionOf(3)).toBe(40);
    expect(axis.totalSize(4)).toBe(45);

    // Replacing with an empty map should clear overrides.
    axis.setOverrides(new Map());
    expect(axis.getSize(1)).toBe(10);
    expect(axis.totalSize(4)).toBe(40);
  });

  it("updates existing overrides without breaking position math", () => {
    const axis = new VariableSizeAxis(10);
    axis.setOverrides(
      new Map([
        [1, 20],
        [3, 5]
      ])
    );

    // Update an existing override (index 1) and ensure downstream positions adjust.
    axis.setSize(1, 25);
    expect(axis.getSize(1)).toBe(25);
    expect(axis.positionOf(2)).toBe(35); // row0=0:10 + row1=25 => row2 starts at 35
    expect(axis.totalSize(4)).toBe(50); // 10 + 25 + 10 + 5

    // Clearing an override by setting it back to the default should restore defaults.
    axis.setSize(1, 10);
    expect(axis.getSize(1)).toBe(10);
    expect(axis.totalSize(4)).toBe(35); // back to 10 + 10 + 10 + 5
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
