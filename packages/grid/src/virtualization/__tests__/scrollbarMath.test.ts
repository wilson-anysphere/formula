import { describe, expect, it } from "vitest";
import { computeScrollbarThumb } from "../scrollbarMath";

describe("computeScrollbarThumb", () => {
  it("supports an optional `out` param to reuse objects", () => {
    const out = { size: -1, offset: -1 };

    const first = computeScrollbarThumb({
      scrollPos: 0,
      viewportSize: 100,
      contentSize: 1000,
      trackSize: 50,
      out
    });

    expect(first).toBe(out);
    expect(first.size).toBeGreaterThan(0);
    expect(first.offset).toBe(0);

    const second = computeScrollbarThumb({
      scrollPos: 450,
      viewportSize: 100,
      contentSize: 1000,
      trackSize: 50,
      out
    });

    expect(second).toBe(out);
    expect(second.size).toBe(first.size);
    expect(second.offset).toBeGreaterThan(0);
  });

  it("returns the full track size when there is no scrollable overflow", () => {
    expect(
      computeScrollbarThumb({
        scrollPos: 0,
        viewportSize: 100,
        contentSize: 100,
        trackSize: 50
      })
    ).toEqual({ size: 50, offset: 0 });
  });
});

