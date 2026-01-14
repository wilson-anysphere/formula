import { describe, expect, it } from "vitest";
import { alignScrollToDevicePixels } from "../alignScrollToDevicePixels";

describe("alignScrollToDevicePixels", () => {
  it("rounds to device-pixel boundaries", () => {
    const max = { maxScrollX: 100, maxScrollY: 100 };
    expect(alignScrollToDevicePixels({ x: 1.2, y: 1.2 }, max, 2)).toEqual({ x: 1, y: 1 });
    expect(alignScrollToDevicePixels({ x: 1.3, y: 1.3 }, max, 2)).toEqual({ x: 1.5, y: 1.5 });
  });

  it("clamps to an aligned max scroll", () => {
    const max = { maxScrollX: 1.3, maxScrollY: 0.9 };
    // At dpr=2, step=0.5, so maxAlignedX=1.0 and maxAlignedY=0.5.
    expect(alignScrollToDevicePixels({ x: 2, y: 2 }, max, 2)).toEqual({ x: 1, y: 0.5 });
    expect(alignScrollToDevicePixels({ x: -5, y: -5 }, max, 2)).toEqual({ x: 0, y: 0 });
  });
});

