import { afterEach, describe, expect, it, vi } from "vitest";

import { hitTestResizeHandle } from "../selectionHandles";
import type { DrawingTransform } from "../types";

describe("drawings selection handle perf", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("caches rotation trig for repeated hit tests", () => {
    const bounds = { x: 100, y: 200, width: 80, height: 40 };
    const transform: DrawingTransform = { rotationDeg: 45, flipH: false, flipV: false };

    const cosSpy = vi.spyOn(Math, "cos");
    const sinSpy = vi.spyOn(Math, "sin");

    // First call primes the cache.
    expect(hitTestResizeHandle(bounds, 140, 220, transform)).toBeNull();
    expect(cosSpy.mock.calls.length).toBeGreaterThan(0);
    expect(sinSpy.mock.calls.length).toBeGreaterThan(0);

    cosSpy.mockClear();
    sinSpy.mockClear();

    for (let i = 0; i < 200; i += 1) {
      expect(hitTestResizeHandle(bounds, 140, 220, transform)).toBeNull();
    }

    expect(cosSpy).toHaveBeenCalledTimes(0);
    expect(sinSpy).toHaveBeenCalledTimes(0);
  });
});

