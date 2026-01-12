import { describe, expect, it } from "vitest";
import { computeAnchoredScrollAfterZoom, computeZoomFromPinchDistance } from "../zoomMath";

describe("zoomMath", () => {
  it("clamps pinch zoom factor", () => {
    expect(
      computeZoomFromPinchDistance({
        startZoom: 1,
        startDistance: 100,
        currentDistance: 500,
        minZoom: 0.5,
        maxZoom: 3
      })
    ).toBe(3);

    expect(
      computeZoomFromPinchDistance({
        startZoom: 1,
        startDistance: 100,
        currentDistance: 10,
        minZoom: 0.5,
        maxZoom: 3
      })
    ).toBe(0.5);
  });

  it("adjusts scroll to keep anchor stable while zooming", () => {
    const startScroll = { x: 100, y: 200 };
    const startAnchor = { x: 250, y: 250 };

    const next = computeAnchoredScrollAfterZoom({
      viewport: { width: 800, height: 600, frozenWidth: 0, frozenHeight: 0 },
      startZoom: 1,
      nextZoom: 2,
      startScroll,
      startAnchor,
      nextAnchor: startAnchor
    });

    expect(next.x).toBe((startScroll.x + startAnchor.x) * 2 - startAnchor.x);
    expect(next.y).toBe((startScroll.y + startAnchor.y) * 2 - startAnchor.y);
  });

  it("adds pan when the pinch midpoint translates", () => {
    const startScroll = { x: 100, y: 200 };
    const startAnchor = { x: 250, y: 250 };
    const nextAnchor = { x: 300, y: 240 };

    const next = computeAnchoredScrollAfterZoom({
      viewport: { width: 800, height: 600, frozenWidth: 0, frozenHeight: 0 },
      startZoom: 1,
      nextZoom: 2,
      startScroll,
      startAnchor,
      nextAnchor
    });

    expect(next.x).toBe((startScroll.x + startAnchor.x) * 2 - nextAnchor.x);
    expect(next.y).toBe((startScroll.y + startAnchor.y) * 2 - nextAnchor.y);
  });
});

