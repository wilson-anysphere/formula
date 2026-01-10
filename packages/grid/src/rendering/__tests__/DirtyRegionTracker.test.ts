import { describe, expect, it } from "vitest";
import { DirtyRegionTracker } from "../DirtyRegionTracker";

describe("DirtyRegionTracker", () => {
  it("merges overlapping regions", () => {
    const tracker = new DirtyRegionTracker();
    tracker.markDirty({ x: 0, y: 0, width: 10, height: 10 });
    tracker.markDirty({ x: 5, y: 5, width: 10, height: 10 });
    expect(tracker.drain()).toEqual([{ x: 0, y: 0, width: 15, height: 15 }]);
  });

  it("merges transitively overlapping regions", () => {
    const tracker = new DirtyRegionTracker();
    tracker.markDirty({ x: 0, y: 0, width: 10, height: 10 });
    tracker.markDirty({ x: 5, y: 0, width: 10, height: 10 });
    tracker.markDirty({ x: 14, y: 0, width: 10, height: 10 });
    expect(tracker.drain()).toEqual([{ x: 0, y: 0, width: 24, height: 10 }]);
  });

  it("keeps separate regions when they do not overlap", () => {
    const tracker = new DirtyRegionTracker();
    tracker.markDirty({ x: 0, y: 0, width: 10, height: 10 });
    tracker.markDirty({ x: 20, y: 0, width: 10, height: 10 });
    const drained = tracker.drain();
    expect(drained).toHaveLength(2);
  });
});

