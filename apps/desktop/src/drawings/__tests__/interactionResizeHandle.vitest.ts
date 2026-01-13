import { describe, expect, it } from "vitest";

import { hitTestResizeHandle, RESIZE_HANDLE_SIZE_PX } from "../selectionHandles";

describe("drawing resize handle hit testing", () => {
  it("matches the rendered handle geometry", () => {
    const bounds = { x: 10, y: 10, width: 20, height: 20 };
    const half = RESIZE_HANDLE_SIZE_PX / 2;

    // Inside the rendered NW handle square.
    expect(hitTestResizeHandle(bounds, bounds.x + half - 1, bounds.y + half - 1)).toBe("nw");

    // Just outside the rendered handle square (this used to hit when the hit box was larger).
    expect(hitTestResizeHandle(bounds, bounds.x + half + 1, bounds.y + half + 1)).toBeNull();
  });
});
