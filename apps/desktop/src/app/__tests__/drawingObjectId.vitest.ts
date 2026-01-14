import { describe, expect, it } from "vitest";

import { insertImageFromFile } from "../../drawings/insertImage";
import { createDrawingObjectId, type Anchor, type ImageStore } from "../../drawings/types";

describe("DrawingObject ids", () => {
  it("createDrawingObjectId generates unique safe integers", () => {
    const ids = new Set<number>();
    const count = 10_000;

    for (let i = 0; i < count; i += 1) {
      const id = createDrawingObjectId();
      expect(Number.isSafeInteger(id)).toBe(true);
      expect(id).toBeGreaterThan(0);
      ids.add(id);
    }

    expect(ids.size).toBe(count);
  });

  it("insertImageFromFile ignores caller-provided ids and generates a safe id", async () => {
    const file = new File([new Uint8Array([1, 2, 3])], "test.png", { type: "image/png" });

    const store = new Map<string, any>();
    const images: ImageStore = {
      get: (id) => store.get(id),
      set: (entry) => store.set(entry.id, entry),
      delete: (id) => store.delete(id),
      clear: () => store.clear(),
    };

    const anchor: Anchor = {
      type: "absolute",
      pos: { xEmu: 0, yEmu: 0 },
      size: { cx: 10, cy: 10 },
    };

    // `2**53` is not a safe integer; if insertion naively used a caller-provided numeric counter,
    // it could persist unsafe/non-collaboration-safe ids.
    const unsafeCounterId = 2 ** 53;

    const result = await insertImageFromFile(file, {
      imageId: "img_1",
      anchor,
      nextObjectId: unsafeCounterId,
      objects: [],
      images,
    });

    expect(result.objects).toHaveLength(1);
    expect(Number.isSafeInteger(result.objects[0]!.id)).toBe(true);
    expect(result.objects[0]!.id).not.toBe(unsafeCounterId);
  });
});
