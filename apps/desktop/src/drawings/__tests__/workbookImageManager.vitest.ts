import { describe, expect, it, vi } from "vitest";

import { InMemoryImageStore, type DrawingObject, type ImageEntry } from "../types";
import { WorkbookImageManager } from "../workbookImageManager";

function createImageObject(id: number, imageId: string): DrawingObject {
  return {
    id,
    kind: { type: "image", imageId },
    anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 0, cy: 0 } },
    zOrder: id,
  };
}

describe("WorkbookImageManager", () => {
  it("retains image bytes when at least one drawing still references the imageId", async () => {
    const store = new InMemoryImageStore();
    const entry: ImageEntry = { id: "img1", bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" };
    store.set(entry);

    const persistence = { delete: vi.fn().mockResolvedValue(undefined) };
    const bitmapCache = { invalidate: vi.fn() };
    const manager = new WorkbookImageManager({
      images: store,
      persistence,
      bitmapCache,
      gcGracePeriodMs: 1_000,
    });

    manager.setSheetDrawings("Sheet1", [createImageObject(1, "img1"), createImageObject(2, "img1")]);
    expect(manager.imageRefCount.get("img1")).toBe(2);

    // Delete one drawing; image still referenced.
    manager.setSheetDrawings("Sheet1", [createImageObject(1, "img1")]);
    expect(manager.imageRefCount.get("img1")).toBe(1);

    await manager.runGcNow();
    expect(store.get("img1")).toBeDefined();
    expect(persistence.delete).not.toHaveBeenCalled();
    expect(bitmapCache.invalidate).not.toHaveBeenCalled();
  });

  it("defers GC when the last reference is removed, then collects after the grace period", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(0);

    const store = new InMemoryImageStore();
    const entry: ImageEntry = { id: "img2", bytes: new Uint8Array([9, 9, 9]), mimeType: "image/png" };
    store.set(entry);

    const persistence = { delete: vi.fn().mockResolvedValue(undefined) };
    const bitmapCache = { invalidate: vi.fn() };
    const manager = new WorkbookImageManager({
      images: store,
      persistence,
      bitmapCache,
      gcGracePeriodMs: 1_000,
    });

    manager.setSheetDrawings("Sheet1", [createImageObject(1, "img2")]);
    expect(manager.imageRefCount.get("img2")).toBe(1);

    // Delete the last drawing referencing img2.
    manager.setSheetDrawings("Sheet1", []);
    expect(manager.imageRefCount.get("img2")).toBeUndefined();

    // Still within grace period; do not collect.
    await manager.runGcNow();
    expect(store.get("img2")).toBeDefined();
    expect(persistence.delete).not.toHaveBeenCalled();

    // Advance beyond the grace period and collect.
    vi.setSystemTime(1_000);
    await manager.runGcNow();
    expect(store.get("img2")).toBeUndefined();
    expect(persistence.delete).toHaveBeenCalledWith("img2");
    expect(bitmapCache.invalidate).toHaveBeenCalledWith("img2");

    manager.dispose();
    vi.useRealTimers();
  });
});

