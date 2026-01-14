// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";

import { ImageBitmapCache } from "../imageBitmapCache";
import type { ImageEntry } from "../types";

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("ImageBitmapCache decode fallback", () => {
  it("falls back to <img> + canvas when createImageBitmap(blob) throws InvalidStateError", async () => {
    const entry: ImageEntry = { id: "img_fallback", bytes: new Uint8Array([1, 2, 3, 4]), mimeType: "image/png" };

    const close = vi.fn();
    const decoded = { close } as unknown as ImageBitmap;

    const createImageBitmapMock = vi.fn((src: any) => {
      if (src instanceof Blob) {
        const err = new Error("decode failed");
        (err as any).name = "InvalidStateError";
        return Promise.reject(err);
      }
      return Promise.resolve(decoded);
    });
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    // jsdom does not implement URL.createObjectURL; stub it (and revokeObjectURL) while keeping the
    // URL constructor intact.
    const URLCtor = globalThis.URL as any;
    const originalCreateObjectURL = URLCtor?.createObjectURL;
    const originalRevokeObjectURL = URLCtor?.revokeObjectURL;
    const createObjectURL = vi.fn(() => "blob:fake");
    const revokeObjectURL = vi.fn();
    URLCtor.createObjectURL = createObjectURL;
    URLCtor.revokeObjectURL = revokeObjectURL;

    try {
      // Provide a deterministic Image stub that fires `onload` asynchronously *after* the handler is
      // assigned (the production code sets `src` before wiring `onload`).
      class FakeImage {
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;
        width = 16;
        height = 8;
        naturalWidth = 16;
        naturalHeight = 8;
        set src(_value: string) {
          queueMicrotask(() => {
            this.onload?.();
          });
        }
      }
      vi.stubGlobal("Image", FakeImage as unknown as typeof Image);

      // jsdom canvas elements return null for getContext; provide a minimal 2D context.
      const originalGetContext = HTMLCanvasElement.prototype.getContext;
      HTMLCanvasElement.prototype.getContext = vi.fn(() => ({ drawImage: vi.fn() }) as any);
      try {
        const cache = new ImageBitmapCache({ maxEntries: 16, negativeCacheMs: 0 });
        await expect(cache.get(entry)).resolves.toBe(decoded);
      } finally {
        HTMLCanvasElement.prototype.getContext = originalGetContext;
      }

      // One attempt via blob, one via canvas fallback.
      expect(createImageBitmapMock).toHaveBeenCalledTimes(2);
      expect(createImageBitmapMock.mock.calls[0]?.[0]).toBeInstanceOf(Blob);
      expect(createObjectURL).toHaveBeenCalledTimes(1);
      expect(revokeObjectURL).toHaveBeenCalledTimes(1);
    } finally {
      // Restore URL static methods.
      if (originalCreateObjectURL === undefined) delete URLCtor.createObjectURL;
      else URLCtor.createObjectURL = originalCreateObjectURL;
      if (originalRevokeObjectURL === undefined) delete URLCtor.revokeObjectURL;
      else URLCtor.revokeObjectURL = originalRevokeObjectURL;
    }
  });
});

