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
      // Provide a deterministic Image stub that fires `onload` synchronously when `src` is assigned.
      // This ensures our fallback decode path wires handlers before setting `src` (robust to
      // synchronous load events in tests/polyfills).
      class FakeImage {
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;
        width = 16;
        height = 8;
        naturalWidth = 16;
        naturalHeight = 8;
        set src(_value: string) {
          this.onload?.();
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

  it("rethrows InvalidStateError when the <img> fallback fails to load, and still revokes the object URL", async () => {
    const entry: ImageEntry = { id: "img_fallback_error", bytes: new Uint8Array([1, 2, 3, 4]), mimeType: "image/png" };

    const err = new Error("decode failed");
    (err as any).name = "InvalidStateError";
    const createImageBitmapMock = vi.fn(() => Promise.reject(err));
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const URLCtor = globalThis.URL as any;
    const originalCreateObjectURL = URLCtor?.createObjectURL;
    const originalRevokeObjectURL = URLCtor?.revokeObjectURL;
    const createObjectURL = vi.fn(() => "blob:fake");
    const revokeObjectURL = vi.fn();
    URLCtor.createObjectURL = createObjectURL;
    URLCtor.revokeObjectURL = revokeObjectURL;

    try {
      class FakeImage {
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;
        set src(_value: string) {
          queueMicrotask(() => {
            this.onerror?.();
          });
        }
      }
      vi.stubGlobal("Image", FakeImage as unknown as typeof Image);

      const cache = new ImageBitmapCache({ maxEntries: 16, negativeCacheMs: 0 });
      await expect(cache.get(entry)).rejects.toMatchObject({ name: "InvalidStateError" });

      // Should attempt the blob decode once; the image-element fallback does not invoke createImageBitmap again
      // because the <img> load failed before we could render into a canvas.
      expect(createImageBitmapMock).toHaveBeenCalledTimes(1);
      expect(createObjectURL).toHaveBeenCalledTimes(1);
      expect(revokeObjectURL).toHaveBeenCalledTimes(1);
    } finally {
      if (originalCreateObjectURL === undefined) delete URLCtor.createObjectURL;
      else URLCtor.createObjectURL = originalCreateObjectURL;
      if (originalRevokeObjectURL === undefined) delete URLCtor.revokeObjectURL;
      else URLCtor.revokeObjectURL = originalRevokeObjectURL;
    }
  });

  it("rethrows InvalidStateError when createImageBitmap(canvas) fails, and still revokes the object URL", async () => {
    const entry: ImageEntry = { id: "img_fallback_canvas_error", bytes: new Uint8Array([1, 2, 3, 4]), mimeType: "image/png" };

    const invalidState = new Error("decode failed");
    (invalidState as any).name = "InvalidStateError";
    const createImageBitmapMock = vi.fn((src: any) => {
      if (src instanceof Blob) {
        return Promise.reject(invalidState);
      }
      return Promise.reject(new Error("bitmap allocation failed"));
    });
    vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

    const URLCtor = globalThis.URL as any;
    const originalCreateObjectURL = URLCtor?.createObjectURL;
    const originalRevokeObjectURL = URLCtor?.revokeObjectURL;
    const createObjectURL = vi.fn(() => "blob:fake");
    const revokeObjectURL = vi.fn();
    URLCtor.createObjectURL = createObjectURL;
    URLCtor.revokeObjectURL = revokeObjectURL;

    try {
      class FakeImage {
        onload: (() => void) | null = null;
        onerror: (() => void) | null = null;
        width = 16;
        height = 8;
        naturalWidth = 16;
        naturalHeight = 8;
        set src(_value: string) {
          this.onload?.();
        }
      }
      vi.stubGlobal("Image", FakeImage as unknown as typeof Image);

      const originalGetContext = HTMLCanvasElement.prototype.getContext;
      HTMLCanvasElement.prototype.getContext = vi.fn(() => ({ drawImage: vi.fn() }) as any);
      try {
        const cache = new ImageBitmapCache({ maxEntries: 16, negativeCacheMs: 0 });
        await expect(cache.get(entry)).rejects.toMatchObject({ name: "InvalidStateError" });
      } finally {
        HTMLCanvasElement.prototype.getContext = originalGetContext;
      }

      expect(createImageBitmapMock).toHaveBeenCalledTimes(2);
      expect(createImageBitmapMock.mock.calls[0]?.[0]).toBeInstanceOf(Blob);
      expect(createImageBitmapMock.mock.calls[1]?.[0]).toBeInstanceOf(HTMLCanvasElement);
      expect(createObjectURL).toHaveBeenCalledTimes(1);
      expect(revokeObjectURL).toHaveBeenCalledTimes(1);
    } finally {
      if (originalCreateObjectURL === undefined) delete URLCtor.createObjectURL;
      else URLCtor.createObjectURL = originalCreateObjectURL;
      if (originalRevokeObjectURL === undefined) delete URLCtor.revokeObjectURL;
      else URLCtor.revokeObjectURL = originalRevokeObjectURL;
    }
  });

  it("times out if the <img> fallback never resolves, and still revokes the object URL", async () => {
    vi.useFakeTimers();
    try {
      const entry: ImageEntry = { id: "img_fallback_timeout", bytes: new Uint8Array([1, 2, 3, 4]), mimeType: "image/png" };

      const invalidState = new Error("decode failed");
      (invalidState as any).name = "InvalidStateError";
      const createImageBitmapMock = vi.fn(() => Promise.reject(invalidState));
      vi.stubGlobal("createImageBitmap", createImageBitmapMock as unknown as typeof createImageBitmap);

      const URLCtor = globalThis.URL as any;
      const originalCreateObjectURL = URLCtor?.createObjectURL;
      const originalRevokeObjectURL = URLCtor?.revokeObjectURL;
      const createObjectURL = vi.fn(() => "blob:fake");
      const revokeObjectURL = vi.fn();
      URLCtor.createObjectURL = createObjectURL;
      URLCtor.revokeObjectURL = revokeObjectURL;

      try {
        class FakeImage {
          onload: (() => void) | null = null;
          onerror: (() => void) | null = null;
          set src(_value: string) {
            // Intentionally never call onload/onerror.
          }
        }
        vi.stubGlobal("Image", FakeImage as unknown as typeof Image);

        const cache = new ImageBitmapCache({ maxEntries: 16, negativeCacheMs: 0 });
        const promise = cache.get(entry);

        // Allow the initial createImageBitmap rejection -> fallback path to schedule its timeout.
        await Promise.resolve();
        await Promise.resolve();

        // The fallback times out after 5s; the cache should rethrow the original InvalidStateError.
        vi.advanceTimersByTime(5_000);
        await expect(promise).rejects.toMatchObject({ name: "InvalidStateError" });

        expect(createImageBitmapMock).toHaveBeenCalledTimes(1);
        expect(createObjectURL).toHaveBeenCalledTimes(1);
        expect(revokeObjectURL).toHaveBeenCalledTimes(1);
      } finally {
        if (originalCreateObjectURL === undefined) delete URLCtor.createObjectURL;
        else URLCtor.createObjectURL = originalCreateObjectURL;
        if (originalRevokeObjectURL === undefined) delete URLCtor.revokeObjectURL;
        else URLCtor.revokeObjectURL = originalRevokeObjectURL;
      }
    } finally {
      vi.useRealTimers();
    }
  });
});
