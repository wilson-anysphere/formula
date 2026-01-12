import type { ImageEntry } from "./types";

/**
 * Cache for decoded images.
 *
 * `createImageBitmap` is asynchronous and relatively expensive, so we keep a
 * Promise per imageId to dedupe concurrent requests.
 */
export class ImageBitmapCache {
  private readonly bitmaps = new Map<string, Promise<ImageBitmap>>();

  async get(entry: ImageEntry): Promise<ImageBitmap> {
    const existing = this.bitmaps.get(entry.id);
    if (existing) return existing;

    const promise = (async () => {
      const buffer = new ArrayBuffer(entry.bytes.byteLength);
      new Uint8Array(buffer).set(entry.bytes);
      const blob = new Blob([buffer], { type: entry.mimeType });
      return await createImageBitmap(blob);
    })();
    this.bitmaps.set(entry.id, promise);
    return promise;
  }

  invalidate(imageId: string): void {
    this.bitmaps.delete(imageId);
  }

  clear(): void {
    this.bitmaps.clear();
  }
}
