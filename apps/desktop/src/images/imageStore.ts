export type StoredImage = { bytes: Uint8Array; mimeType: string };

/**
 * Lightweight in-memory image store for in-cell images.
 *
 * This intentionally mirrors the shape we expect from the workbook backend
 * (imageId -> {bytes, mimeType}) so replacing demo seeding with real workbook
 * image hydration is trivial.
 */
export class DesktopImageStore {
  private readonly images = new Map<string, StoredImage>();

  clear(): void {
    this.images.clear();
  }

  set(imageId: string, entry: StoredImage): void {
    this.images.set(String(imageId), entry);
  }

  delete(imageId: string): void {
    this.images.delete(String(imageId));
  }

  get(imageId: string): StoredImage | null {
    return this.images.get(String(imageId)) ?? null;
  }

  getImageBlob(imageId: string): Blob | null {
    const entry = this.get(imageId);
    if (!entry) return null;
    if (typeof Blob === "undefined") return null;
    try {
      return new Blob([entry.bytes], { type: entry.mimeType || "application/octet-stream" });
    } catch {
      return null;
    }
  }
}
