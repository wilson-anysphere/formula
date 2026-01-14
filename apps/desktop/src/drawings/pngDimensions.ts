// Guardrails for PNG decompression bombs: a PNG can advertise extremely large dimensions while the
// compressed payload remains small. Decoding such images can allocate huge bitmaps and hang/crash
// the renderer.
//
// These limits are shared across:
// - insert/paste flows (so rejected images are never persisted)
// - collaboration hydration (so remote peers can't DoS by publishing huge-dimension PNGs)
// - bitmap decode (defense-in-depth for legacy/persisted images)
export const MAX_PNG_DIMENSION = 10_000;
export const MAX_PNG_PIXELS = 50_000_000;

export function readPngDimensions(bytes: Uint8Array): { width: number; height: number } | null {
  // PNG signature (8 bytes) + IHDR chunk header (8 bytes) + width/height (8 bytes).
  if (!(bytes instanceof Uint8Array) || bytes.byteLength < 24) return null;

  if (
    bytes[0] !== 0x89 ||
    bytes[1] !== 0x50 ||
    bytes[2] !== 0x4e ||
    bytes[3] !== 0x47 ||
    bytes[4] !== 0x0d ||
    bytes[5] !== 0x0a ||
    bytes[6] !== 0x1a ||
    bytes[7] !== 0x0a
  ) {
    return null;
  }

  // The first chunk should be IHDR: length (4) + type (4) + data...
  if (bytes[12] !== 0x49 || bytes[13] !== 0x48 || bytes[14] !== 0x44 || bytes[15] !== 0x52) {
    return null;
  }

  try {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const width = view.getUint32(16, false);
    const height = view.getUint32(20, false);
    if (width === 0 || height === 0) return null;
    return { width, height };
  } catch {
    return null;
  }
}

