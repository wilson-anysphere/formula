import { createDrawingObjectId, type Anchor, type DrawingObject, type ImageEntry, type ImageStore } from "./types";
import { MAX_INSERT_IMAGE_BYTES } from "./insertImageLimits.js";
import { MAX_PNG_DIMENSION, MAX_PNG_PIXELS, readImageDimensions } from "./pngDimensions";

// Keep this in sync with the clipboard image guard (currently 5MiB raw PNG bytes).
// This is only the *default* for base64 decoding helpers; callers may override via `opts.maxBytes`.
const DEFAULT_MAX_BASE64_IMAGE_BYTES = 5 * 1024 * 1024; // 5MiB

export { MAX_INSERT_IMAGE_BYTES };

function nextUniqueDrawingObjectId(existing: Set<number>): number {
  let max = 0;
  for (const id of existing) {
    if (id > max) max = id;
  }

  // Prefer collision-resistant ids for multi-user safety, but guarantee termination even if
  // WebCrypto is stubbed/deterministic in tests.
  for (let attempt = 0; attempt < 10; attempt += 1) {
    const candidate = createDrawingObjectId();
    if (!existing.has(candidate)) return candidate;
  }

  // Fallback: find a deterministic unused safe integer id.
  const candidate = max + 1;
  if (Number.isSafeInteger(candidate) && candidate > 0 && !existing.has(candidate)) return candidate;

  let next = 1;
  while (existing.has(next) && next < Number.MAX_SAFE_INTEGER) next += 1;
  return next;
}

export async function insertImageFromFile(
  file: File,
  opts: {
    imageId: string;
    anchor: Anchor;
    /**
     * @deprecated Drawing object ids must be globally unique across collaborators; callers should not
     *             pass incrementing counters anymore. This field is ignored.
     */
    nextObjectId?: number;
    objects: DrawingObject[];
    images: ImageStore;
  },
): Promise<{ objects: DrawingObject[]; image: ImageEntry }> {
  // Defensive guard: callers should filter before calling (so we can present good UX),
  // but keep a hard stop here to avoid unbounded allocations if a new entry point
  // forgets to enforce limits.
  if (typeof file.size === "number" && file.size > MAX_INSERT_IMAGE_BYTES) {
    throw new Error(`insertImageFromFile: image too large (${file.size} bytes)`);
  }
  const bytes = await readFileBytes(file);
  if (bytes.byteLength > MAX_INSERT_IMAGE_BYTES) {
    throw new Error(`insertImageFromFile: image too large (${bytes.byteLength} bytes)`);
  }
  const dims = readImageDimensions(bytes);
  if (
    dims &&
    (dims.width > MAX_PNG_DIMENSION || dims.height > MAX_PNG_DIMENSION || dims.width * dims.height > MAX_PNG_PIXELS)
  ) {
    throw new Error(`insertImageFromFile: image dimensions too large (${dims.width}x${dims.height})`);
  }
  const mimeType = file.type || guessMimeType(file.name);
  const image: ImageEntry = { id: opts.imageId, bytes, mimeType };
  opts.images.set(image);

  const nextZOrder = (() => {
    // Use max-zOrder so inserts remain on top even if prior objects were re-ordered or deleted.
    let max = -1;
    for (const obj of opts.objects) {
      if (typeof obj.zOrder === "number" && Number.isFinite(obj.zOrder)) {
        max = Math.max(max, obj.zOrder);
      }
    }
    return max + 1;
  })();

  // Random ids are collision-resistant across collaborators, but still ensure we don't collide
  // with existing objects within this sheet.
  const usedIds = new Set(opts.objects.map((o) => o.id));
  const objectId = nextUniqueDrawingObjectId(usedIds);

  const object: DrawingObject = {
    id: objectId,
    kind: { type: "image", imageId: image.id },
    anchor: opts.anchor,
    zOrder: nextZOrder,
    size: opts.anchor.type === "oneCell" || opts.anchor.type === "absolute" ? opts.anchor.size : undefined,
  };

  return { objects: [...opts.objects, object], image };
}

export function insertImageFromBytes(
  bytes: Uint8Array,
  opts: {
    imageId: string;
    mimeType: string;
    anchor: Anchor;
    /**
     * @deprecated Drawing object ids must be globally unique across collaborators; callers should not
     *             pass incrementing counters anymore. This field is ignored.
     */
    nextObjectId?: number;
    objects: DrawingObject[];
    images: ImageStore;
  },
): { objects: DrawingObject[]; image: ImageEntry } {
  // Defensive guard: callers should enforce limits upstream, but keep hard stops here so new
  // entry points cannot accidentally persist oversized images.
  if (!(bytes instanceof Uint8Array)) {
    throw new Error("insertImageFromBytes: bytes must be a Uint8Array");
  }
  if (bytes.byteLength > MAX_INSERT_IMAGE_BYTES) {
    throw new Error(`insertImageFromBytes: image too large (${bytes.byteLength} bytes)`);
  }
  const dims = readImageDimensions(bytes);
  if (
    dims &&
    (dims.width > MAX_PNG_DIMENSION || dims.height > MAX_PNG_DIMENSION || dims.width * dims.height > MAX_PNG_PIXELS)
  ) {
    throw new Error(`insertImageFromBytes: image dimensions too large (${dims.width}x${dims.height})`);
  }

  const image: ImageEntry = { id: opts.imageId, bytes, mimeType: opts.mimeType };
  opts.images.set(image);

  const nextZOrder = (() => {
    let max = -1;
    for (const obj of opts.objects) {
      if (typeof obj.zOrder === "number" && Number.isFinite(obj.zOrder)) {
        max = Math.max(max, obj.zOrder);
      }
    }
    return max + 1;
  })();

  const usedIds = new Set(opts.objects.map((o) => o.id));
  const objectId = nextUniqueDrawingObjectId(usedIds);

  const object: DrawingObject = {
    id: objectId,
    kind: { type: "image", imageId: image.id },
    anchor: opts.anchor,
    zOrder: nextZOrder,
    size: opts.anchor.type === "oneCell" || opts.anchor.type === "absolute" ? opts.anchor.size : undefined,
  };

  return { objects: [...opts.objects, object], image };
}

/**
 * Decode a (potentially `data:*;base64,`-prefixed) base64 string into bytes.
 *
 * Intended for legacy clipboard paths that still surface `pngBase64`.
 */
export function decodeBase64ToBytes(base64: string, opts: { maxBytes?: number } = {}): Uint8Array | null {
  if (typeof base64 !== "string") return null;
  const maxBytes = Number.isFinite(opts.maxBytes) ? Number(opts.maxBytes) : DEFAULT_MAX_BASE64_IMAGE_BYTES;

  let trimmed = base64.trim();
  if (!trimmed) return null;

  // Strip `data:*;base64,` prefix if present.
  if (trimmed.startsWith("data:")) {
    const comma = trimmed.indexOf(",");
    if (comma === -1) return null;
    trimmed = trimmed.slice(comma + 1).trim();
    if (!trimmed) return null;
  }

  // Rough size estimate before decode to avoid allocating huge buffers.
  const len = trimmed.length;
  const padding = trimmed.endsWith("==") ? 2 : trimmed.endsWith("=") ? 1 : 0;
  const estimated = Math.max(0, Math.floor((len * 3) / 4) - padding);
  if (estimated > maxBytes) return null;

  try {
    if (typeof Buffer !== "undefined") {
      // eslint-disable-next-line no-undef
      const buf = Buffer.from(trimmed, "base64");
      if (buf.byteLength > maxBytes) return null;
      return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
    }
  } catch {
    // Fall through to atob.
  }

  try {
    if (typeof atob === "function") {
      const bin = atob(trimmed);
      if (bin.length > maxBytes) return null;
      const out = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i += 1) {
        out[i] = bin.charCodeAt(i);
      }
      return out;
    }
  } catch {
    // Ignore.
  }

  return null;
}

async function readFileBytes(file: Blob): Promise<Uint8Array> {
  const anyFile = file as any;
  if (typeof anyFile?.arrayBuffer === "function") {
    const buffer: ArrayBuffer = await anyFile.arrayBuffer();
    return new Uint8Array(buffer);
  }

  // JSDOM's `File` implementation does not always provide `arrayBuffer()`. Prefer
  // `FileReader` when available since it exists in both browsers and jsdom.
  const FileReaderCtor = (globalThis as any)?.FileReader as typeof FileReader | undefined;
  if (typeof FileReaderCtor === "function") {
    return await new Promise((resolve, reject) => {
      const reader = new FileReaderCtor();
      reader.onload = () => {
        const result = reader.result;
        if (result instanceof ArrayBuffer) {
          resolve(new Uint8Array(result));
          return;
        }
        reject(new Error("FileReader did not return an ArrayBuffer"));
      };
      reader.onerror = () => reject(reader.error ?? new Error("Failed to read file bytes"));
      try {
        reader.readAsArrayBuffer(file);
      } catch (err) {
        reject(err);
      }
    });
  }

  // Final fallback: use Fetch's Body mixin if available.
  const ResponseCtor = (globalThis as any)?.Response as typeof Response | undefined;
  if (typeof ResponseCtor === "function") {
    const buffer = await new ResponseCtor(file as any).arrayBuffer();
    return new Uint8Array(buffer);
  }

  throw new Error("Unable to read file bytes: File.arrayBuffer/FileReader/Response unavailable");
}

function guessMimeType(name: string): string {
  const ext = name.split(".").pop()?.toLowerCase();
  switch (ext) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "webp":
      return "image/webp";
    case "svg":
      return "image/svg+xml";
    default:
      return "application/octet-stream";
  }
}
