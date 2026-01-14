import { createDrawingObjectId, type Anchor, type DrawingObject, type ImageEntry, type ImageStore } from "./types";

// Keep this in sync with clipboard provider / Tauri clipboard guards.
const DEFAULT_MAX_IMAGE_BYTES = 5 * 1024 * 1024; // 5MB (raw bytes)

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
  const bytes = await readFileBytes(file);
  const mimeType = file.type || guessMimeType(file.name);
  const image: ImageEntry = { id: opts.imageId, bytes, mimeType };
  opts.images.set(image);

  const object: DrawingObject = {
    id: createDrawingObjectId(),
    kind: { type: "image", imageId: image.id },
    anchor: opts.anchor,
    zOrder: opts.objects.length,
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
  const image: ImageEntry = { id: opts.imageId, bytes, mimeType: opts.mimeType };
  opts.images.set(image);

  const object: DrawingObject = {
    id: createDrawingObjectId(),
    kind: { type: "image", imageId: image.id },
    anchor: opts.anchor,
    zOrder: opts.objects.length,
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
  const maxBytes = Number.isFinite(opts.maxBytes) ? Number(opts.maxBytes) : DEFAULT_MAX_IMAGE_BYTES;

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

async function readFileBytes(file: File): Promise<Uint8Array> {
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
    default:
      return "application/octet-stream";
  }
}
