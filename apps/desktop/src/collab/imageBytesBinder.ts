import * as Y from "yjs";
import { getYMap } from "@formula/collab-yjs-utils";

import type { CollabSession } from "@formula/collab-session";
import type { ImageEntry, ImageStore } from "../drawings/types";
import { MAX_PNG_DIMENSION, MAX_PNG_PIXELS, readImageDimensions } from "../drawings/pngDimensions";

export type ImageBytesBinder = {
  /**
   * Publish a locally-inserted image to the collaborative Yjs metadata store (size-capped).
   *
   * Call this after inserting the image into the local in-memory `ImageStore`.
   */
  onLocalImageInserted: (image: ImageEntry) => void;
  destroy: () => void;
};

type StoredImageEntry = {
  mimeType: string;
  bytesBase64: string;
};

const DRAWING_IMAGES_KEY = "drawingImages";

// MVP caps: keep updates bounded to avoid giant Yjs updates.
const DEFAULT_MAX_IMAGE_BYTES = 1_000_000; // 1MB raw bytes (base64 is larger)
const DEFAULT_MAX_IMAGES = 100;

function isRecord(value: unknown): value is Record<string, any> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function ensureNestedYMap(parent: any, key: string): Y.Map<any> | null {
  const existing = parent?.get?.(key);
  const existingMap = getYMap(existing);
  if (existingMap) return existingMap;

  // Prefer constructing nested Yjs types using the parent's constructor to tolerate
  // mixed-module environments (ESM + CJS) where multiple `yjs` instances exist.
  // (See `@formula/collab-yjs-utils` for similar cross-instance considerations.)
  const ParentCtor = (parent as any)?.constructor as { new (): any } | undefined;
  const next = (() => {
    if (typeof ParentCtor === "function" && ParentCtor !== Object) {
      try {
        return new ParentCtor();
      } catch {
        // Fall through.
      }
    }
    return new Y.Map();
  })();

  // Best-effort: if the existing value was a plain object, preserve entries.
  if (isRecord(existing)) {
    for (const [k, v] of Object.entries(existing)) {
      next.set(k, v);
    }
  }

  try {
    parent?.set?.(key, next);
  } catch {
    return null;
  }
  return next;
}

/**
 * Estimate decoded bytes without decoding.
 *
 * Assumes the input has already been normalized to a raw base64 string without a `data:` prefix.
 */
function estimateBase64Bytes(base64: string): number {
  const len = base64.length;
  if (len === 0) return 0;
  const padding =
    base64.endsWith("==") ? 2 : base64.endsWith("=") ? 1 : 0;
  return Math.floor((len * 3) / 4) - padding;
}

function normalizeBase64String(raw: string): string | null {
  if (typeof raw !== "string") return null;
  const trimmed = raw.trim();
  if (!trimmed) return null;

  // Strip optional data URL prefix.
  let base64 = trimmed;
  if (base64.startsWith("data:")) {
    const comma = base64.indexOf(",");
    if (comma === -1) return null;
    base64 = base64.slice(comma + 1);
  }

  base64 = base64.trim();
  if (!base64) return null;

  // Be tolerant of whitespace/newlines.
  if (/\s/.test(base64)) {
    base64 = base64.replace(/\s+/g, "");
  }

  // Support base64url (common when copied through URL-safe layers). Normalize to standard
  // base64 and restore optional padding.
  if (base64.includes("-") || base64.includes("_")) {
    base64 = base64.replace(/-/g, "+").replace(/_/g, "/");
  }

  const mod = base64.length % 4;
  if (mod === 1) return null;
  if (mod === 2) base64 += "==";
  else if (mod === 3) base64 += "=";

  return base64 || null;
}

function encodeBase64(bytes: Uint8Array): string | null {
  try {
    if (typeof Buffer !== "undefined") {
      return Buffer.from(bytes).toString("base64");
    }
  } catch {
    // Ignore.
  }

  if (typeof btoa !== "function") return null;

  // Avoid stack overflows by chunking `fromCharCode` calls.
  try {
    let binary = "";
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      // eslint-disable-next-line unicorn/prefer-code-point
      binary += String.fromCharCode(...chunk);
    }
    return btoa(binary);
  } catch {
    return null;
  }
}

function decodeBase64(base64Raw: string, maxBytes: number): Uint8Array | null {
  // Fast length guard before we perform any normalization that could allocate large intermediate
  // strings (trim/replace). In adversarial collaborative documents a malicious peer could insert an
  // extremely long base64 string and attempt to DoS clients by forcing them to repeatedly normalize
  // and decode it.
  //
  // Base64 expands bytes by ~4/3 plus padding. Allow generous overhead for optional `data:` prefixes
  // and whitespace/newlines.
  const roughMaxChars = Math.ceil(((maxBytes + 2) * 4) / 3) + 128;
  if (base64Raw.length > roughMaxChars * 2) return null;

  const base64 = normalizeBase64String(base64Raw);
  if (!base64) return null;

  // Fast size guard before decoding.
  if (estimateBase64Bytes(base64) > maxBytes) return null;

  try {
    if (typeof Buffer !== "undefined") {
      const buf = Buffer.from(base64, "base64");
      if (buf.byteLength > maxBytes) return null;
      // Copy into a right-sized Uint8Array to avoid retaining a larger Buffer slab.
      return new Uint8Array(buf);
    }
  } catch {
    // Ignore.
  }

  if (typeof atob !== "function") return null;

  try {
    const bin = atob(base64);
    if (bin.length > maxBytes) return null;
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i += 1) out[i] = bin.charCodeAt(i);
    return out;
  } catch {
    return null;
  }
}

function coerceStoredImageEntry(raw: unknown, maxBytes: number): { mimeType: string; bytes: Uint8Array } | null {
  // Variant: direct bytes. (mimeType unknown; accept but use a generic fallback.)
  if (raw instanceof Uint8Array) {
    const dims = readImageDimensions(raw);
    if (dims && (dims.width > MAX_PNG_DIMENSION || dims.height > MAX_PNG_DIMENSION || dims.width * dims.height > MAX_PNG_PIXELS)) {
      return null;
    }
    return { mimeType: "application/octet-stream", bytes: raw };
  }

  const map = getYMap(raw);
  if (map) {
    const mimeType = typeof map.get("mimeType") === "string" ? (map.get("mimeType") as string) : "application/octet-stream";
    const bytes = map.get("bytes");
    if (bytes instanceof Uint8Array) {
      const dims = readImageDimensions(bytes);
      if (dims && (dims.width > MAX_PNG_DIMENSION || dims.height > MAX_PNG_DIMENSION || dims.width * dims.height > MAX_PNG_PIXELS)) {
        return null;
      }
      return { mimeType, bytes };
    }
    const bytesBase64 = map.get("bytesBase64");
    if (typeof bytesBase64 === "string") {
      const decoded = decodeBase64(bytesBase64, maxBytes);
      if (!decoded) return null;
      const dims = readImageDimensions(decoded);
      if (dims && (dims.width > MAX_PNG_DIMENSION || dims.height > MAX_PNG_DIMENSION || dims.width * dims.height > MAX_PNG_PIXELS)) {
        return null;
      }
      return { mimeType, bytes: decoded };
    }
    return null;
  }

  if (isRecord(raw)) {
    const mimeType = typeof raw.mimeType === "string" ? (raw.mimeType as string) : "application/octet-stream";
    if (raw.bytes instanceof Uint8Array) {
      const dims = readImageDimensions(raw.bytes);
      if (dims && (dims.width > MAX_PNG_DIMENSION || dims.height > MAX_PNG_DIMENSION || dims.width * dims.height > MAX_PNG_PIXELS)) {
        return null;
      }
      return { mimeType, bytes: raw.bytes };
    }
    if (typeof raw.bytesBase64 === "string") {
      const decoded = decodeBase64(raw.bytesBase64, maxBytes);
      if (!decoded) return null;
      const dims = readImageDimensions(decoded);
      if (dims && (dims.width > MAX_PNG_DIMENSION || dims.height > MAX_PNG_DIMENSION || dims.width * dims.height > MAX_PNG_PIXELS)) {
        return null;
      }
      return { mimeType, bytes: decoded };
    }
  }

  return null;
}

function coerceStoredImageMeta(raw: unknown): StoredImageEntry | null {
  const map = getYMap(raw);
  if (map) {
    const mimeType = typeof map.get("mimeType") === "string" ? (map.get("mimeType") as string) : null;
    const bytesBase64 = typeof map.get("bytesBase64") === "string" ? (map.get("bytesBase64") as string) : null;
    if (!mimeType || !bytesBase64) return null;
    return { mimeType, bytesBase64 };
  }

  if (isRecord(raw)) {
    const mimeType = typeof raw.mimeType === "string" ? (raw.mimeType as string) : null;
    const bytesBase64 = typeof raw.bytesBase64 === "string" ? (raw.bytesBase64 as string) : null;
    if (!mimeType || !bytesBase64) return null;
    return { mimeType, bytesBase64 };
  }

  return null;
}

function enforceMaxImages(map: Y.Map<any>, maxImages: number, keepId: string): void {
  if (!Number.isFinite(maxImages) || maxImages <= 0) return;
  if (map.size <= maxImages) return;

  // Best-effort eviction: remove the oldest entries in iteration order, but never evict `keepId`.
  try {
    const keys = Array.from(map.keys());
    for (const key of keys) {
      if (map.size <= maxImages) break;
      if (key === keepId) continue;
      map.delete(key);
    }
  } catch {
    // Ignore eviction failures.
  }
}

/**
 * Bind an in-memory `ImageStore` to a CollabSession's Yjs metadata map so inserted image bytes
 * can propagate to other collaborators without relying on per-client IndexedDB.
 *
 * This is a best-effort MVP implementation:
 *  - never throws
 *  - size caps (per image + max count)
 *  - idempotent hydration
 */
export function bindImageBytesToCollabSession(options: {
  session: Pick<CollabSession, "doc" | "metadata" | "localOrigins"> | null | undefined;
  images: ImageStore | null | undefined;
  /**
   * Optional stable Yjs transaction origin used for local writes.
   *
   * When omitted, a new per-binder origin token is created.
   */
  origin?: any;
  maxImageBytes?: number;
  maxImages?: number;
}): ImageBytesBinder {
  const session = options?.session ?? null;
  const images = options?.images ?? null;

  if (!session || !images) {
    return { onLocalImageInserted: () => {}, destroy: () => {} };
  }

  const doc = (session as any).doc as Y.Doc | undefined;
  const metadata = (session as any).metadata as Y.Map<any> | undefined;
  if (!doc || !metadata) {
    return { onLocalImageInserted: () => {}, destroy: () => {} };
  }

  const maxImageBytes = Number.isFinite(options?.maxImageBytes ?? NaN)
    ? (options?.maxImageBytes as number)
    : DEFAULT_MAX_IMAGE_BYTES;
  const maxImages = Number.isFinite(options?.maxImages ?? NaN) ? (options?.maxImages as number) : DEFAULT_MAX_IMAGES;

  const ownsOrigin = options?.origin == null;
  const binderOrigin = options?.origin ?? { type: "collab:image-bytes-binder" };
  try {
    (session as any).localOrigins?.add?.(binderOrigin);
  } catch {
    // ignore
  }

  let destroyed = false;

  // Track the last hydrated raw Yjs value to avoid repeatedly decoding base64 when other
  // metadata keys change.
  const hydratedRawValues = new Map<string, unknown>();

  const currentImagesIdsCapped = (): string[] => {
    const limit = Math.max(0, Math.trunc(maxImages));
    if (limit === 0) return [];
    const imagesMap = getImagesMap();
    if (imagesMap) {
      const out: string[] = [];
      try {
        for (const key of imagesMap.keys()) {
          if (out.length >= limit) break;
          if (typeof key === "string") out.push(key);
        }
      } catch {
        // ignore
      }
      return out;
    }
    const record = getImagesRecord();
    if (record) {
      try {
        return Object.keys(record).slice(0, limit);
      } catch {
        return [];
      }
    }
    return [];
  };

  const pruneHydratedRawValues = (): void => {
    // Keep this map bounded to avoid retaining large base64 strings for images that were
    // later evicted/deleted from the shared store.
    try {
      const keepIds = new Set(currentImagesIdsCapped());
      if (keepIds.size === 0) {
        hydratedRawValues.clear();
        return;
      }
      for (const key of Array.from(hydratedRawValues.keys())) {
        if (!keepIds.has(key)) hydratedRawValues.delete(key);
      }
    } catch {
      // ignore
    }
  };

  const getImagesContainerRaw = (): unknown => {
    try {
      return metadata.get(DRAWING_IMAGES_KEY);
    } catch {
      return null;
    }
  };

  const getImagesMap = (): Y.Map<any> | null => getYMap(getImagesContainerRaw());

  const getImagesRecord = (): Record<string, any> | null => {
    const raw = getImagesContainerRaw();
    return isRecord(raw) ? (raw as Record<string, any>) : null;
  };

  // Only create the nested Y.Map when we have bytes to publish. Avoid writing an empty map
  // during binder setup so we don't emit spurious Yjs updates on document open.
  const ensureImagesMapForWrite = (): Y.Map<any> | null => ensureNestedYMap(metadata, DRAWING_IMAGES_KEY);

  const hydrateImageIds = (ids: Iterable<string> | null): void => {
    if (destroyed) return;
    const imagesMap = getImagesMap();
    const imagesRecord = imagesMap ? null : getImagesRecord();
    if (!imagesMap && !imagesRecord) return;

    const toHydrate = ids
      ? Array.from(ids)
      : imagesMap
        ? Array.from(imagesMap.keys())
        : Object.keys(imagesRecord ?? {});

    // Defensive cap: avoid decoding an unbounded number of images even if a doc is corrupt/malicious.
    const capped = toHydrate.slice(0, Math.max(0, Math.trunc(maxImages)));

    for (const imageId of capped) {
      try {
        const raw = imagesMap ? imagesMap.get(imageId) : (imagesRecord as any)?.[imageId];
        if (!raw) continue;

        // Avoid re-decoding the exact same Yjs value (common when multiple observe events fire).
        const prevRaw = hydratedRawValues.get(imageId);
        // For nested Y.Maps, updates can change inner fields without changing the outer value reference.
        // In that case, rely on the Yjs event filtering (changedIds) and re-hydrate.
        if (prevRaw === raw && !getYMap(raw)) continue;

        // If we already have bytes locally for this image id (e.g. from IndexedDB persistence),
        // avoid re-decoding base64 and re-writing the ImageStore.
        //
        // For the MVP picture insertion flow, image ids are treated as immutable, so matching
        // byteLength + mimeType is sufficient to consider the image "already hydrated".
        const existingLocal = images.get(imageId);
        if (existingLocal?.bytes instanceof Uint8Array) {
          const meta = coerceStoredImageMeta(raw);
          if (meta) {
            const normalized = normalizeBase64String(meta.bytesBase64);
            if (normalized) {
              const estimatedBytes = estimateBase64Bytes(normalized);
              if (
                estimatedBytes <= maxImageBytes &&
                existingLocal.bytes.byteLength === estimatedBytes &&
                String(existingLocal.mimeType ?? "") === meta.mimeType
              ) {
                hydratedRawValues.set(imageId, raw);
                continue;
              }
            }
          } else if (raw instanceof Uint8Array && existingLocal.bytes.byteLength === raw.byteLength) {
            hydratedRawValues.set(imageId, raw);
            continue;
          }
        }

        const entry = coerceStoredImageEntry(raw, maxImageBytes);
        if (!entry) {
          hydratedRawValues.set(imageId, raw);
          continue;
        }

        if (entry.bytes.byteLength > maxImageBytes) {
          hydratedRawValues.set(imageId, raw);
          continue;
        }

        // Populate/overwrite the in-memory store. (Idempotent for Map-backed stores.)
        images.set({ id: imageId, bytes: entry.bytes, mimeType: entry.mimeType });
        hydratedRawValues.set(imageId, raw);
      } catch {
        // Ignore hydration errors.
      }
    }
  };

  const hydrateAll = () => hydrateImageIds(null);

  const handleMetadataDeepChange = (events: any[], transaction: Y.Transaction) => {
    if (destroyed) return;
    if (!events || events.length === 0) return;

    const origin = transaction?.origin ?? null;
    if (origin === binderOrigin) return;

    const imagesMap = getImagesMap();

    let shouldHydrateAll = false;
    const changedIds = new Set<string>();

    for (const event of events) {
      const path = event?.path;
      if (Array.isArray(path) && path.length >= 2 && path[0] === DRAWING_IMAGES_KEY && typeof path[1] === "string") {
        // If a future schema stores nested Yjs types (e.g. a Y.Map per image id), we can still
        // hydrate the updated image by looking at the second path component.
        changedIds.add(path[1]);
      }

      if (event?.target === metadata) {
        const keys = event?.changes?.keys;
        if (keys && typeof keys.has === "function" && keys.has(DRAWING_IMAGES_KEY)) {
          shouldHydrateAll = true;
        }
      }

      if (imagesMap && event?.target === imagesMap) {
        const keys = event?.changes?.keys;
        if (keys && typeof keys.keys === "function") {
          for (const key of keys.keys()) {
            if (typeof key === "string") changedIds.add(key);
          }
        } else {
          shouldHydrateAll = true;
        }
      }
    }

    // If we don't have an images container yet, only hydrate on root-level changes.
    if (!imagesMap && !shouldHydrateAll) return;

    if (shouldHydrateAll) {
      hydrateAll();
      pruneHydratedRawValues();
      return;
    }

    if (changedIds.size > 0) {
      hydrateImageIds(changedIds);
    }

    // Even when only a subset of ids changed, prune any cached raw values for ids that may have
    // been deleted/evicted from the Yjs map so we don't retain large base64 strings indefinitely.
    pruneHydratedRawValues();
  };

  try {
    metadata.observeDeep(handleMetadataDeepChange);
  } catch {
    // ignore
  }

  // Initial hydration (and for cases where the provider has already applied state).
  hydrateAll();

  const onLocalImageInserted = (image: ImageEntry) => {
    if (destroyed) return;

    try {
      const imageId = typeof image?.id === "string" ? image.id : null;
      if (!imageId) return;
      const mimeType = typeof image?.mimeType === "string" ? image.mimeType : "application/octet-stream";
      const bytes = image?.bytes;
      if (!(bytes instanceof Uint8Array)) return;

      if (bytes.byteLength > maxImageBytes) return;
      const bytesBase64 = encodeBase64(bytes);
      if (!bytesBase64) return;

      doc.transact(
        () => {
          const imagesMap = ensureImagesMapForWrite();
          if (!imagesMap) return;

          const existing = imagesMap.get(imageId);
          // Avoid redundant overwrites (and repeated Yjs updates).
          const existingMeta = coerceStoredImageMeta(existing);
          if (existingMeta && existingMeta.mimeType === mimeType && existingMeta.bytesBase64 === bytesBase64) {
            return;
          }

          const entry: StoredImageEntry = { mimeType, bytesBase64 };
          imagesMap.set(imageId, entry);
          enforceMaxImages(imagesMap, maxImages, imageId);
        },
        binderOrigin,
      );

       // Local writes use `binderOrigin` and therefore won't flow through the deep-observer
       // (it early-returns to avoid echo). Prune here so we don't retain cached raw values
       // for images that were evicted by `enforceMaxImages` in this transaction.
       pruneHydratedRawValues();
    } catch {
      // ignore
    }
  };

  return {
    onLocalImageInserted,
    destroy() {
      if (destroyed) return;
      destroyed = true;
      hydratedRawValues.clear();
      try {
        metadata.unobserveDeep(handleMetadataDeepChange);
      } catch {
        // ignore
      }
      if (ownsOrigin) {
        try {
          (session as any).localOrigins?.delete?.(binderOrigin);
        } catch {
          // ignore
        }
      }
    },
  };
}
