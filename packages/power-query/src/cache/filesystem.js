import { fnv1a64 } from "./key.js";

/**
 * @typedef {import("./cache.js").CacheEntry} CacheEntry
 */

const BINARY_MARKER_KEY = "__pq_cache_binary";

/**
 * @param {unknown} value
 * @returns {value is Record<string, unknown> & { [BINARY_MARKER_KEY]: string }}
 */
function isBinaryMarker(value) {
  return (
    value != null &&
    typeof value === "object" &&
    !Array.isArray(value) &&
    // @ts-ignore - runtime indexing
    typeof value[BINARY_MARKER_KEY] === "string"
  );
}

/**
 * @param {unknown} value
 * @returns {boolean}
 */
function containsBinaryMarker(value) {
  if (isBinaryMarker(value)) return true;
  if (Array.isArray(value)) return value.some((item) => containsBinaryMarker(item));
  if (value && typeof value === "object") {
    for (const v of Object.values(value)) {
      if (containsBinaryMarker(v)) return true;
    }
  }
  return false;
}

/**
 * Replace all Uint8Array values in an object graph with marker objects that
 * reference a sibling `.bin` file.
 *
 * @param {unknown} value
 * @param {string} binFileName
 * @returns {{ value: unknown, segments: Array<{ offset: number, length: number, bytes: Uint8Array }> }}
 */
function extractBinarySegments(value, binFileName) {
  /** @type {Array<{ offset: number, length: number, bytes: Uint8Array }>} */
  const segments = [];
  let offset = 0;

  /**
   * @param {unknown} current
   * @returns {unknown}
   */
  function visit(current) {
    if (current instanceof Uint8Array) {
      const segmentOffset = offset;
      const length = current.byteLength;
      offset += length;
      segments.push({ offset: segmentOffset, length, bytes: current });
      return { [BINARY_MARKER_KEY]: binFileName, offset: segmentOffset, length };
    }

    if (Array.isArray(current)) {
      return current.map((item) => visit(item));
    }

    if (current && typeof current === "object") {
      // Respect `toJSON()` for non-plain objects (e.g. Date, URL) so we don't
      // accidentally serialize them as `{}` when extracting binary segments.
      // (Buffers are handled above via the Uint8Array branch.)
      // @ts-ignore - runtime inspection
      if (typeof current.toJSON === "function" && current.constructor && current.constructor !== Object) {
        // @ts-ignore - runtime
        return visit(current.toJSON());
      }

      const out = {};
      for (const [k, v] of Object.entries(current)) {
        // @ts-ignore - runtime indexing
        out[k] = visit(v);
      }
      return out;
    }

    return current;
  }

  return { value: visit(value), segments };
}

/**
 * @param {unknown} value
 * @param {Uint8Array} binBytes
 * @param {string} binFileName
 * @returns {unknown}
 */
function hydrateBinarySegments(value, binBytes, binFileName) {
  /**
   * @param {unknown} current
   * @returns {unknown}
   */
  function visit(current) {
    if (isBinaryMarker(current)) {
      // @ts-ignore - runtime indexing
      const markerFileName = current[BINARY_MARKER_KEY];
      if (markerFileName !== binFileName) {
        throw new Error("Invalid cache binary marker");
      }

      // Backwards compat: the original marker shape only included the filename
      // and implied the entire `.bin` file.
      const offset = typeof current.offset === "number" ? current.offset : 0;
      const length = typeof current.length === "number" ? current.length : binBytes.byteLength;

      if (
        !Number.isFinite(offset) ||
        !Number.isFinite(length) ||
        offset < 0 ||
        length < 0 ||
        offset + length > binBytes.byteLength
      ) {
        throw new Error("Invalid cache binary marker range");
      }

      return new Uint8Array(binBytes.buffer, binBytes.byteOffset + offset, length);
    }

    if (Array.isArray(current)) {
      return current.map((item) => visit(item));
    }

    if (current && typeof current === "object") {
      for (const [k, v] of Object.entries(current)) {
        // @ts-ignore - runtime indexing
        current[k] = visit(v);
      }
      return current;
    }

    return current;
  }

  return visit(value);
}

/**
 * Very small filesystem cache store for Node environments.
 *
 * It stores one JSON file per key, plus an optional `.bin` blob when the cached
 * value includes binary payloads (`Uint8Array`), such as Arrow IPC bytes.
 *
 * Keys are hashed to filenames to avoid filesystem character issues.
 */
export class FileSystemCacheStore {
  /**
   * @param {{ directory: string }} options
   */
  constructor(options) {
    this.directory = options.directory;
    /** @type {{ fs: typeof import("node:fs/promises"), path: typeof import("node:path") } | null} */
    this._deps = null;
  }

  /**
   * @returns {Promise<{ fs: typeof import("node:fs/promises"), path: typeof import("node:path") }>}
   */
  async deps() {
    if (this._deps) return this._deps;
    const fs = await import("node:fs/promises");
    const path = await import("node:path");
    this._deps = { fs, path };
    return this._deps;
  }

  /**
   * @param {string} key
   * @returns {Promise<{ jsonPath: string, binPath: string, binFileName: string }>}
   */
  async pathsForKey(key) {
    const hashed = fnv1a64(key);
    const { path } = await this.deps();
    return {
      jsonPath: path.join(this.directory, `${hashed}.json`),
      binPath: path.join(this.directory, `${hashed}.bin`),
      binFileName: `${hashed}.bin`,
    };
  }

  /**
   * Backwards-compatible helper for callers that relied on the original JSON-only
   * cache implementation.
   *
   * @param {string} key
   * @returns {Promise<string>}
   */
  async filePathForKey(key) {
    const { jsonPath } = await this.pathsForKey(key);
    return jsonPath;
  }

  async ensureDir() {
    const { fs } = await this.deps();
    await fs.mkdir(this.directory, { recursive: true });
  }

  /**
   * Write a file to disk in a best-effort atomic fashion (write to a temporary
   * file then rename).
   *
   * This helps avoid partially-written cache files if the process crashes or is
   * interrupted mid-write.
   *
   * @param {string} finalPath
   * @param {string | Uint8Array} data
   * @param {BufferEncoding | undefined} encoding
   */
  async writeFileAtomic(finalPath, data, encoding) {
    const { fs, path } = await this.deps();
    const tmpPath = path.join(
      path.dirname(finalPath),
      `${path.basename(finalPath)}.tmp-${Date.now()}-${Math.random().toString(16).slice(2)}`,
    );

    try {
      await fs.writeFile(tmpPath, data, encoding);

      try {
        await fs.rename(tmpPath, finalPath);
      } catch (err) {
        // On Windows, rename does not reliably overwrite existing files.
        if (err && typeof err === "object" && "code" in err && (err.code === "EEXIST" || err.code === "EPERM")) {
          await fs.rm(finalPath, { force: true });
          await fs.rename(tmpPath, finalPath);
          return;
        }
        throw err;
      }
    } catch (err) {
      // Best-effort cleanup if we fail mid-write.
      await fs.rm(tmpPath, { force: true }).catch(() => {});
      throw err;
    }
  }

  /**
   * @param {string} key
   * @returns {Promise<CacheEntry | null>}
   */
  async get(key) {
    await this.ensureDir();
    const { jsonPath, binPath, binFileName } = await this.pathsForKey(key);
    try {
      const { fs } = await this.deps();
      const text = await fs.readFile(jsonPath, "utf8");
      const parsed = JSON.parse(text);
      if (!parsed || parsed.key !== key) return null;

      const entry = parsed.entry ?? null;
      const value = entry?.value;

      if (containsBinaryMarker(value)) {
        const bytes = await fs.readFile(binPath);
        const restored = new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
        entry.value = hydrateBinarySegments(value, restored, binFileName);
      }

      return entry;
    } catch {
      return null;
    }
  }

  /**
   * @param {string} key
   * @param {CacheEntry} entry
   */
  async set(key, entry) {
    await this.ensureDir();
    const { fs } = await this.deps();
    const { jsonPath, binPath, binFileName } = await this.pathsForKey(key);

    const extracted = extractBinarySegments(entry?.value, binFileName);

    if (extracted.segments.length > 0) {
      const total = extracted.segments.reduce((sum, seg) => sum + seg.length, 0);
      const combined = new Uint8Array(total);
      for (const seg of extracted.segments) {
        combined.set(seg.bytes, seg.offset);
      }

      await this.writeFileAtomic(binPath, combined);
      await this.writeFileAtomic(jsonPath, JSON.stringify({ key, entry: { ...entry, value: extracted.value } }), "utf8");
      return;
    }

    // If we are writing a JSON-only entry, clean up any previous binary blob.
    await fs.rm(binPath, { force: true });
    await this.writeFileAtomic(jsonPath, JSON.stringify({ key, entry }), "utf8");
  }

  /**
   * @param {string} key
   */
  async delete(key) {
    const { fs } = await this.deps();
    const { jsonPath, binPath } = await this.pathsForKey(key);
    await fs.rm(jsonPath, { force: true });
    await fs.rm(binPath, { force: true });
  }

  async clear() {
    const { fs } = await this.deps();
    await fs.rm(this.directory, { recursive: true, force: true });
  }
}
