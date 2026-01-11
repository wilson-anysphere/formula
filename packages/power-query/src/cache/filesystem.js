import { fnv1a64 } from "./key.js";

/**
 * @typedef {import("./cache.js").CacheEntry} CacheEntry
 */

const BINARY_MARKER_KEY = "__pq_cache_binary";

/**
 * Very small filesystem cache store for Node environments.
 *
 * It stores one JSON file per key, plus an optional `.bin` blob when the cached
 * value includes Arrow IPC bytes.
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
      const bytesMarker = value?.version === 2 && value?.table?.kind === "arrow" ? value?.table?.bytes : null;
      if (
        bytesMarker &&
        typeof bytesMarker === "object" &&
        !Array.isArray(bytesMarker) &&
        typeof bytesMarker[BINARY_MARKER_KEY] === "string"
      ) {
        const markerFileName = bytesMarker[BINARY_MARKER_KEY];
        // Only allow loading from within the cache directory.
        if (markerFileName !== binFileName) return null;
        const bytes = await fs.readFile(binPath);
        const restored = new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
        entry.value = { ...value, table: { ...value.table, bytes: restored } };
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

    const value = entry?.value;
    const arrowBytes =
      value?.version === 2 && value?.table?.kind === "arrow" && value?.table?.bytes instanceof Uint8Array
        ? value.table.bytes
        : null;

    if (arrowBytes) {
      await this.writeFileAtomic(binPath, arrowBytes);
      const patchedValue = {
        ...value,
        table: {
          ...value.table,
          bytes: { [BINARY_MARKER_KEY]: binFileName },
        },
      };
      await this.writeFileAtomic(
        jsonPath,
        JSON.stringify({ key, entry: { ...entry, value: patchedValue } }),
        "utf8",
      );
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
