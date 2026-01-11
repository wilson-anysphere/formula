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
 * JSON.stringify replacer that extracts Uint8Array payloads into a sibling `.bin`
 * file and replaces them with marker objects.
 *
 * @param {{ key: string, entry: CacheEntry }} payload
 * @param {string} binFileName
 * @returns {{ jsonText: string, segments: Array<{ offset: number, length: number, bytes: Uint8Array }>, totalBytes: number }}
 */
function jsonWithBinarySegments(payload, binFileName) {
  /** @type {Array<{ offset: number, length: number, bytes: Uint8Array }>} */
  const segments = [];
  let offset = 0;

  const jsonText = JSON.stringify(payload, (_key, value) => {
    if (value instanceof Uint8Array) {
      const segmentOffset = offset;
      const length = value.byteLength;
      offset += length;
      segments.push({ offset: segmentOffset, length, bytes: value });
      return { [BINARY_MARKER_KEY]: binFileName, offset: segmentOffset, length };
    }

    // Node Buffers define a `toJSON()` hook that runs before replacers, so we may
    // see the `{ type: "Buffer", data: number[] }` shape here instead of the
    // original `Uint8Array`. Treat it as binary to avoid JSON bloat.
    if (
      value &&
      typeof value === "object" &&
      !Array.isArray(value) &&
      // @ts-ignore - runtime inspection
      value.type === "Buffer" &&
      // @ts-ignore - runtime inspection
      Array.isArray(value.data)
    ) {
      // @ts-ignore - runtime inspection
      const data = value.data;
      const segmentOffset = offset;
      const bytes = Uint8Array.from(data);
      const length = bytes.byteLength;
      offset += length;
      segments.push({ offset: segmentOffset, length, bytes });
      return { [BINARY_MARKER_KEY]: binFileName, offset: segmentOffset, length };
    }

    return value;
  });

  return { jsonText, segments, totalBytes: offset };
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
      const hasBinaryMarkers = text.includes(`"${BINARY_MARKER_KEY}":`);
      const parsed = JSON.parse(text);
      if (!parsed || parsed.key !== key) return null;

      const entry = parsed.entry ?? null;
      const value = entry?.value;

      if (hasBinaryMarkers) {
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

    const serialized = jsonWithBinarySegments({ key, entry }, binFileName);

    if (serialized.segments.length > 0) {
      const combined = new Uint8Array(serialized.totalBytes);
      for (const seg of serialized.segments) {
        combined.set(seg.bytes, seg.offset);
      }
      await this.writeFileAtomic(binPath, combined);
      await this.writeFileAtomic(jsonPath, serialized.jsonText, "utf8");
      return;
    }

    // If we are writing a JSON-only entry, clean up any previous binary blob.
    await fs.rm(binPath, { force: true });
    await this.writeFileAtomic(jsonPath, serialized.jsonText, "utf8");
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

  /**
   * Best-effort TTL pruning by scanning cache files on disk.
   *
   * @param {number} [nowMs]
   */
  async pruneExpired(nowMs = Date.now()) {
    await this.ensureDir();
    const { fs, path } = await this.deps();
    const tmpGraceMs = 5 * 60 * 1000;

    /** @type {import("node:fs").Dirent[]} */
    let entries = [];
    try {
      entries = await fs.readdir(this.directory, { withFileTypes: true });
    } catch {
      return;
    }

    for (const entry of entries) {
      if (!entry.isFile()) continue;

      // Best-effort cleanup of stale temp files left behind by interrupted writes.
      // We only remove temp files older than a small grace period to avoid racing
      // concurrent writers.
      if (entry.name.includes(".tmp-") || entry.name.endsWith(".tmp")) {
        const tmpPath = path.join(this.directory, entry.name);
        try {
          const stats = await fs.stat(tmpPath);
          const ageMs = nowMs - stats.mtimeMs;
          if (Number.isFinite(stats.mtimeMs) && ageMs > tmpGraceMs) {
            await fs.rm(tmpPath, { force: true });
          }
        } catch {
          // ignore
        }
        continue;
      }

      if (!entry.name.endsWith(".json")) continue;
      const jsonPath = path.join(this.directory, entry.name);
      const binPath = path.join(this.directory, `${entry.name.slice(0, -".json".length)}.bin`);

      try {
        const text = await fs.readFile(jsonPath, "utf8");
        const parsed = JSON.parse(text);
        const expiresAtMs =
          parsed && typeof parsed === "object" && parsed.entry && typeof parsed.entry === "object" ? parsed.entry.expiresAtMs : null;
        if (typeof expiresAtMs === "number" && expiresAtMs <= nowMs) {
          await fs.rm(jsonPath, { force: true });
          await fs.rm(binPath, { force: true });
        }
      } catch {
        // Corrupted or unreadable cache entries are treated as misses; remove them
        // so they don't linger indefinitely.
        await fs.rm(jsonPath, { force: true }).catch(() => {});
        await fs.rm(binPath, { force: true }).catch(() => {});
      }
    }
  }
}
