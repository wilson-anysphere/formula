import { fnv1a64 } from "./key.js";

/**
 * @typedef {import("./cache.js").CacheEntry} CacheEntry
 */

/**
 * Very small filesystem cache store for Node environments.
 *
 * It stores one JSON file per key. Keys are hashed to filenames to avoid
 * filesystem character issues.
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
   * @returns {Promise<string>}
   */
  async filePathForKey(key) {
    const hashed = fnv1a64(key);
    const { path } = await this.deps();
    return path.join(this.directory, `${hashed}.json`);
  }

  async ensureDir() {
    const { fs } = await this.deps();
    await fs.mkdir(this.directory, { recursive: true });
  }

  /**
   * @param {string} key
   * @returns {Promise<CacheEntry | null>}
   */
  async get(key) {
    await this.ensureDir();
    const filePath = await this.filePathForKey(key);
    try {
      const { fs } = await this.deps();
      const text = await fs.readFile(filePath, "utf8");
      const parsed = JSON.parse(text);
      if (!parsed || parsed.key !== key) return null;
      return parsed.entry ?? null;
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
    const filePath = await this.filePathForKey(key);
    await fs.writeFile(filePath, JSON.stringify({ key, entry }), "utf8");
  }

  /**
   * @param {string} key
   */
  async delete(key) {
    const { fs } = await this.deps();
    const filePath = await this.filePathForKey(key);
    await fs.rm(filePath, { force: true });
  }

  async clear() {
    const { fs } = await this.deps();
    await fs.rm(this.directory, { recursive: true, force: true });
  }
}
