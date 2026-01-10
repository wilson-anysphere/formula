import { promises as fs } from "node:fs";
import path from "node:path";

/**
 * @typedef {"snapshot" | "checkpoint" | "restore"} VersionKind
 *
 * @typedef {{
 *   id: string;
 *   kind: VersionKind;
 *   timestampMs: number;
 *   userId?: string | null;
 *   userName?: string | null;
 *   description?: string | null;
 *   checkpointName?: string | null;
 *   checkpointLocked?: boolean | null;
 *   checkpointAnnotations?: string | null;
 *   snapshotBase64: string;
 * }} StoredVersion
 */

/**
 * @typedef {{
 *   id: string;
 *   kind: VersionKind;
 *   timestampMs: number;
 *   userId: string | null;
 *   userName: string | null;
 *   description: string | null;
 *   checkpointName: string | null;
 *   checkpointLocked: boolean | null;
 *   checkpointAnnotations: string | null;
 *   snapshot: Uint8Array;
 * }} VersionRecord
 */

/**
 * A minimal persistence layer that writes versions to a JSON file.
 *
 * The production design in docs calls for SQLite (desktop) / IndexedDB (web),
 * but for this repo we keep it dependency-free and still storage-backed.
 */
export class FileVersionStore {
  /**
   * @param {{ filePath: string }} opts
   */
  constructor(opts) {
    this.filePath = opts.filePath;
  }

  async _ensureFile() {
    await fs.mkdir(path.dirname(this.filePath), { recursive: true });
    try {
      await fs.access(this.filePath);
    } catch {
      await fs.writeFile(this.filePath, JSON.stringify({ versions: [] }, null, 2), "utf8");
    }
  }

  /**
   * @returns {Promise<{ versions: StoredVersion[] }>}
   */
  async _read() {
    await this._ensureFile();
    const raw = await fs.readFile(this.filePath, "utf8");
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object" || !Array.isArray(parsed.versions)) {
      throw new Error(`Corrupt version store file: ${this.filePath}`);
    }
    return parsed;
  }

  /**
   * @param {{ versions: StoredVersion[] }} data
   */
  async _write(data) {
    const tmp = `${this.filePath}.tmp`;
    await fs.writeFile(tmp, JSON.stringify(data, null, 2), "utf8");
    await fs.rename(tmp, this.filePath);
  }

  /**
   * @param {VersionRecord} version
   */
  async saveVersion(version) {
    const data = await this._read();
    /** @type {StoredVersion} */
    const stored = {
      id: version.id,
      kind: version.kind,
      timestampMs: version.timestampMs,
      userId: version.userId ?? null,
      userName: version.userName ?? null,
      description: version.description ?? null,
      checkpointName: version.checkpointName ?? null,
      checkpointLocked: version.checkpointLocked ?? null,
      checkpointAnnotations: version.checkpointAnnotations ?? null,
      snapshotBase64: Buffer.from(version.snapshot).toString("base64"),
    };
    data.versions.push(stored);
    await this._write(data);
  }

  /**
   * @param {string} versionId
   * @returns {Promise<VersionRecord | null>}
   */
  async getVersion(versionId) {
    const data = await this._read();
    const found = data.versions.find((v) => v.id === versionId);
    if (!found) return null;
    return {
      id: found.id,
      kind: found.kind,
      timestampMs: found.timestampMs,
      userId: found.userId ?? null,
      userName: found.userName ?? null,
      description: found.description ?? null,
      checkpointName: found.checkpointName ?? null,
      checkpointLocked: found.checkpointLocked ?? null,
      checkpointAnnotations: found.checkpointAnnotations ?? null,
      snapshot: Buffer.from(found.snapshotBase64, "base64"),
    };
  }

  /**
   * @returns {Promise<VersionRecord[]>}
   */
  async listVersions() {
    const data = await this._read();
    const versions = data.versions.map((v) => ({
      id: v.id,
      kind: v.kind,
      timestampMs: v.timestampMs,
      userId: v.userId ?? null,
      userName: v.userName ?? null,
      description: v.description ?? null,
      checkpointName: v.checkpointName ?? null,
      checkpointLocked: v.checkpointLocked ?? null,
      checkpointAnnotations: v.checkpointAnnotations ?? null,
      snapshot: Buffer.from(v.snapshotBase64, "base64"),
    }));
    versions.sort((a, b) => b.timestampMs - a.timestampMs);
    return versions;
  }

  /**
   * @param {string} versionId
   * @param {{ checkpointLocked?: boolean }} patch
   * @returns {Promise<void>}
   */
  async updateVersion(versionId, patch) {
    const data = await this._read();
    const idx = data.versions.findIndex((v) => v.id === versionId);
    if (idx === -1) throw new Error(`Version not found: ${versionId}`);
    const existing = data.versions[idx];
    data.versions[idx] = {
      ...existing,
      checkpointLocked:
        patch.checkpointLocked === undefined ? existing.checkpointLocked : patch.checkpointLocked,
    };
    await this._write(data);
  }
}

