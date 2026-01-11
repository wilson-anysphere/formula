import initSqlJs from "sql.js";
import { createRequire } from "node:module";
import { promises as fs } from "node:fs";
import path from "node:path";
import { applyPatch, diffDocumentStates } from "../patch.js";
import { emptyDocumentState, normalizeDocumentState } from "../state.js";
import { randomUUID } from "../uuid.js";
import {
  decodeEncryptedFileBytes,
  encodeEncryptedFileBytes,
  isEncryptedFileBytes
} from "../../../../security/crypto/encryptedFile.js";

/**
 * SQLite-backed store.
 *
 * Uses the same approach as `packages/versioning/src/store/sqliteVersionStore.js`:
 * `sql.js` (WASM SQLite) + file persistence. This keeps the implementation
 * dependency-light and Node-compatible without requiring native modules.
 */

/**
 * @typedef {import("../types.js").Branch} Branch
 * @typedef {import("../types.js").Commit} Commit
 * @typedef {import("../types.js").DocumentState} DocumentState
 * @typedef {import("../patch.js").Patch} Patch
 */

export class SQLiteBranchStore {
  /**
   * @param {{
   *   filePath?: string,
   *   filename?: string,
   *   snapshotEveryNCommits?: number,
   *   snapshotWhenPatchExceedsBytes?: number,
   *   encryption?: (
   *     | { mode: "off" }
   *     | { mode: "keyring", keyRing: import("../../../../security/crypto/keyring.js").KeyRing, aadContext?: any }
   *   )
   * }} options
   */
  constructor(options) {
    const filePath = options.filePath ?? options.filename;
    if (!filePath) {
      throw new Error("SQLiteBranchStore requires { filePath }");
    }
    this.filePath = filePath;
    this.snapshotEveryNCommits =
      options.snapshotEveryNCommits == null ? 50 : options.snapshotEveryNCommits;
    this.snapshotWhenPatchExceedsBytes =
      options.snapshotWhenPatchExceedsBytes == null ? null : options.snapshotWhenPatchExceedsBytes;

    this.encryption = options.encryption ?? { mode: "off" };
    if (this.encryption.mode === "keyring" && !this.encryption.keyRing) {
      throw new Error("SQLiteBranchStore encryption.mode='keyring' requires { keyRing }");
    }

    /** @type {any | null} */
    this._db = null;
    /** @type {Promise<any> | null} */
    this._initPromise = null;
    /** @type {Promise<void>} */
    this._persistChain = Promise.resolve();
  }

  _aadContext() {
    return this.encryption.mode === "keyring"
      ? (this.encryption.aadContext ?? { scope: "formula.branches.sqlite", schemaVersion: 1 })
      : null;
  }

  /**
   * @returns {Promise<any>}
   */
  async _open() {
    if (this._db) return this._db;
    if (this._initPromise) return this._initPromise;
    this._initPromise = this._openInner();
    return this._initPromise;
  }

  async _openInner() {
    await fs.mkdir(path.dirname(this.filePath), { recursive: true });

    const SQL = await initSqlJs({ locateFile: locateSqlJsFile });

    /** @type {Uint8Array | null} */
    let existing = null;
    try {
      existing = await fs.readFile(this.filePath);
    } catch {
      existing = null;
    }

    let bytes = existing ? Buffer.from(existing) : null;
    if (bytes && isEncryptedFileBytes(bytes)) {
      if (this.encryption.mode !== "keyring") {
        throw new Error("Encrypted SQLiteBranchStore requires encryption.mode='keyring'");
      }
      const decoded = decodeEncryptedFileBytes(bytes);
      bytes = this.encryption.keyRing.decryptBytes(decoded, { aadContext: this._aadContext() });
    }

    const db = bytes ? new SQL.Database(bytes) : new SQL.Database();
    this._db = db;
    this._ensureSchema();
    return db;
  }

  _ensureSchema() {
    if (!this._db) return;
    this._db.run(`
      PRAGMA foreign_keys = ON;

      CREATE TABLE IF NOT EXISTS commits (
        id TEXT PRIMARY KEY,
        doc_id TEXT NOT NULL,
        parent_commit_id TEXT,
        merge_parent_commit_id TEXT,
        created_by TEXT NOT NULL,
        created_at INTEGER NOT NULL,
        message TEXT,
        patch_json TEXT NOT NULL,
        snapshot_json TEXT
      );
      CREATE INDEX IF NOT EXISTS idx_commits_doc ON commits(doc_id);

      CREATE TABLE IF NOT EXISTS branches (
        id TEXT PRIMARY KEY,
        doc_id TEXT NOT NULL,
        name TEXT NOT NULL,
        created_by TEXT NOT NULL,
        created_at INTEGER NOT NULL,
        description TEXT,
        head_commit_id TEXT NOT NULL,
        UNIQUE(doc_id, name)
      );
      CREATE INDEX IF NOT EXISTS idx_branches_doc ON branches(doc_id);
    `);

    // Schema migration: add snapshot_json column for existing stores.
    const info = this._db.exec("PRAGMA table_info(commits);");
    const cols = new Set();
    if (info[0]?.values) {
      for (const row of info[0].values) {
        cols.add(row[1]);
      }
    }
    if (!cols.has("snapshot_json")) {
      this._db.run("ALTER TABLE commits ADD COLUMN snapshot_json TEXT;");
    }
  }

  async _persist() {
    const db = await this._open();
    const data = Buffer.from(db.export());
    let out = data;

    if (this.encryption.mode === "keyring") {
      const encrypted = this.encryption.keyRing.encryptBytes(out, { aadContext: this._aadContext() });
      out = encodeEncryptedFileBytes({
        keyVersion: encrypted.keyVersion,
        iv: encrypted.iv,
        tag: encrypted.tag,
        ciphertext: encrypted.ciphertext
      });
    }

    const tmp = `${this.filePath}.tmp`;
    await fs.writeFile(tmp, out);
    await fs.rename(tmp, this.filePath);
  }

  async _queuePersist() {
    this._persistChain = this._persistChain.then(() => this._persist());
    return this._persistChain;
  }

  /**
   * @param {string} docId
   * @param {import("../types.js").Actor} actor
   * @param {DocumentState} initialState
   */
  async ensureDocument(docId, actor, initialState) {
    const db = await this._open();
    const existingStmt = db.prepare("SELECT id FROM branches WHERE doc_id = ? AND name = 'main' LIMIT 1");
    existingStmt.bind([docId]);
    const hasExisting = existingStmt.step();
    existingStmt.free();
    if (hasExisting) return;

    const now = Date.now();
    const rootCommitId = randomUUID();
    const mainBranchId = randomUUID();

    const patch = diffDocumentStates(emptyDocumentState(), initialState);
    const snapshot = applyPatch(emptyDocumentState(), patch);

    db.run("BEGIN");
    try {
      const insertCommit = db.prepare(
        `INSERT INTO commits
          (id, doc_id, parent_commit_id, merge_parent_commit_id, created_by, created_at, message, patch_json, snapshot_json)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
      );
      insertCommit.run([
        rootCommitId,
        docId,
        null,
        null,
        actor.userId,
        now,
        "root",
        JSON.stringify(patch),
        JSON.stringify(snapshot),
      ]);
      insertCommit.free();

      const insertBranch = db.prepare(
        `INSERT INTO branches
          (id, doc_id, name, created_by, created_at, description, head_commit_id)
          VALUES (?, ?, ?, ?, ?, ?, ?)`
      );
      insertBranch.run([mainBranchId, docId, "main", actor.userId, now, null, rootCommitId]);
      insertBranch.free();

      db.run("COMMIT");
    } catch (e) {
      db.run("ROLLBACK");
      throw e;
    }

    await this._queuePersist();
  }

  /**
   * @param {string} docId
   * @returns {Promise<Branch[]>}
   */
  async listBranches(docId) {
    const db = await this._open();
    const stmt = db.prepare(
      `SELECT id, doc_id, name, created_by, created_at, description, head_commit_id
       FROM branches WHERE doc_id = ? ORDER BY created_at ASC`
    );
    stmt.bind([docId]);

    /** @type {Branch[]} */
    const out = [];
    while (stmt.step()) {
      const row = stmt.getAsObject();
      out.push({
        id: row.id,
        docId: row.doc_id,
        name: row.name,
        createdBy: row.created_by,
        createdAt: row.created_at,
        description: row.description ?? null,
        headCommitId: row.head_commit_id,
      });
    }
    stmt.free();
    return out;
  }

  /**
   * @param {string} docId
   * @param {string} name
   * @returns {Promise<Branch | null>}
   */
  async getBranch(docId, name) {
    const db = await this._open();
    const stmt = db.prepare(
      `SELECT id, doc_id, name, created_by, created_at, description, head_commit_id
       FROM branches WHERE doc_id = ? AND name = ? LIMIT 1`
    );
    stmt.bind([docId, name]);
    if (!stmt.step()) {
      stmt.free();
      return null;
    }
    const row = stmt.getAsObject();
    stmt.free();
    return {
      id: row.id,
      docId: row.doc_id,
      name: row.name,
      createdBy: row.created_by,
      createdAt: row.created_at,
      description: row.description ?? null,
      headCommitId: row.head_commit_id,
    };
  }

  /**
   * @param {{ docId: string, name: string, createdBy: string, createdAt: number, description: string | null, headCommitId: string }} input
   * @returns {Promise<Branch>}
   */
  async createBranch(input) {
    const db = await this._open();
    const id = randomUUID();
    const stmt = db.prepare(
      `INSERT INTO branches
        (id, doc_id, name, created_by, created_at, description, head_commit_id)
        VALUES (?, ?, ?, ?, ?, ?, ?)`
    );
    stmt.run([
      id,
      input.docId,
      input.name,
      input.createdBy,
      input.createdAt,
      input.description,
      input.headCommitId,
    ]);
    stmt.free();
    await this._queuePersist();

    return {
      id,
      docId: input.docId,
      name: input.name,
      createdBy: input.createdBy,
      createdAt: input.createdAt,
      description: input.description,
      headCommitId: input.headCommitId,
    };
  }

  /**
   * @param {string} docId
   * @param {string} oldName
   * @param {string} newName
   */
  async renameBranch(docId, oldName, newName) {
    const db = await this._open();
    const stmt = db.prepare("UPDATE branches SET name = ? WHERE doc_id = ? AND name = ?");
    stmt.run([newName, docId, oldName]);
    stmt.free();
    await this._queuePersist();
  }

  /**
   * @param {string} docId
   * @param {string} name
   */
  async deleteBranch(docId, name) {
    const db = await this._open();
    const stmt = db.prepare("DELETE FROM branches WHERE doc_id = ? AND name = ?");
    stmt.run([docId, name]);
    stmt.free();
    await this._queuePersist();
  }

  /**
   * @param {string} docId
   * @param {string} name
   * @param {string} headCommitId
   */
  async updateBranchHead(docId, name, headCommitId) {
    const db = await this._open();
    const stmt = db.prepare("UPDATE branches SET head_commit_id = ? WHERE doc_id = ? AND name = ?");
    stmt.run([headCommitId, docId, name]);
    stmt.free();
    await this._queuePersist();
  }

  /**
   * @param {{ docId: string, parentCommitId: string | null, mergeParentCommitId: string | null, createdBy: string, createdAt: number, message: string | null, patch: Patch, nextState?: DocumentState }} input
   * @returns {Promise<Commit>}
   */
  async createCommit(input) {
    const db = await this._open();
    const id = randomUUID();

    const patchJson = JSON.stringify(input.patch);
    const shouldSnapshot = await this.#shouldSnapshotCommit({ parentCommitId: input.parentCommitId, patchJson });
    const snapshotState = shouldSnapshot
      ? await this.#resolveSnapshotState({ parentCommitId: input.parentCommitId, patch: input.patch, nextState: input.nextState })
      : null;
    const snapshotJson = snapshotState ? JSON.stringify(snapshotState) : null;
    const stmt = db.prepare(
      `INSERT INTO commits
        (id, doc_id, parent_commit_id, merge_parent_commit_id, created_by, created_at, message, patch_json, snapshot_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
    );
    stmt.run([
      id,
      input.docId,
      input.parentCommitId,
      input.mergeParentCommitId,
      input.createdBy,
      input.createdAt,
      input.message,
      patchJson,
      snapshotJson,
    ]);
    stmt.free();
    await this._queuePersist();

    return {
      id,
      docId: input.docId,
      parentCommitId: input.parentCommitId,
      mergeParentCommitId: input.mergeParentCommitId,
      createdBy: input.createdBy,
      createdAt: input.createdAt,
      message: input.message,
      patch: structuredClone(input.patch),
    };
  }

  /**
   * @param {string} commitId
   * @returns {Promise<Commit | null>}
   */
  async getCommit(commitId) {
    const db = await this._open();
    const stmt = db.prepare(
      `SELECT id, doc_id, parent_commit_id, merge_parent_commit_id,
        created_by, created_at, message, patch_json
      FROM commits WHERE id = ? LIMIT 1`
    );
    stmt.bind([commitId]);
    if (!stmt.step()) {
      stmt.free();
      return null;
    }
    const row = stmt.getAsObject();
    stmt.free();
    return {
      id: row.id,
      docId: row.doc_id,
      parentCommitId: row.parent_commit_id ?? null,
      mergeParentCommitId: row.merge_parent_commit_id ?? null,
      createdBy: row.created_by,
      createdAt: row.created_at,
      message: row.message ?? null,
      patch: JSON.parse(row.patch_json),
    };
  }

  /**
   * @param {string} commitId
   * @returns {Promise<DocumentState>}
   */
  async getDocumentStateAtCommit(commitId) {
    const row = await this.#getCommitForState(commitId);
    if (!row) throw new Error(`Commit not found: ${commitId}`);

    if (row.snapshotJson) {
      return normalizeDocumentState(JSON.parse(row.snapshotJson));
    }

    /** @type {{ id: string, patch: Patch }[]} */
    const chain = [];
    let current = row;
    while (current && !current.snapshotJson) {
      chain.push({ id: current.id, patch: JSON.parse(current.patchJson) });
      if (!current.parentCommitId) break;
      const parent = await this.#getCommitForState(current.parentCommitId);
      if (!parent) throw new Error(`Commit not found: ${current.parentCommitId}`);
      current = parent;
    }

    chain.reverse();

    /** @type {DocumentState} */
    let state = current?.snapshotJson ? normalizeDocumentState(JSON.parse(current.snapshotJson)) : emptyDocumentState();
    for (const c of chain) {
      state = this._applyPatch(state, c.patch);
    }
    return state;
  }

  /**
   * @param {DocumentState} state
   * @param {Patch} patch
   * @returns {DocumentState}
   */
  _applyPatch(state, patch) {
    return applyPatch(state, patch);
  }

  async #shouldSnapshotCommit({ parentCommitId, patchJson }) {
    if (this.snapshotWhenPatchExceedsBytes != null && this.snapshotWhenPatchExceedsBytes > 0) {
      const patchBytes = Buffer.byteLength(patchJson, "utf8");
      if (patchBytes > this.snapshotWhenPatchExceedsBytes) return true;
    }

    if (this.snapshotEveryNCommits != null && this.snapshotEveryNCommits > 0) {
      const distance = await this.#distanceFromSnapshotCommit(parentCommitId);
      if (distance + 1 >= this.snapshotEveryNCommits) return true;
    }

    return false;
  }

  async #distanceFromSnapshotCommit(startCommitId) {
    if (!startCommitId) return 0;
    let distance = 0;
    let currentId = startCommitId;
    while (currentId) {
      const row = await this.#getCommitForState(currentId);
      if (!row) throw new Error(`Commit not found: ${currentId}`);
      if (row.snapshotJson) return distance;
      if (!row.parentCommitId) return distance;
      distance += 1;
      currentId = row.parentCommitId;
    }
    return distance;
  }

  async #resolveSnapshotState({ parentCommitId, patch, nextState }) {
    if (nextState) return normalizeDocumentState(nextState);
    const base = parentCommitId ? await this.getDocumentStateAtCommit(parentCommitId) : emptyDocumentState();
    return this._applyPatch(base, patch);
  }

  /**
   * Fetches a lightweight commit row for state reconstruction / snapshot traversal.
   *
   * @param {string} commitId
   * @returns {Promise<{ id: string, parentCommitId: string | null, patchJson: string, snapshotJson: string | null } | null>}
   */
  async #getCommitForState(commitId) {
    const db = await this._open();
    const stmt = db.prepare(
      `SELECT id, parent_commit_id, patch_json, snapshot_json
       FROM commits WHERE id = ? LIMIT 1`
    );
    stmt.bind([commitId]);
    if (!stmt.step()) {
      stmt.free();
      return null;
    }
    const row = stmt.getAsObject();
    stmt.free();
    return {
      id: row.id,
      parentCommitId: row.parent_commit_id ?? null,
      patchJson: row.patch_json,
      snapshotJson: row.snapshot_json ?? null,
    };
  }

  close() {
    if (!this._db) return;
    this._db.close();
    this._db = null;
  }
}

function locateSqlJsFile(file) {
  try {
    const require = createRequire(import.meta.url);
    return require.resolve(`sql.js/dist/${file}`);
  } catch {
    return file;
  }
}
