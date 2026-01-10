import initSqlJs from "sql.js";
import { createRequire } from "node:module";
import { promises as fs } from "node:fs";
import path from "node:path";
import { applyPatch } from "../patch.js";
import { randomUUID } from "../uuid.js";

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
   * @param {{ filePath?: string, filename?: string }} options
   */
  constructor(options) {
    const filePath = options.filePath ?? options.filename;
    if (!filePath) {
      throw new Error("SQLiteBranchStore requires { filePath }");
    }
    this.filePath = filePath;

    /** @type {any | null} */
    this._db = null;
    /** @type {Promise<any> | null} */
    this._initPromise = null;
    /** @type {Promise<void>} */
    this._persistChain = Promise.resolve();
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

    const db = existing ? new SQL.Database(existing) : new SQL.Database();
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
        patch_json TEXT NOT NULL
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
  }

  async _persist() {
    const db = await this._open();
    const data = db.export();
    const tmp = `${this.filePath}.tmp`;
    await fs.writeFile(tmp, data);
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

    const patch = { sheets: structuredClone(initialState.sheets ?? {}) };

    db.run("BEGIN");
    try {
      const insertCommit = db.prepare(
        `INSERT INTO commits
          (id, doc_id, parent_commit_id, merge_parent_commit_id, created_by, created_at, message, patch_json)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
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
   * @param {{ docId: string, parentCommitId: string | null, mergeParentCommitId: string | null, createdBy: string, createdAt: number, message: string | null, patch: Patch }} input
   * @returns {Promise<Commit>}
   */
  async createCommit(input) {
    const db = await this._open();
    const id = randomUUID();
    const stmt = db.prepare(
      `INSERT INTO commits
        (id, doc_id, parent_commit_id, merge_parent_commit_id, created_by, created_at, message, patch_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
    );
    stmt.run([
      id,
      input.docId,
      input.parentCommitId,
      input.mergeParentCommitId,
      input.createdBy,
      input.createdAt,
      input.message,
      JSON.stringify(input.patch),
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
    const commit = await this.getCommit(commitId);
    if (!commit) throw new Error(`Commit not found: ${commitId}`);

    /** @type {Commit[]} */
    const chain = [];
    let current = commit;
    while (current) {
      chain.push(current);
      if (!current.parentCommitId) break;
      const parent = await this.getCommit(current.parentCommitId);
      if (!parent) throw new Error(`Commit not found: ${current.parentCommitId}`);
      current = parent;
    }
    chain.reverse();

    /** @type {DocumentState} */
    let state = { sheets: {} };
    for (const c of chain) {
      state = applyPatch(state, c.patch);
    }
    return state;
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
