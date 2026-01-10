import crypto from "node:crypto";
import { DatabaseSync } from "node:sqlite";
import { applyPatch } from "../patch.js";

/**
 * SQLite-backed store.
 *
 * Notes:
 * - Uses Node's built-in (currently experimental) `node:sqlite` module so the
 *   repo can run tests without native dependencies.
 * - Stores commits as patches (JSON) relative to the first parent to avoid
 *   duplicating full document snapshots on every version.
 */

/**
 * @typedef {import("../types.js").Branch} Branch
 * @typedef {import("../types.js").Commit} Commit
 * @typedef {import("../types.js").DocumentState} DocumentState
 * @typedef {import("../patch.js").Patch} Patch
 */

export class SQLiteBranchStore {
  /** @type {DatabaseSync} */
  #db;

  /**
   * @param {{ filename: string }} options
   */
  constructor({ filename }) {
    this.#db = new DatabaseSync(filename);
    this.#migrate();
  }

  #migrate() {
    this.#db.exec(`
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

  /**
   * @param {string} docId
   * @param {import("../types.js").Actor} actor
   * @param {DocumentState} initialState
   */
  async ensureDocument(docId, actor, initialState) {
    const existing = this.#db
      .prepare("SELECT id FROM branches WHERE doc_id = ? AND name = 'main'")
      .get(docId);
    if (existing) return;

    const now = Date.now();
    const rootCommitId = crypto.randomUUID();
    const mainBranchId = crypto.randomUUID();

    const patch = { sheets: structuredClone(initialState.sheets ?? {}) };

    this.#db.exec("BEGIN");
    try {
      this.#db
        .prepare(
          `INSERT INTO commits
          (id, doc_id, parent_commit_id, merge_parent_commit_id, created_by, created_at, message, patch_json)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
        )
        .run(
          rootCommitId,
          docId,
          null,
          null,
          actor.userId,
          now,
          "root",
          JSON.stringify(patch)
        );

      this.#db
        .prepare(
          `INSERT INTO branches
          (id, doc_id, name, created_by, created_at, description, head_commit_id)
          VALUES (?, ?, ?, ?, ?, ?, ?)`
        )
        .run(mainBranchId, docId, "main", actor.userId, now, null, rootCommitId);

      this.#db.exec("COMMIT");
    } catch (e) {
      this.#db.exec("ROLLBACK");
      throw e;
    }
  }

  /**
   * @param {string} docId
   * @returns {Promise<Branch[]>}
   */
  async listBranches(docId) {
    const rows = this.#db
      .prepare(
        `SELECT id, doc_id as docId, name, created_by as createdBy, created_at as createdAt,
          description, head_commit_id as headCommitId
        FROM branches WHERE doc_id = ? ORDER BY created_at ASC`
      )
      .all(docId);
    return rows.map((r) => ({
      ...r,
      description: r.description ?? null
    }));
  }

  /**
   * @param {string} docId
   * @param {string} name
   * @returns {Promise<Branch | null>}
   */
  async getBranch(docId, name) {
    const row = this.#db
      .prepare(
        `SELECT id, doc_id as docId, name, created_by as createdBy, created_at as createdAt,
          description, head_commit_id as headCommitId
        FROM branches WHERE doc_id = ? AND name = ?`
      )
      .get(docId, name);
    if (!row) return null;
    return { ...row, description: row.description ?? null };
  }

  /**
   * @param {{ docId: string, name: string, createdBy: string, createdAt: number, description: string | null, headCommitId: string }} input
   * @returns {Promise<Branch>}
   */
  async createBranch(input) {
    const id = crypto.randomUUID();
    this.#db
      .prepare(
        `INSERT INTO branches
        (id, doc_id, name, created_by, created_at, description, head_commit_id)
        VALUES (?, ?, ?, ?, ?, ?, ?)`
      )
      .run(
        id,
        input.docId,
        input.name,
        input.createdBy,
        input.createdAt,
        input.description,
        input.headCommitId
      );

    return {
      id,
      docId: input.docId,
      name: input.name,
      createdBy: input.createdBy,
      createdAt: input.createdAt,
      description: input.description,
      headCommitId: input.headCommitId
    };
  }

  /**
   * @param {string} docId
   * @param {string} oldName
   * @param {string} newName
   */
  async renameBranch(docId, oldName, newName) {
    this.#db
      .prepare("UPDATE branches SET name = ? WHERE doc_id = ? AND name = ?")
      .run(newName, docId, oldName);
  }

  /**
   * @param {string} docId
   * @param {string} name
   */
  async deleteBranch(docId, name) {
    this.#db
      .prepare("DELETE FROM branches WHERE doc_id = ? AND name = ?")
      .run(docId, name);
  }

  /**
   * @param {string} docId
   * @param {string} name
   * @param {string} headCommitId
   */
  async updateBranchHead(docId, name, headCommitId) {
    this.#db
      .prepare("UPDATE branches SET head_commit_id = ? WHERE doc_id = ? AND name = ?")
      .run(headCommitId, docId, name);
  }

  /**
   * @param {{ docId: string, parentCommitId: string | null, mergeParentCommitId: string | null, createdBy: string, createdAt: number, message: string | null, patch: Patch }} input
   * @returns {Promise<Commit>}
   */
  async createCommit(input) {
    const id = crypto.randomUUID();
    this.#db
      .prepare(
        `INSERT INTO commits
        (id, doc_id, parent_commit_id, merge_parent_commit_id, created_by, created_at, message, patch_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      )
      .run(
        id,
        input.docId,
        input.parentCommitId,
        input.mergeParentCommitId,
        input.createdBy,
        input.createdAt,
        input.message,
        JSON.stringify(input.patch)
      );

    return {
      id,
      docId: input.docId,
      parentCommitId: input.parentCommitId,
      mergeParentCommitId: input.mergeParentCommitId,
      createdBy: input.createdBy,
      createdAt: input.createdAt,
      message: input.message,
      patch: structuredClone(input.patch)
    };
  }

  /**
   * @param {string} commitId
   * @returns {Promise<Commit | null>}
   */
  async getCommit(commitId) {
    const row = this.#db
      .prepare(
        `SELECT id, doc_id as docId, parent_commit_id as parentCommitId,
          merge_parent_commit_id as mergeParentCommitId, created_by as createdBy,
          created_at as createdAt, message, patch_json as patchJson
        FROM commits WHERE id = ?`
      )
      .get(commitId);
    if (!row) return null;
    return {
      id: row.id,
      docId: row.docId,
      parentCommitId: row.parentCommitId ?? null,
      mergeParentCommitId: row.mergeParentCommitId ?? null,
      createdBy: row.createdBy,
      createdAt: row.createdAt,
      message: row.message ?? null,
      patch: JSON.parse(row.patchJson)
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
}

