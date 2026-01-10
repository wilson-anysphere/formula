import crypto from "node:crypto";
import { applyPatch } from "../patch.js";

/**
 * @typedef {import("../types.js").Branch} Branch
 * @typedef {import("../types.js").Commit} Commit
 * @typedef {import("../types.js").DocumentState} DocumentState
 * @typedef {import("../types.js").Actor} Actor
 * @typedef {import("../patch.js").Patch} Patch
 */

/**
 * @typedef {{
 *   ensureDocument(docId: string, actor: Actor, initialState: DocumentState): Promise<void>,
 *   listBranches(docId: string): Promise<Branch[]>,
 *   getBranch(docId: string, name: string): Promise<Branch | null>,
 *   createBranch(input: { docId: string, name: string, createdBy: string, createdAt: number, description: string | null, headCommitId: string }): Promise<Branch>,
 *   renameBranch(docId: string, oldName: string, newName: string): Promise<void>,
 *   deleteBranch(docId: string, name: string): Promise<void>,
 *   updateBranchHead(docId: string, name: string, headCommitId: string): Promise<void>,
 *   createCommit(input: { docId: string, parentCommitId: string | null, mergeParentCommitId: string | null, createdBy: string, createdAt: number, message: string | null, patch: Patch }): Promise<Commit>,
 *   getCommit(commitId: string): Promise<Commit | null>,
 *   getDocumentStateAtCommit(commitId: string): Promise<DocumentState>
 * }} BranchStore
 */

/**
 * In-memory store used for tests and as a reference implementation. It stores
 * commits as patches relative to the first parent; this keeps branch operations
 * lightweight and avoids duplicating full document snapshots.
 */
export class InMemoryBranchStore {
  /** @type {Map<string, Branch[]>} */
  #branchesByDoc = new Map();

  /** @type {Map<string, Commit>} */
  #commits = new Map();

  /** @type {Map<string, string>} */
  #rootCommitByDoc = new Map();

  async ensureDocument(docId, actor, initialState) {
    if (this.#rootCommitByDoc.has(docId)) return;

    const rootCommitId = crypto.randomUUID();
    const rootCommit = {
      id: rootCommitId,
      docId,
      parentCommitId: null,
      mergeParentCommitId: null,
      createdBy: actor.userId,
      createdAt: Date.now(),
      message: "root",
      patch: { sheets: structuredClone(initialState.sheets ?? {}) }
    };

    this.#commits.set(rootCommitId, rootCommit);
    this.#rootCommitByDoc.set(docId, rootCommitId);

    /** @type {Branch} */
    const main = {
      id: crypto.randomUUID(),
      docId,
      name: "main",
      createdBy: actor.userId,
      createdAt: Date.now(),
      description: null,
      headCommitId: rootCommitId
    };

    this.#branchesByDoc.set(docId, [main]);
  }

  async listBranches(docId) {
    return structuredClone(this.#branchesByDoc.get(docId) ?? []);
  }

  async getBranch(docId, name) {
    const branches = this.#branchesByDoc.get(docId) ?? [];
    return branches.find((b) => b.name === name) ?? null;
  }

  async createBranch({ docId, name, createdBy, createdAt, description, headCommitId }) {
    const branches = this.#branchesByDoc.get(docId) ?? [];
    if (branches.some((b) => b.name === name)) {
      throw new Error(`Branch already exists: ${name}`);
    }

    /** @type {Branch} */
    const branch = {
      id: crypto.randomUUID(),
      docId,
      name,
      createdBy,
      createdAt,
      description,
      headCommitId
    };

    branches.push(branch);
    this.#branchesByDoc.set(docId, branches);
    return structuredClone(branch);
  }

  async renameBranch(docId, oldName, newName) {
    const branches = this.#branchesByDoc.get(docId) ?? [];
    if (branches.some((b) => b.name === newName)) {
      throw new Error(`Branch already exists: ${newName}`);
    }
    const branch = branches.find((b) => b.name === oldName);
    if (!branch) throw new Error(`Branch not found: ${oldName}`);
    branch.name = newName;
  }

  async deleteBranch(docId, name) {
    const branches = this.#branchesByDoc.get(docId) ?? [];
    const filtered = branches.filter((b) => b.name !== name);
    this.#branchesByDoc.set(docId, filtered);
  }

  async updateBranchHead(docId, name, headCommitId) {
    const branches = this.#branchesByDoc.get(docId) ?? [];
    const branch = branches.find((b) => b.name === name);
    if (!branch) throw new Error(`Branch not found: ${name}`);
    branch.headCommitId = headCommitId;
  }

  async createCommit({ docId, parentCommitId, mergeParentCommitId, createdBy, createdAt, message, patch }) {
    const id = crypto.randomUUID();
    /** @type {Commit} */
    const commit = {
      id,
      docId,
      parentCommitId,
      mergeParentCommitId,
      createdBy,
      createdAt,
      message,
      patch: structuredClone(patch)
    };
    this.#commits.set(id, commit);
    return structuredClone(commit);
  }

  async getCommit(commitId) {
    return structuredClone(this.#commits.get(commitId) ?? null);
  }

  async getDocumentStateAtCommit(commitId) {
    const commit = this.#commits.get(commitId);
    if (!commit) throw new Error(`Commit not found: ${commitId}`);

    /** @type {Commit[]} */
    const chain = [];
    let current = commit;
    while (current) {
      chain.push(current);
      if (!current.parentCommitId) break;
      const parent = this.#commits.get(current.parentCommitId);
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

