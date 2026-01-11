import { applyPatch, diffDocumentStates } from "../patch.js";
import { emptyDocumentState, normalizeDocumentState } from "../state.js";
import { randomUUID } from "../uuid.js";

const UTF8_ENCODER = new TextEncoder();

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
 *   hasDocument?(docId: string): Promise<boolean>,
 *   listBranches(docId: string): Promise<Branch[]>,
 *   getBranch(docId: string, name: string): Promise<Branch | null>,
 *   createBranch(input: { docId: string, name: string, createdBy: string, createdAt: number, description: string | null, headCommitId: string }): Promise<Branch>,
 *   renameBranch(docId: string, oldName: string, newName: string): Promise<void>,
 *   deleteBranch(docId: string, name: string): Promise<void>,
 *   updateBranchHead(docId: string, name: string, headCommitId: string): Promise<void>,
 *   getCurrentBranchName?(docId: string): Promise<string>,
 *   setCurrentBranchName?(docId: string, name: string): Promise<void>,
 *   createCommit(input: { docId: string, parentCommitId: string | null, mergeParentCommitId: string | null, createdBy: string, createdAt: number, message: string | null, patch: Patch, nextState?: DocumentState }): Promise<Commit>,
 *   getCommit(commitId: string): Promise<Commit | null>,
 *   getDocumentStateAtCommit(commitId: string): Promise<DocumentState>
 * }} BranchStore
 */

/**
 * In-memory store used for tests and as a reference implementation. It stores
 * commits as patches relative to the first parent, with periodic full-state
 * snapshots to keep checkouts from becoming O(history length).
 */
export class InMemoryBranchStore {
  /**
   * @param {{
   *   snapshotEveryNCommits?: number,
   *   snapshotWhenPatchExceedsBytes?: number
   * }=} options
   */
  constructor(options = {}) {
    this.snapshotEveryNCommits =
      options.snapshotEveryNCommits == null ? 50 : options.snapshotEveryNCommits;
    this.snapshotWhenPatchExceedsBytes =
      options.snapshotWhenPatchExceedsBytes == null ? null : options.snapshotWhenPatchExceedsBytes;
  }

  /** @type {Map<string, Branch[]>} */
  #branchesByDoc = new Map();

  /** @type {Map<string, Commit>} */
  #commits = new Map();

  /** @type {Map<string, DocumentState>} */
  #snapshotsByCommitId = new Map();

  /** @type {Map<string, string>} */
  #rootCommitByDoc = new Map();

  async ensureDocument(docId, actor, initialState) {
    if (this.#rootCommitByDoc.has(docId)) return;

    const rootCommitId = randomUUID();
    const patch = diffDocumentStates(emptyDocumentState(), initialState);
    const snapshot = applyPatch(emptyDocumentState(), patch);
    const rootCommit = {
      id: rootCommitId,
      docId,
      parentCommitId: null,
      mergeParentCommitId: null,
      createdBy: actor.userId,
      createdAt: Date.now(),
      message: "root",
      patch
    };

    this.#commits.set(rootCommitId, rootCommit);
    this.#snapshotsByCommitId.set(rootCommitId, snapshot);
    this.#rootCommitByDoc.set(docId, rootCommitId);

    /** @type {Branch} */
    const main = {
      id: randomUUID(),
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
      id: randomUUID(),
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

  async createCommit({
    docId,
    parentCommitId,
    mergeParentCommitId,
    createdBy,
    createdAt,
    message,
    patch,
    nextState,
  }) {
    const id = randomUUID();

    const shouldSnapshot = await this.#shouldSnapshotCommit({ parentCommitId, patch });
    const snapshot = shouldSnapshot ? await this.#resolveSnapshotState({ parentCommitId, patch, nextState }) : null;
    /** @type {Commit} */
    const commit = {
      id,
      docId,
      parentCommitId,
      mergeParentCommitId,
      createdBy,
      createdAt,
      message,
      patch: structuredClone(patch),
    };
    this.#commits.set(id, commit);
    if (snapshot) this.#snapshotsByCommitId.set(id, snapshot);
    return structuredClone(commit);
  }

  async getCommit(commitId) {
    return structuredClone(this.#commits.get(commitId) ?? null);
  }

  async getDocumentStateAtCommit(commitId) {
    const commit = this.#commits.get(commitId);
    if (!commit) throw new Error(`Commit not found: ${commitId}`);

    const directSnapshot = this.#snapshotsByCommitId.get(commitId);
    if (directSnapshot) return structuredClone(directSnapshot);

    /** @type {Commit[]} */
    const chain = [];
    let current = commit;
    while (current && !this.#snapshotsByCommitId.has(current.id)) {
      chain.push(current);
      if (!current.parentCommitId) break;
      const parent = this.#commits.get(current.parentCommitId);
      if (!parent) throw new Error(`Commit not found: ${current.parentCommitId}`);
      current = parent;
    }

    chain.reverse();

    /** @type {DocumentState} */
    const baseSnapshot = current ? this.#snapshotsByCommitId.get(current.id) : null;
    let state = baseSnapshot ? structuredClone(baseSnapshot) : emptyDocumentState();
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

  async #shouldSnapshotCommit({ parentCommitId, patch }) {
    if (this.snapshotWhenPatchExceedsBytes != null && this.snapshotWhenPatchExceedsBytes > 0) {
      const patchBytes = UTF8_ENCODER.encode(JSON.stringify(patch)).length;
      if (patchBytes > this.snapshotWhenPatchExceedsBytes) return true;
    }

    if (this.snapshotEveryNCommits != null && this.snapshotEveryNCommits > 0) {
      const distance = this.#distanceFromSnapshotCommit(parentCommitId);
      if (distance + 1 >= this.snapshotEveryNCommits) return true;
    }

    return false;
  }

  #distanceFromSnapshotCommit(startCommitId) {
    if (!startCommitId) return 0;
    let distance = 0;
    let currentId = startCommitId;
    while (currentId) {
      const commit = this.#commits.get(currentId);
      if (!commit) throw new Error(`Commit not found: ${currentId}`);
      if (this.#snapshotsByCommitId.has(commit.id)) return distance;
      if (!commit.parentCommitId) return distance;
      distance += 1;
      currentId = commit.parentCommitId;
    }
    return distance;
  }

  async #resolveSnapshotState({ parentCommitId, patch, nextState }) {
    if (nextState) return normalizeDocumentState(nextState);
    const base = parentCommitId ? await this.getDocumentStateAtCommit(parentCommitId) : emptyDocumentState();
    return this._applyPatch(base, patch);
  }
}
