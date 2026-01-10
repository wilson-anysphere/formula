import { diffDocumentStates } from "./patch.js";
import { applyConflictResolutions, mergeDocumentStates } from "./merge.js";

/**
 * @typedef {import("./types.js").Actor} Actor
 * @typedef {import("./types.js").DocumentState} DocumentState
 * @typedef {import("./types.js").Branch} Branch
 * @typedef {import("./types.js").Commit} Commit
 * @typedef {import("./types.js").MergeResult} MergeResult
 * @typedef {import("./merge.js").ConflictResolution} ConflictResolution
 */

function assertCanManageBranches(actor) {
  if (actor.role !== "owner" && actor.role !== "admin") {
    throw new Error("Branch operations require owner/admin permissions");
  }
}

/**
 * BranchService provides high-level branch/merge operations for a single
 * document.
 *
 * It is intentionally UI-agnostic; the desktop panel wires user interactions to
 * these methods.
 */
export class BranchService {
  /** @type {string} */
  #docId;
  /** @type {any} */
  #store;
  /** @type {string} */
  #currentBranchName = "main";

  /**
   * @param {{ docId: string, store: any }} input
   */
  constructor({ docId, store }) {
    this.#docId = docId;
    this.#store = store;
  }

  /**
   * @param {Actor} actor
   * @param {DocumentState} initialState
   */
  async init(actor, initialState) {
    await this.#store.ensureDocument(this.#docId, actor, initialState);
  }

  async listBranches() {
    return this.#store.listBranches(this.#docId);
  }

  /**
   * @returns {Promise<Branch>}
   */
  async getCurrentBranch() {
    const branch = await this.#store.getBranch(this.#docId, this.#currentBranchName);
    if (!branch) throw new Error(`Current branch not found: ${this.#currentBranchName}`);
    return branch;
  }

  /**
   * @param {Actor} actor
   * @param {{ name: string, description?: string }} input
   */
  async createBranch(actor, { name, description }) {
    assertCanManageBranches(actor);
    const current = await this.getCurrentBranch();
    return this.#store.createBranch({
      docId: this.#docId,
      name,
      createdBy: actor.userId,
      createdAt: Date.now(),
      description: description ?? null,
      headCommitId: current.headCommitId
    });
  }

  /**
   * @param {Actor} actor
   * @param {{ oldName: string, newName: string }} input
   */
  async renameBranch(actor, { oldName, newName }) {
    assertCanManageBranches(actor);
    if (oldName === this.#currentBranchName) this.#currentBranchName = newName;
    await this.#store.renameBranch(this.#docId, oldName, newName);
  }

  /**
   * @param {Actor} actor
   * @param {{ name: string }} input
   */
  async deleteBranch(actor, { name }) {
    assertCanManageBranches(actor);
    if (name === "main") throw new Error("Cannot delete main branch");
    if (name === this.#currentBranchName) {
      throw new Error("Cannot delete the currently checked-out branch");
    }
    await this.#store.deleteBranch(this.#docId, name);
  }

  /**
   * Checks out the specified branch and returns its current state.
   *
   * @param {Actor} actor
   * @param {{ name: string }} input
   * @returns {Promise<DocumentState>}
   */
  async checkoutBranch(actor, { name }) {
    assertCanManageBranches(actor);
    const branch = await this.#store.getBranch(this.#docId, name);
    if (!branch) throw new Error(`Branch not found: ${name}`);
    this.#currentBranchName = name;
    return this.#store.getDocumentStateAtCommit(branch.headCommitId);
  }

  /**
   * Creates a new commit on the current branch and advances the branch head.
   *
   * @param {Actor} actor
   * @param {{ nextState: DocumentState, message?: string }} input
   */
  async commit(actor, { nextState, message }) {
    const branch = await this.getCurrentBranch();
    const currentState = await this.#store.getDocumentStateAtCommit(branch.headCommitId);
    const patch = diffDocumentStates(currentState, nextState);
    const commit = await this.#store.createCommit({
      docId: this.#docId,
      parentCommitId: branch.headCommitId,
      mergeParentCommitId: null,
      createdBy: actor.userId,
      createdAt: Date.now(),
      message: message ?? null,
      patch
    });
    await this.#store.updateBranchHead(this.#docId, branch.name, commit.id);
    return commit;
  }

  /**
   * @param {Actor} actor
   * @param {{ sourceBranch: string }} input
   * @returns {Promise<MergeResult & { baseCommitId: string, oursHeadCommitId: string, theirsHeadCommitId: string }>}
   */
  async previewMerge(actor, { sourceBranch }) {
    assertCanManageBranches(actor);
    const oursBranch = await this.getCurrentBranch();
    const theirsBranch = await this.#store.getBranch(this.#docId, sourceBranch);
    if (!theirsBranch) throw new Error(`Branch not found: ${sourceBranch}`);

    const baseCommitId = await this.#findCommonAncestorCommitId(
      oursBranch.headCommitId,
      theirsBranch.headCommitId
    );

    const [baseState, oursState, theirsState] = await Promise.all([
      this.#store.getDocumentStateAtCommit(baseCommitId),
      this.#store.getDocumentStateAtCommit(oursBranch.headCommitId),
      this.#store.getDocumentStateAtCommit(theirsBranch.headCommitId)
    ]);

    const merge = mergeDocumentStates({ base: baseState, ours: oursState, theirs: theirsState });
    return {
      ...merge,
      baseCommitId,
      oursHeadCommitId: oursBranch.headCommitId,
      theirsHeadCommitId: theirsBranch.headCommitId
    };
  }

  /**
   * Applies a merge (with conflict resolutions) and creates a new merge commit on
   * the current branch.
   *
   * @param {Actor} actor
   * @param {{ sourceBranch: string, resolutions: ConflictResolution[], message?: string }} input
   * @returns {Promise<{ commit: Commit, state: DocumentState }>}
   */
  async merge(actor, { sourceBranch, resolutions, message }) {
    const preview = await this.previewMerge(actor, { sourceBranch });
    const finalState = applyConflictResolutions(preview, resolutions);

    // Ensure all conflicts were addressed by validating that each conflict index
    // is present in `resolutions`.
    const resolved = new Set(resolutions.map((r) => r.conflictIndex));
    for (let i = 0; i < preview.conflicts.length; i += 1) {
      if (!resolved.has(i)) throw new Error("All merge conflicts must be resolved before merging");
    }

    const oursState = await this.#store.getDocumentStateAtCommit(preview.oursHeadCommitId);
    const patch = diffDocumentStates(oursState, finalState);

    const commit = await this.#store.createCommit({
      docId: this.#docId,
      parentCommitId: preview.oursHeadCommitId,
      mergeParentCommitId: preview.theirsHeadCommitId,
      createdBy: actor.userId,
      createdAt: Date.now(),
      message: message ?? `Merge branch '${sourceBranch}'`,
      patch
    });

    const currentBranch = await this.getCurrentBranch();
    await this.#store.updateBranchHead(this.#docId, currentBranch.name, commit.id);

    return { commit, state: finalState };
  }

  /**
   * @param {string} oursHead
   * @param {string} theirsHead
   * @returns {Promise<string>}
   */
  async #findCommonAncestorCommitId(oursHead, theirsHead) {
    /** @type {Map<string, number>} */
    const oursDepth = new Map();
    const queue = [{ id: oursHead, depth: 0 }];
    while (queue.length > 0) {
      const { id, depth } = queue.shift();
      if (oursDepth.has(id)) continue;
      oursDepth.set(id, depth);
      const commit = await this.#store.getCommit(id);
      if (!commit) throw new Error(`Commit not found: ${id}`);
      if (commit.parentCommitId) queue.push({ id: commit.parentCommitId, depth: depth + 1 });
      if (commit.mergeParentCommitId)
        queue.push({ id: commit.mergeParentCommitId, depth: depth + 1 });
    }

    let best = null;
    let bestScore = Infinity;
    const queue2 = [{ id: theirsHead, depth: 0 }];
    /** @type {Set<string>} */
    const seen = new Set();
    while (queue2.length > 0) {
      const { id, depth } = queue2.shift();
      if (seen.has(id)) continue;
      seen.add(id);
      const ours = oursDepth.get(id);
      if (ours !== undefined) {
        const score = ours + depth;
        if (score < bestScore) {
          best = id;
          bestScore = score;
        }
      }
      const commit = await this.#store.getCommit(id);
      if (!commit) throw new Error(`Commit not found: ${id}`);
      if (commit.parentCommitId) queue2.push({ id: commit.parentCommitId, depth: depth + 1 });
      if (commit.mergeParentCommitId)
        queue2.push({ id: commit.mergeParentCommitId, depth: depth + 1 });
    }

    if (!best) throw new Error("No common ancestor found (corrupt history)");
    return best;
  }
}

