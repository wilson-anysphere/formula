import { diffDocumentStates } from "./patch.js";
import { applyConflictResolutions, mergeDocumentStates } from "./merge.js";
import { normalizeDocumentState } from "./state.js";

/**
 * @typedef {import("./types.js").Actor} Actor
 * @typedef {import("./types.js").DocumentState} DocumentState
 * @typedef {import("./types.js").Branch} Branch
 * @typedef {import("./types.js").Commit} Commit
 * @typedef {import("./types.js").MergeResult} MergeResult
 * @typedef {import("./merge.js").ConflictResolution} ConflictResolution
 */

/**
 * Branch mutations (create/rename/delete/checkout/merge) are restricted to
 * document owners/admins. Editors can still create commits (see
 * `assertCanCommit`).
 *
 * @param {Actor} actor
 * @param {string} operation
 */
function assertCanManageBranches(actor, operation) {
  if (actor.role !== "owner" && actor.role !== "admin") {
    throw new Error(`${operation} requires owner/admin permissions (role=${actor.role})`);
  }
}

/**
 * Commits mutate document history but are allowed for any role that can edit the
 * document contents.
 *
 * @param {Actor} actor
 */
function assertCanCommit(actor) {
  if (actor.role !== "owner" && actor.role !== "admin" && actor.role !== "editor") {
    throw new Error(`Commit requires edit permission (role=${actor.role})`);
  }
}

/**
 * @param {any} value
 * @returns {value is Record<string, any>}
 */
function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

/**
 * Backwards-compatible detection for the BranchService v0 state shape:
 * `{ sheets: Record<sheetId, CellMap> }`.
 *
 * Old clients only know about cell edits; they should not be able to accidentally
 * wipe workbook metadata (sheet names/order, namedRanges, comments) by omitting
 * those fields from commits.
 *
 * @param {any} value
 */
function isLegacyCellsOnlyState(value) {
  return (
    isRecord(value) &&
    value.schemaVersion !== 1 &&
    !("cells" in value) &&
    isRecord(value.sheets)
  );
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

  async #getCurrentBranchName() {
    const store = this.#store;
    if (store && typeof store.getCurrentBranchName === "function") {
      const name = await store.getCurrentBranchName(this.#docId);
      if (typeof name === "string" && name.length > 0) {
        // Keep the local cache in sync even when the store is authoritative.
        this.#currentBranchName = name;
        return name;
      }
    }
    return this.#currentBranchName;
  }

  /**
   * @param {string} name
   */
  async #setCurrentBranchName(name) {
    const store = this.#store;
    if (store && typeof store.setCurrentBranchName === "function") {
      await store.setCurrentBranchName(this.#docId, name);
      this.#currentBranchName = name;
      return;
    }
    this.#currentBranchName = name;
  }

  /**
   * @param {Actor} actor
   * @param {DocumentState} initialState
   */
  async init(actor, initialState) {
    // `ensureDocument` will create the root commit + main branch if the document
    // doesn't exist yet. Creating a new document is an admin-level action, but
    // calling `init` on an existing document is safe for any role (it becomes a
    // no-op in the store).
    const existingMain = await this.#store.getBranch(this.#docId, "main");
    if (!existingMain) {
      assertCanManageBranches(actor, "init");
    }
    await this.#store.ensureDocument(this.#docId, actor, initialState);
  }

  async listBranches() {
    return this.#store.listBranches(this.#docId);
  }

  /**
   * @returns {Promise<Branch>}
   */
  async getCurrentBranch() {
    const name = await this.#getCurrentBranchName();
    const branch = await this.#store.getBranch(this.#docId, name);
    if (!branch) throw new Error(`Current branch not found: ${name}`);
    return branch;
  }

  /**
   * Convenience helper: load the current branch head state.
   *
   * This is intentionally *not* permission-gated (it's a read of already
   * reachable history) and is useful for adapters that can't represent the full
   * workbook metadata surface area (e.g. DocumentController).
   *
   * @returns {Promise<DocumentState>}
   */
  async getCurrentState() {
    const branch = await this.getCurrentBranch();
    return this.#store.getDocumentStateAtCommit(branch.headCommitId);
  }

  /**
   * @param {Actor} actor
   * @param {{ name: string, description?: string }} input
   */
  async createBranch(actor, { name, description }) {
    assertCanManageBranches(actor, "createBranch");
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
    assertCanManageBranches(actor, "renameBranch");
    await this.#store.renameBranch(this.#docId, oldName, newName);
    const current = await this.#getCurrentBranchName();
    if (oldName === current) await this.#setCurrentBranchName(newName);
  }

  /**
   * @param {Actor} actor
   * @param {{ name: string }} input
   */
  async deleteBranch(actor, { name }) {
    assertCanManageBranches(actor, "deleteBranch");
    if (name === "main") throw new Error("Cannot delete main branch");
    const current = await this.#getCurrentBranchName();
    if (name === current) {
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
    assertCanManageBranches(actor, "checkoutBranch");
    const branch = await this.#store.getBranch(this.#docId, name);
    if (!branch) throw new Error(`Branch not found: ${name}`);
    await this.#setCurrentBranchName(name);
    return this.#store.getDocumentStateAtCommit(branch.headCommitId);
  }

  /**
   * Creates a new commit on the current branch and advances the branch head.
   *
   * @param {Actor} actor
   * @param {{ nextState: DocumentState, message?: string }} input
   */
  async commit(actor, { nextState, message }) {
    assertCanCommit(actor);
    const branch = await this.getCurrentBranch();
    const currentState = normalizeDocumentState(
      await this.#store.getDocumentStateAtCommit(branch.headCommitId)
    );

    let effectiveNextState = nextState;
    if (isLegacyCellsOnlyState(nextState)) {
      const legacy = normalizeDocumentState(nextState);

      // Legacy commits only provide per-sheet cell maps. Treat them as an overlay
      // on the current branch head so older callers cannot accidentally delete
      // workbook metadata or unrelated sheets.
      const merged = structuredClone(currentState);
      for (const [sheetId, cellMap] of Object.entries(legacy.cells ?? {})) {
        merged.cells[sheetId] = structuredClone(cellMap);
        if (!merged.sheets.metaById[sheetId]) {
          merged.sheets.metaById[sheetId] = { id: sheetId, name: sheetId };
        }
        if (!merged.sheets.order.includes(sheetId)) merged.sheets.order.push(sheetId);
      }
      effectiveNextState = merged;
    }

    const patch = diffDocumentStates(currentState, effectiveNextState);
    const commit = await this.#store.createCommit({
      docId: this.#docId,
      parentCommitId: branch.headCommitId,
      mergeParentCommitId: null,
      createdBy: actor.userId,
      createdAt: Date.now(),
      message: message ?? null,
      patch,
      nextState: effectiveNextState,
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
    assertCanManageBranches(actor, "previewMerge");
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
    assertCanManageBranches(actor, "merge");
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
      patch,
      nextState: finalState,
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
