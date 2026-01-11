import * as Y from "yjs";
import { applyPatch, diffDocumentStates } from "../patch.js";
import { emptyDocumentState, normalizeDocumentState } from "../state.js";
import { randomUUID } from "../uuid.js";

/**
 * @typedef {import("../types.js").Actor} Actor
 * @typedef {import("../types.js").Branch} Branch
 * @typedef {import("../types.js").Commit} Commit
 * @typedef {import("../types.js").DocumentState} DocumentState
 * @typedef {import("../patch.js").Patch} Patch
 */

const UTF8_ENCODER = new TextEncoder();

/**
 * @param {unknown} value
 * @returns {Y.Map<any> | null}
 */
function getYMap(value) {
  if (value instanceof Y.Map) return value;

  // See CollabSession#getYMapCell for why we can't rely solely on instanceof.
  if (!value || typeof value !== "object") return null;
  const maybe = /** @type {any} */ (value);
  if (maybe.constructor?.name !== "YMap") return null;
  if (typeof maybe.get !== "function") return null;
  if (typeof maybe.set !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  return /** @type {Y.Map<any>} */ (maybe);
}

/**
 * Yjs-backed implementation of the BranchStore interface.
 *
 * Stores the branch + commit graph inside a shared Y.Doc so history syncs and
 * persists automatically via the collaboration layer.
 */
export class YjsBranchStore {
  /** @type {Y.Doc} */
  #ydoc;
  /** @type {string} */
  #rootName;
  /** @type {Y.Map<any>} */
  #branches;
  /** @type {Y.Map<any>} */
  #commits;
  /** @type {Y.Map<any>} */
  #meta;
  /** @type {number | null} */
  #snapshotEveryNCommits;
  /** @type {number | null} */
  #snapshotWhenPatchExceedsBytes;

  /**
   * @param {{
   *   ydoc: Y.Doc,
   *   rootName?: string,
   *   snapshotEveryNCommits?: number,
   *   snapshotWhenPatchExceedsBytes?: number
   * }} input
   */
  constructor({ ydoc, rootName, snapshotEveryNCommits, snapshotWhenPatchExceedsBytes }) {
    if (!ydoc) throw new Error("YjsBranchStore requires { ydoc }");
    this.#ydoc = ydoc;
    this.#rootName = rootName ?? "branching";
    this.#branches = ydoc.getMap(`${this.#rootName}:branches`);
    this.#commits = ydoc.getMap(`${this.#rootName}:commits`);
    this.#meta = ydoc.getMap(`${this.#rootName}:meta`);
    this.#snapshotEveryNCommits =
      snapshotEveryNCommits == null ? 50 : snapshotEveryNCommits;
    this.#snapshotWhenPatchExceedsBytes =
      snapshotWhenPatchExceedsBytes == null ? null : snapshotWhenPatchExceedsBytes;
  }

  /**
   * @param {string} docId
   * @param {Actor} actor
   * @param {DocumentState} initialState
   */
  async ensureDocument(docId, actor, initialState) {
    const existingRoot = this.#meta.get("rootCommitId");
    if (typeof existingRoot === "string" && existingRoot.length > 0) {
      // Backwards-compatible migration: older docs may have been created before
      // we started persisting the global checked-out branch name.
      const existingBranch = this.#meta.get("currentBranchName");
      if (typeof existingBranch !== "string" || existingBranch.length === 0) {
        this.#ydoc.transact(() => {
          const current = this.#meta.get("currentBranchName");
          if (typeof current === "string" && current.length > 0) return;
          this.#meta.set("currentBranchName", "main");
        });
      }

      // Migration: older docs may not have stored commit snapshots.
      const rootCommitMap = getYMap(this.#commits.get(existingRoot));
      if (rootCommitMap && rootCommitMap.get("snapshot") === undefined) {
        const patch = rootCommitMap.get("patch");
        if (patch) {
          const snapshot = this._applyPatch(emptyDocumentState(), patch);
          this.#ydoc.transact(() => {
            const commit = getYMap(this.#commits.get(existingRoot));
            if (!commit) return;
            if (commit.get("snapshot") !== undefined) return;
            commit.set("snapshot", structuredClone(snapshot));
          });
        }
      }

      return;
    }

    const now = Date.now();
    const rootCommitId = randomUUID();
    const mainBranchId = randomUUID();

    /** @type {Patch} */
    const patch = diffDocumentStates(emptyDocumentState(), initialState);
    const snapshot = this._applyPatch(emptyDocumentState(), patch);

    this.#ydoc.transact(() => {
      const rootAfter = this.#meta.get("rootCommitId");
      if (typeof rootAfter === "string" && rootAfter.length > 0) return;

      const commit = new Y.Map();
      commit.set("id", rootCommitId);
      commit.set("docId", docId);
      commit.set("parentCommitId", null);
      commit.set("mergeParentCommitId", null);
      commit.set("createdBy", actor.userId);
      commit.set("createdAt", now);
      commit.set("message", "root");
      commit.set("patch", structuredClone(patch));
      commit.set("snapshot", structuredClone(snapshot));
      this.#commits.set(rootCommitId, commit);

      const main = new Y.Map();
      main.set("id", mainBranchId);
      main.set("docId", docId);
      main.set("name", "main");
      main.set("createdBy", actor.userId);
      main.set("createdAt", now);
      main.set("description", null);
      main.set("headCommitId", rootCommitId);
      this.#branches.set("main", main);

      this.#meta.set("rootCommitId", rootCommitId);
      this.#meta.set("currentBranchName", "main");
    });
  }

  async getCurrentBranchName(_docId) {
    const name = this.#meta.get("currentBranchName");
    return typeof name === "string" && name.length > 0 ? name : "main";
  }

  /**
   * @param {string} _docId
   * @param {string} name
   */
  async setCurrentBranchName(_docId, name) {
    this.#ydoc.transact(() => {
      this.#meta.set("currentBranchName", name);
    });
  }

  /**
   * @param {Y.Map<any>} branchMap
   * @returns {Branch}
   */
  #branchFromYMap(branchMap) {
    return {
      id: String(branchMap.get("id") ?? ""),
      docId: String(branchMap.get("docId") ?? ""),
      name: String(branchMap.get("name") ?? ""),
      createdBy: String(branchMap.get("createdBy") ?? ""),
      createdAt: Number(branchMap.get("createdAt") ?? 0),
      description: (branchMap.get("description") ?? null) === null ? null : String(branchMap.get("description")),
      headCommitId: String(branchMap.get("headCommitId") ?? "")
    };
  }

  /**
   * @param {Y.Map<any>} commitMap
   * @returns {Commit}
   */
  #commitFromYMap(commitMap) {
    return {
      id: String(commitMap.get("id") ?? ""),
      docId: String(commitMap.get("docId") ?? ""),
      parentCommitId: commitMap.get("parentCommitId") ?? null,
      mergeParentCommitId: commitMap.get("mergeParentCommitId") ?? null,
      createdBy: String(commitMap.get("createdBy") ?? ""),
      createdAt: Number(commitMap.get("createdAt") ?? 0),
      message: commitMap.get("message") ?? null,
      patch: structuredClone(commitMap.get("patch") ?? { schemaVersion: 1 })
    };
  }

  /**
   * @param {string} docId
   * @returns {Promise<Branch[]>}
   */
  async listBranches(docId) {
    /** @type {Branch[]} */
    const out = [];
    this.#branches.forEach((value) => {
      const branchMap = getYMap(value);
      if (!branchMap) return;
      const branch = this.#branchFromYMap(branchMap);
      if (branch.docId !== docId) return;
      out.push(branch);
    });
    out.sort((a, b) => (a.createdAt - b.createdAt === 0 ? a.name.localeCompare(b.name) : a.createdAt - b.createdAt));
    return structuredClone(out);
  }

  /**
   * @param {string} docId
   * @param {string} name
   * @returns {Promise<Branch | null>}
   */
  async getBranch(docId, name) {
    const branchMap = getYMap(this.#branches.get(name));
    if (!branchMap) return null;
    const branch = this.#branchFromYMap(branchMap);
    if (branch.docId !== docId) return null;
    return structuredClone(branch);
  }

  /**
   * @param {{ docId: string, name: string, createdBy: string, createdAt: number, description: string | null, headCommitId: string }} input
   * @returns {Promise<Branch>}
   */
  async createBranch({ docId, name, createdBy, createdAt, description, headCommitId }) {
    if (this.#branches.has(name)) {
      throw new Error(`Branch already exists: ${name}`);
    }

    const id = randomUUID();
    this.#ydoc.transact(() => {
      if (this.#branches.has(name)) {
        throw new Error(`Branch already exists: ${name}`);
      }
      const branch = new Y.Map();
      branch.set("id", id);
      branch.set("docId", docId);
      branch.set("name", name);
      branch.set("createdBy", createdBy);
      branch.set("createdAt", createdAt);
      branch.set("description", description ?? null);
      branch.set("headCommitId", headCommitId);
      this.#branches.set(name, branch);
    });

    return {
      id,
      docId,
      name,
      createdBy,
      createdAt,
      description: description ?? null,
      headCommitId
    };
  }

  /**
   * @param {string} docId
   * @param {string} oldName
   * @param {string} newName
   */
  async renameBranch(docId, oldName, newName) {
    this.#ydoc.transact(() => {
      if (this.#branches.has(newName)) {
        throw new Error(`Branch already exists: ${newName}`);
      }

      const branchMap = getYMap(this.#branches.get(oldName));
      if (!branchMap) throw new Error(`Branch not found: ${oldName}`);
      if (String(branchMap.get("docId") ?? "") !== docId) {
        throw new Error(`Branch not found: ${oldName}`);
      }

      const next = new Y.Map();
      branchMap.forEach((v, k) => {
        if (k === "name") return;
        next.set(k, v);
      });
      next.set("name", newName);

      this.#branches.delete(oldName);
      this.#branches.set(newName, next);
    });
  }

  /**
   * @param {string} docId
   * @param {string} name
   */
  async deleteBranch(docId, name) {
    this.#ydoc.transact(() => {
      const branchMap = getYMap(this.#branches.get(name));
      if (!branchMap) return;
      if (String(branchMap.get("docId") ?? "") !== docId) return;
      this.#branches.delete(name);
    });
  }

  /**
   * @param {string} docId
   * @param {string} name
   * @param {string} headCommitId
   */
  async updateBranchHead(docId, name, headCommitId) {
    this.#ydoc.transact(() => {
      const branchMap = getYMap(this.#branches.get(name));
      if (!branchMap) throw new Error(`Branch not found: ${name}`);
      if (String(branchMap.get("docId") ?? "") !== docId) {
        throw new Error(`Branch not found: ${name}`);
      }
      branchMap.set("headCommitId", headCommitId);
    });
  }

  /**
   * @param {{ docId: string, parentCommitId: string | null, mergeParentCommitId: string | null, createdBy: string, createdAt: number, message: string | null, patch: Patch, nextState?: DocumentState }} input
   * @returns {Promise<Commit>}
   */
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
    const snapshotState = shouldSnapshot
      ? await this.#resolveSnapshotState({ parentCommitId, patch, nextState })
      : null;

    this.#ydoc.transact(() => {
      const commit = new Y.Map();
      commit.set("id", id);
      commit.set("docId", docId);
      commit.set("parentCommitId", parentCommitId);
      commit.set("mergeParentCommitId", mergeParentCommitId);
      commit.set("createdBy", createdBy);
      commit.set("createdAt", createdAt);
      commit.set("message", message ?? null);
      commit.set("patch", structuredClone(patch));
      if (snapshotState) commit.set("snapshot", structuredClone(snapshotState));
      this.#commits.set(id, commit);
    });

    return {
      id,
      docId,
      parentCommitId,
      mergeParentCommitId,
      createdBy,
      createdAt,
      message: message ?? null,
      patch: structuredClone(patch)
    };
  }

  /**
   * @param {string} commitId
   * @returns {Promise<Commit | null>}
   */
  async getCommit(commitId) {
    const commitMap = getYMap(this.#commits.get(commitId));
    if (!commitMap) return null;
    return structuredClone(this.#commitFromYMap(commitMap));
  }

  /**
   * @param {string} commitId
   * @returns {Promise<DocumentState>}
   */
  async getDocumentStateAtCommit(commitId) {
    const direct = getYMap(this.#commits.get(commitId));
    if (!direct) throw new Error(`Commit not found: ${commitId}`);

    const directSnapshot = direct.get("snapshot");
    if (directSnapshot !== undefined) {
      return normalizeDocumentState(structuredClone(directSnapshot));
    }

    /** @type {Patch[]} */
    const chain = [];
    let currentId = commitId;
    /** @type {any} */
    let baseSnapshot = null;

    while (currentId) {
      const commitMap = getYMap(this.#commits.get(currentId));
      if (!commitMap) throw new Error(`Commit not found: ${currentId}`);

      const snapshot = commitMap.get("snapshot");
      if (snapshot !== undefined) {
        baseSnapshot = snapshot;
        break;
      }

      chain.push(structuredClone(commitMap.get("patch") ?? {}));

      const parent = commitMap.get("parentCommitId");
      if (!parent) break;
      currentId = parent;
    }

    chain.reverse();

    /** @type {DocumentState} */
    let state = baseSnapshot !== null ? normalizeDocumentState(structuredClone(baseSnapshot)) : emptyDocumentState();
    for (const patch of chain) {
      state = this._applyPatch(state, patch);
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
    if (this.#snapshotWhenPatchExceedsBytes != null && this.#snapshotWhenPatchExceedsBytes > 0) {
      const patchBytes = UTF8_ENCODER.encode(JSON.stringify(patch)).length;
      if (patchBytes > this.#snapshotWhenPatchExceedsBytes) return true;
    }

    if (this.#snapshotEveryNCommits != null && this.#snapshotEveryNCommits > 0) {
      const distance = this.#distanceFromSnapshotCommit(parentCommitId);
      if (distance + 1 >= this.#snapshotEveryNCommits) return true;
    }

    return false;
  }

  #distanceFromSnapshotCommit(startCommitId) {
    if (!startCommitId) return 0;
    let distance = 0;
    let currentId = startCommitId;
    while (currentId) {
      const commitMap = getYMap(this.#commits.get(currentId));
      if (!commitMap) throw new Error(`Commit not found: ${currentId}`);
      if (commitMap.get("snapshot") !== undefined) return distance;
      const parentId = commitMap.get("parentCommitId");
      if (!parentId) return distance;
      distance += 1;
      currentId = parentId;
    }
    return distance;
  }

  async #resolveSnapshotState({ parentCommitId, patch, nextState }) {
    if (nextState) return normalizeDocumentState(nextState);
    const base = parentCommitId ? await this.getDocumentStateAtCommit(parentCommitId) : emptyDocumentState();
    return this._applyPatch(base, patch);
  }
}
