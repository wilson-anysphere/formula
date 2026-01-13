import {
  applyDocumentStateToYjsDoc,
  yjsDocToDocumentState,
} from "../../versioning/branches/src/browser.js";
import * as Y from "yjs";

function getYMap(value) {
  if (value instanceof Y.Map) return value;
  if (!value || typeof value !== "object") return null;
  const maybe = value;
  if (typeof maybe.get !== "function") return null;
  if (typeof maybe.set !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  if (typeof maybe.keys !== "function") return null;
  if (typeof maybe.forEach !== "function") return null;
  if (typeof maybe.observeDeep !== "function") return null;
  if (typeof maybe.unobserveDeep !== "function") return null;
  return maybe;
}

function isYAbstractType(value) {
  if (value instanceof Y.AbstractType) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = value;
  if (typeof maybe.observeDeep !== "function") return false;
  if (typeof maybe.unobserveDeep !== "function") return false;
  return Boolean(maybe._map instanceof Map || maybe._start || maybe._item || maybe._length != null);
}

function replaceForeignRootType({ doc, name, existing, create }) {
  const t = create();
  t._map = existing?._map;
  t._start = existing?._start;
  t._length = existing?._length;

  const map = existing?._map;
  if (map instanceof Map) {
    map.forEach((item) => {
      for (let n = item; n !== null; n = n.left) {
        n.parent = t;
      }
    });
  }

  for (let n = existing?._start ?? null; n !== null; n = n.right) {
    n.parent = t;
  }

  doc.share.set(name, t);
  t._integrate?.(doc, null);
  return t;
}

function getMapRoot(doc, name) {
  const existing = doc.share.get(name);
  if (!existing) return doc.getMap(name);

  const map = getYMap(existing);
  if (map) {
    if (map instanceof Y.Map) return map;
    if (doc instanceof Y.Doc) {
      return replaceForeignRootType({ doc, name, existing: map, create: () => new Y.Map() });
    }
    return map;
  }

  if (isYAbstractType(existing) && doc instanceof Y.Doc) {
    return replaceForeignRootType({ doc, name, existing, create: () => new Y.Map() });
  }
  if (isYAbstractType(existing)) return doc.getMap(name);

  return doc.getMap(name);
}

/**
 * @typedef {import("@formula/collab-session").CollabSession} CollabSession
 * @typedef {import("../../versioning/branches/src/types.js").Actor} Actor
 * @typedef {import("../../versioning/branches/src/BranchService.js").BranchService} BranchService
 * @typedef {import("../../versioning/branches/src/merge.js").ConflictResolution} ConflictResolution
 */

export class CollabBranchingWorkflow {
  /** @type {CollabSession} */
  #session;
  /** @type {BranchService} */
  #branchService;
  /** @type {string} */
  #rootName;
  /** @type {any} */
  #applyOrigin;

  /**
   * Origin configuration for checkout/merge:
   * - default uses a dedicated origin token (`"branching-apply"`) so bulk rewrites aren't
   *   captured by collaborative undo tracking (`session.origin`). CollabSession also treats
   *   this origin as an ignored "time travel" operation for conflict monitors.
   * - pass `applyWithSessionOrigin: true` (or `applyOrigin: session.origin`) to opt back
   *   into the historical undoable behavior.
   *
   * @param {{
   *   session: CollabSession,
   *   branchService: BranchService,
   *   rootName?: string,
   *   applyOrigin?: any,
   *   applyWithSessionOrigin?: boolean,
   * }} input
   */
  constructor({ session, branchService, rootName, applyOrigin, applyWithSessionOrigin }) {
    this.#session = session;
    this.#branchService = branchService;
    this.#rootName = rootName ?? "branching";
    this.#applyOrigin = applyWithSessionOrigin
      ? this.#session.origin
      : applyOrigin !== undefined
        ? applyOrigin
        : "branching-apply";
  }

  /**
   * Returns the globally checked-out branch name stored in Yjs metadata.
   *
   * @returns {string}
   */
  getCurrentBranchName() {
    const name = this.#getGlobalCurrentBranchName();
    const branches = getMapRoot(this.#session.doc, `${this.#rootName}:branches`);
    return branches.get(name) !== undefined ? name : "main";
  }

  /**
   * Async variant backed by the underlying BranchService/store.
   *
   * Prefer this when you want store-level validation/self-healing (e.g.
   * YjsBranchStore will normalize dangling pointers).
   *
   * @returns {Promise<string>}
   */
  async getCurrentBranchNameAsync() {
    return this.#branchService.getCurrentBranchName();
  }

  /**
   * @returns {Promise<ReturnType<BranchService["listBranches"]>>}
   */
  async listBranches() {
    return this.#branchService.listBranches();
  }

  /**
   * @param {Actor} actor
   * @param {{ name: string, description?: string }} input
   */
  async createBranch(actor, input) {
    return this.#branchService.createBranch(actor, input);
  }

  /**
   * @param {Actor} actor
   * @param {{ oldName: string, newName: string }} input
   */
  async renameBranch(actor, { oldName, newName }) {
    await this.#branchService.renameBranch(actor, { oldName, newName });
  }

  /**
   * @param {Actor} actor
   * @param {{ name: string }} input
   */
  async deleteBranch(actor, { name }) {
    await this.#branchService.deleteBranch(actor, { name });
  }

  async getCurrentBranch() {
    return this.#branchService.getCurrentBranch();
  }

  /**
   * Returns the state of the globally checked-out branch head.
   */
  async getCurrentState() {
    return this.#branchService.getCurrentState();
  }

  /**
   * @returns {string}
   */
  #getGlobalCurrentBranchName() {
    const meta = getMapRoot(this.#session.doc, `${this.#rootName}:meta`);
    const name = meta.get("currentBranchName");
    return typeof name === "string" && name.length > 0 ? name : "main";
  }

  /**
   * Snapshot the current collaborative workbook state into a new commit.
   *
   * @param {Actor} actor
   * @param {string} [message]
   */
  async commitCurrentState(actor, message) {
    const nextState = yjsDocToDocumentState(this.#session.doc);
    return this.#branchService.commit(actor, { nextState, message });
  }

  /**
   * @param {Actor} actor
   * @param {{ name: string }} input
   */
  async checkoutBranch(actor, { name }) {
    const state = await this.#branchService.checkoutBranch(actor, { name });
    applyDocumentStateToYjsDoc(this.#session.doc, state, { origin: this.#applyOrigin });
    return state;
  }

  /**
   * @param {Actor} actor
   * @param {{ sourceBranch: string }} input
   */
  async previewMerge(actor, { sourceBranch }) {
    return this.#branchService.previewMerge(actor, { sourceBranch });
  }

  /**
   * @param {Actor} actor
   * @param {{ sourceBranch: string, resolutions: ConflictResolution[], message?: string }} input
   */
  async merge(actor, { sourceBranch, resolutions, message }) {
    const result = await this.#branchService.merge(actor, { sourceBranch, resolutions, message });
    applyDocumentStateToYjsDoc(this.#session.doc, result.state, { origin: this.#applyOrigin });
    return result;
  }
}
