import {
  applyDocumentStateToYjsDoc,
  yjsDocToDocumentState,
} from "../../versioning/branches/src/yjs/index.js";

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

  /**
   * @param {{ session: CollabSession, branchService: BranchService, rootName?: string }} input
   */
  constructor({ session, branchService, rootName }) {
    this.#session = session;
    this.#branchService = branchService;
    this.#rootName = rootName ?? "branching";
  }

  /**
   * Returns the globally checked-out branch name stored in Yjs metadata.
   *
   * @returns {string}
   */
  getCurrentBranchName() {
    const name = this.#getGlobalCurrentBranchName();
    const branches = this.#session.doc.getMap(`${this.#rootName}:branches`);
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
    const meta = this.#session.doc.getMap(`${this.#rootName}:meta`);
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
    applyDocumentStateToYjsDoc(this.#session.doc, state, { origin: this.#session.origin });
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
    applyDocumentStateToYjsDoc(this.#session.doc, result.state, { origin: this.#session.origin });
    return result;
  }
}
