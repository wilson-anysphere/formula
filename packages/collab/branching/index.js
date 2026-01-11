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
   * @returns {string}
   */
  #getGlobalCurrentBranchName() {
    const meta = this.#session.doc.getMap(`${this.#rootName}:meta`);
    const name = meta.get("currentBranchName");
    return typeof name === "string" && name.length > 0 ? name : "main";
  }

  /**
   * @param {string} name
   */
  #setGlobalCurrentBranchName(name) {
    const meta = this.#session.doc.getMap(`${this.#rootName}:meta`);
    this.#session.doc.transact(() => {
      meta.set("currentBranchName", name);
    }, this.#session.origin);
  }

  #syncBranchServiceToGlobalBranch() {
    const name = this.#getGlobalCurrentBranchName();
    if (this.#branchService.getCurrentBranchName?.() !== name) {
      this.#branchService.setCurrentBranchName?.(name);
    }
  }

  /**
   * Snapshot the current collaborative workbook state into a new commit.
   *
   * @param {Actor} actor
   * @param {string} [message]
   */
  async commitCurrentState(actor, message) {
    this.#syncBranchServiceToGlobalBranch();
    const nextState = yjsDocToDocumentState(this.#session.doc);
    return this.#branchService.commit(actor, { nextState, message });
  }

  /**
   * @param {Actor} actor
   * @param {{ name: string }} input
   */
  async checkoutBranch(actor, { name }) {
    const state = await this.#branchService.checkoutBranch(actor, { name });
    this.#setGlobalCurrentBranchName(name);
    applyDocumentStateToYjsDoc(this.#session.doc, state, { origin: this.#session.origin });
    return state;
  }

  /**
   * @param {Actor} actor
   * @param {{ sourceBranch: string }} input
   */
  async previewMerge(actor, { sourceBranch }) {
    this.#syncBranchServiceToGlobalBranch();
    return this.#branchService.previewMerge(actor, { sourceBranch });
  }

  /**
   * @param {Actor} actor
   * @param {{ sourceBranch: string, resolutions: ConflictResolution[], message?: string }} input
   */
  async merge(actor, { sourceBranch, resolutions, message }) {
    this.#syncBranchServiceToGlobalBranch();
    const result = await this.#branchService.merge(actor, { sourceBranch, resolutions, message });
    applyDocumentStateToYjsDoc(this.#session.doc, result.state, { origin: this.#session.origin });
    return result;
  }
}
