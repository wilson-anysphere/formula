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

  /**
   * @param {{ session: CollabSession, branchService: BranchService }} input
   */
  constructor({ session, branchService }) {
    this.#session = session;
    this.#branchService = branchService;
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

