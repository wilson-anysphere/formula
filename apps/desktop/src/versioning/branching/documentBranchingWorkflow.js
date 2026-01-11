import {
  applyBranchStateToDocumentController,
  documentControllerToBranchState,
} from "./branchStateAdapter.js";

/**
 * @typedef {import("../../document/documentController.js").DocumentController} DocumentController
 * @typedef {import("../../../../../packages/versioning/branches/src/BranchService.js").BranchService} BranchService
 * @typedef {import("../../../../../packages/versioning/branches/src/types.js").Actor} Actor
 * @typedef {import("../../../../../packages/versioning/branches/src/merge.js").ConflictResolution} ConflictResolution
 */

/**
 * Small helper that wires the git-like BranchService workflow into a live
 * DocumentController by applying branch states on checkout/merge.
 */
export class DocumentBranchingWorkflow {
  /**
   * @param {{ doc: DocumentController, branchService: BranchService }} input
   */
  constructor({ doc, branchService }) {
    this.doc = doc;
    this.branchService = branchService;
  }

  /**
   * @param {Actor} actor
   * @param {string} [message]
   */
  async commitCurrentState(actor, message) {
    const nextState = documentControllerToBranchState(this.doc);
    return this.branchService.commit(actor, { nextState, message });
  }

  /**
   * @param {Actor} actor
   * @param {string} branchName
   */
  async checkoutIntoDoc(actor, branchName) {
    const state = await this.branchService.checkoutBranch(actor, { name: branchName });
    applyBranchStateToDocumentController(this.doc, state);
    return state;
  }

  /**
   * @param {Actor} actor
   * @param {string} sourceBranch
   * @param {ConflictResolution[]} resolutions
   * @param {string} [message]
   */
  async mergeIntoDoc(actor, sourceBranch, resolutions, message) {
    const result = await this.branchService.merge(actor, {
      sourceBranch,
      resolutions,
      message,
    });
    applyBranchStateToDocumentController(this.doc, result.state);
    return result;
  }
}

