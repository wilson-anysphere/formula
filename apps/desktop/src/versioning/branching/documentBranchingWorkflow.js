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
   * Convenience passthroughs so UI code can treat this as a BranchService-like
   * object while ensuring checkout/merge are applied to the live document.
   */
  async listBranches() {
    return this.branchService.listBranches();
  }

  /**
   * @param {Actor} actor
   * @param {{ name: string, description?: string }} input
   */
  async createBranch(actor, input) {
    return this.branchService.createBranch(actor, input);
  }

  /**
   * @param {Actor} actor
   * @param {{ oldName: string, newName: string }} input
   */
  async renameBranch(actor, input) {
    return this.branchService.renameBranch(actor, input);
  }

  /**
   * @param {Actor} actor
   * @param {{ name: string }} input
   */
  async deleteBranch(actor, input) {
    return this.branchService.deleteBranch(actor, input);
  }

  /**
   * @param {Actor} actor
   * @param {{ sourceBranch: string }} input
   */
  async previewMerge(actor, input) {
    return this.branchService.previewMerge(actor, input);
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
   * Alias for BranchService UI integrations that expect `checkoutBranch`.
   *
   * @param {Actor} actor
   * @param {{ name: string }} input
   */
  async checkoutBranch(actor, input) {
    return this.checkoutIntoDoc(actor, input.name);
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

  /**
   * Alias for BranchService UI integrations that expect `merge`.
   *
   * @param {Actor} actor
   * @param {{ sourceBranch: string, resolutions: ConflictResolution[], message?: string }} input
   */
  async merge(actor, input) {
    return this.mergeIntoDoc(actor, input.sourceBranch, input.resolutions, input.message);
  }
}
