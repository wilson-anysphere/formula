import {
  applyBranchStateToDocumentController,
  documentControllerToBranchState,
} from "./branchStateAdapter.js";
import { normalizeDocumentState } from "../../../../../packages/versioning/branches/src/state.js";
import { parseA1 } from "../../document/coords.js";

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
    const nextState = normalizeDocumentState(documentControllerToBranchState(this.doc));

    // DocumentController doesn't model workbook-level metadata like empty sheets,
    // named ranges, or comments. Preserve whatever is currently stored in the
    // branch head and only overlay the cell edits we can observe locally.
    const baseState = normalizeDocumentState(await this.branchService.getCurrentState());

    /** @type {import("../../../../../packages/versioning/branches/src/types.js").DocumentState} */
    const merged = structuredClone(baseState);

    const MASKED_CELL_VALUE = "###";

    // DocumentController snapshot is authoritative for cell contents *except*
    // for masked/unreadable cells ("###"). Preserve those cells from the branch
    // head so permissions/encryption don't get committed as literal placeholders.
    for (const [sheetId, nextSheet] of Object.entries(nextState.cells)) {
      const baseSheet = merged.cells[sheetId] ?? {};
      /** @type {Record<string, any>} */
      const mergedSheet = {};

      for (const [addr, cell] of Object.entries(nextSheet ?? {})) {
        const baseCell = baseSheet[addr];

        let canEdit = true;
        if (typeof this.doc.canEditCell === "function") {
          try {
            const coord = parseA1(addr);
            canEdit = this.doc.canEditCell({ sheetId, row: coord.row, col: coord.col });
          } catch {
            canEdit = true;
          }
        }

        const isMasked =
          cell &&
          typeof cell === "object" &&
          cell.enc == null &&
          cell.formula == null &&
          cell.value === MASKED_CELL_VALUE &&
          (canEdit === false || (baseCell && typeof baseCell === "object" && baseCell.enc != null));

        if (isMasked && baseSheet[addr] !== undefined) {
          mergedSheet[addr] = baseSheet[addr];
        } else {
          mergedSheet[addr] = cell;
        }
      }

      merged.cells[sheetId] = mergedSheet;

      if (!merged.sheets.metaById[sheetId]) {
        const nextMeta = nextState.sheets.metaById[sheetId];
        merged.sheets.metaById[sheetId] = {
          id: sheetId,
          name: sheetId,
          view: nextMeta?.view ? structuredClone(nextMeta.view) : { frozenRows: 0, frozenCols: 0 },
        };
      } else {
        // DocumentController only knows about sheet ids (names/order live in BranchService),
        // but it *does* own per-sheet view state (e.g. frozen panes). Preserve the existing
        // sheet name while syncing view state from the live workbook.
        const nextMeta = nextState.sheets.metaById[sheetId];
        if (nextMeta?.view) {
          merged.sheets.metaById[sheetId] = { ...merged.sheets.metaById[sheetId], view: structuredClone(nextMeta.view) };
        }
      }
    }

    // Ensure any new sheets are present in the ordering (DocumentController
    // doesn't maintain sheet order, so append).
    for (const sheetId of nextState.sheets.order) {
      if (!merged.sheets.order.includes(sheetId)) merged.sheets.order.push(sheetId);
    }

    return this.branchService.commit(actor, { nextState: normalizeDocumentState(merged), message });
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
