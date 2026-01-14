import {
  applyBranchStateToDocumentController,
  documentControllerToBranchState,
} from "./branchStateAdapter.js";
import { normalizeDocumentState } from "../../../../../packages/versioning/branches/src/browser.js";
import { parseA1 } from "../../document/coords.js";

/**
 * @typedef {import("../../document/documentController.js").DocumentController} DocumentController
 * @typedef {import("../../../../../packages/versioning/branches/src/browser.js").BranchService} BranchService
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

    // DocumentController does not model some workbook-level metadata like named ranges
    // or comments. Preserve whatever is currently stored in the branch head and only
    // overlay the cell + sheet metadata we can observe locally.
    const baseState = normalizeDocumentState(await this.branchService.getCurrentState());

    /** @type {import("../../../../../packages/versioning/branches/src/types.js").DocumentState} */
    const merged = structuredClone(baseState);

    const supportsSheetMetadata =
      typeof this.doc.getSheetMeta === "function" ||
      typeof this.doc.getSheetMetadata === "function" ||
      this.doc.sheetMetaById != null ||
      this.doc.sheetMetadataById != null ||
      this.doc.sheetMeta != null;

    // Sheet order can be preserved independently of richer sheet metadata (names/visibility/tab color).
    // Feature-detect a dedicated reordering API to avoid treating legacy insertion order as
    // user-controlled tab order.
    const supportsSheetOrder = typeof this.doc.reorderSheets === "function";

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
          // Fail closed: treat any defined `enc` field (including `null`) as an
          // encryption marker so masked placeholder detection never applies to
          // encrypted cells.
          cell.enc === undefined &&
          cell.formula == null &&
          cell.value === MASKED_CELL_VALUE &&
          (canEdit === false || (baseCell && typeof baseCell === "object" && baseCell.enc !== undefined));

        if (isMasked && baseSheet[addr] !== undefined) {
          mergedSheet[addr] = baseSheet[addr];
        } else {
          mergedSheet[addr] = cell;
        }
      }

      merged.cells[sheetId] = mergedSheet;

      if (!merged.sheets.metaById[sheetId]) {
        const nextMeta = nextState.sheets.metaById[sheetId];
        /** @type {import("../../../../../packages/versioning/branches/src/types.js").SheetMeta} */
        const sheetMeta = {
          id: sheetId,
          // Before DocumentController tracked sheet metadata, BranchService owned sheet
          // names. Now DocumentController is authoritative for display names; keep the
          // feature-detection so older controller instances (or persisted branch histories)
          // don't get their existing sheet names clobbered.
          name: supportsSheetMetadata ? (nextMeta?.name ?? sheetId) : sheetId,
          // Avoid deep-cloning untrusted view payloads here; we normalize the merged state before committing.
          view: nextMeta?.view ? nextMeta.view : { frozenRows: 0, frozenCols: 0 },
        };
        if (supportsSheetMetadata && nextMeta?.visibility) {
          sheetMeta.visibility = nextMeta.visibility;
        }
        if (supportsSheetMetadata && nextMeta && "tabColor" in nextMeta) {
          if (nextMeta.tabColor === null) sheetMeta.tabColor = null;
          else if (typeof nextMeta.tabColor === "string") sheetMeta.tabColor = nextMeta.tabColor;
        }
        merged.sheets.metaById[sheetId] = sheetMeta;
      } else {
        // DocumentController always owns per-sheet view state (e.g. frozen panes).
        // Sheet metadata is also owned by DocumentController, but keep the defensive
        // feature-detection to support older docs/clients.
        const nextMeta = nextState.sheets.metaById[sheetId];
        if (supportsSheetMetadata && nextMeta?.name != null) {
          const existingName = merged.sheets.metaById[sheetId]?.name;
          // Avoid clobbering existing BranchService sheet names if the
          // DocumentController metadata surface exists but hasn't populated a
          // custom display name yet (common during gradual rollouts).
          if (nextMeta.name !== sheetId || existingName == null || existingName === sheetId) {
            merged.sheets.metaById[sheetId] = { ...merged.sheets.metaById[sheetId], name: nextMeta.name };
          }
        }
        if (supportsSheetMetadata && nextMeta?.visibility) {
          merged.sheets.metaById[sheetId] = { ...merged.sheets.metaById[sheetId], visibility: nextMeta.visibility };
        }
        if (supportsSheetMetadata && nextMeta && "tabColor" in nextMeta) {
          if (nextMeta.tabColor === null) {
            merged.sheets.metaById[sheetId] = { ...merged.sheets.metaById[sheetId], tabColor: null };
          } else if (typeof nextMeta.tabColor === "string") {
            merged.sheets.metaById[sheetId] = { ...merged.sheets.metaById[sheetId], tabColor: nextMeta.tabColor };
          }
        }
        if (nextMeta?.view) {
          merged.sheets.metaById[sheetId] = { ...merged.sheets.metaById[sheetId], view: nextMeta.view };
        }
      }
    }

    if (supportsSheetOrder) {
      // Treat the DocumentController's sheet ordering as canonical so branch commits
      // preserve user-visible tab order.
      const desired = nextState.sheets.order;
      /** @type {string[]} */
      const nextOrder = [];
      const seen = new Set();
      for (const sheetId of desired) {
        if (seen.has(sheetId)) continue;
        nextOrder.push(sheetId);
        seen.add(sheetId);
      }
      // Preserve any sheets that still only exist in BranchService (e.g. legacy
      // empty sheets / workbook metadata not represented in DocumentController).
      for (const sheetId of merged.sheets.order) {
        if (seen.has(sheetId)) continue;
        nextOrder.push(sheetId);
        seen.add(sheetId);
      }
      merged.sheets.order = nextOrder;
    } else {
      // Legacy DocumentController: it does not maintain sheet order, so append any
      // new sheets discovered during this commit.
      for (const sheetId of nextState.sheets.order) {
        if (!merged.sheets.order.includes(sheetId)) merged.sheets.order.push(sheetId);
      }
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
