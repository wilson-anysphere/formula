import { applyBranchStateToYjsDoc, branchStateFromYjsDoc } from "./branchStateAdapter.js";

/**
 * @typedef {import("yjs").Doc} YDoc
 * @typedef {import("../types.js").DocumentState} DocumentState
 */

/**
 * Convert a Yjs spreadsheet document into a versioning DocumentState snapshot.
 *
 * @param {YDoc} ydoc
 * @returns {DocumentState}
 */
export function yjsDocToDocumentState(ydoc) {
  return branchStateFromYjsDoc(ydoc);
}

/**
 * Apply a DocumentState snapshot into a Yjs spreadsheet document.
 *
 * This mutates the live workbook state (global checkout/merge semantics).
 *
 * @param {YDoc} ydoc
 * @param {DocumentState} state
 * @param {{ origin?: any }} [opts]
 */
export function applyDocumentStateToYjsDoc(ydoc, state, opts = {}) {
  applyBranchStateToYjsDoc(ydoc, state, opts);
}

export {};

