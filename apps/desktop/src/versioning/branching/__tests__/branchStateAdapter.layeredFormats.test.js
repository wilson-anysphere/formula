import test from "node:test";
import assert from "node:assert/strict";
import { TextDecoder } from "node:util";

import {
  applyBranchStateToDocumentController,
  documentControllerToBranchState,
} from "../branchStateAdapter.js";
import { StyleTable } from "../../../formatting/styleTable.js";

class FakeDocumentController {
  constructor() {
    /** @type {{ sheets: Map<string, any> }} */
    this.model = { sheets: new Map() };
    this.styleTable = new StyleTable();
    /** @type {any} */
    this.lastAppliedSnapshot = null;
  }

  getSheetIds() {
    return Array.from(this.model.sheets.keys());
  }

  getSheetView(sheetId) {
    return this.model.sheets.get(sheetId)?.view ?? { frozenRows: 0, frozenCols: 0 };
  }

  /**
   * Minimal `DocumentController.applyState` implementation that stores the decoded snapshot.
   * The real DocumentController interns style objects into its style table; we mimic that
   * behavior for layered formats so `documentControllerToBranchState` can test styleId -> style
   * object conversion on export.
   *
   * @param {Uint8Array} encoded
   */
  applyState(encoded) {
    const decoded = new TextDecoder().decode(encoded);
    const snapshot = JSON.parse(decoded);
    this.lastAppliedSnapshot = snapshot;

    const sheets = Array.isArray(snapshot?.sheets) ? snapshot.sheets : [];
    for (const sheet of sheets) {
      const sheetId = String(sheet?.id ?? "");
      if (!sheetId) continue;
      let model = this.model.sheets.get(sheetId);
      if (!model) {
        model = { cells: new Map(), view: { frozenRows: 0, frozenCols: 0 } };
        this.model.sheets.set(sheetId, model);
      }

      /** @type {any} */
      const view = {
        frozenRows: sheet?.frozenRows ?? 0,
        frozenCols: sheet?.frozenCols ?? 0,
      };
      if (sheet?.colWidths) view.colWidths = sheet.colWidths;
      if (sheet?.rowHeights) view.rowHeights = sheet.rowHeights;

      // Layered formats: intern style objects back into style ids (DocumentController behavior).
      if (sheet?.defaultFormat) {
        view.defaultFormat = this.styleTable.intern(sheet.defaultFormat);
      }
      if (sheet?.rowFormats) {
        view.rowFormats = {};
        for (const [key, format] of Object.entries(sheet.rowFormats)) {
          view.rowFormats[key] = this.styleTable.intern(format);
        }
      }
      if (sheet?.colFormats) {
        view.colFormats = {};
        for (const [key, format] of Object.entries(sheet.colFormats)) {
          view.colFormats[key] = this.styleTable.intern(format);
        }
      }

      model.view = view;
    }
  }
}

test("branchStateAdapter preserves column default format through a branch state roundtrip", () => {
  const doc = new FakeDocumentController();
  const boldStyleId = doc.styleTable.intern({ bold: true });

  doc.model.sheets.set("Sheet1", {
    cells: new Map(),
    view: { frozenRows: 0, frozenCols: 0, colFormats: { "0": boldStyleId } },
  });

  const state = documentControllerToBranchState(doc);

  assert.deepEqual(state.sheets.metaById.Sheet1.view.colFormats, { "0": { bold: true } });

  const restored = new FakeDocumentController();
  applyBranchStateToDocumentController(restored, state);

  // Verify the snapshot sent to DocumentController includes the layered formats.
  assert.deepEqual(restored.lastAppliedSnapshot.sheets[0].colFormats, { "0": { bold: true } });

  // Export back to BranchService state. Our fake applyState interns formats into style ids,
  // so this verifies the adapter converts style ids back into style objects.
  const stateRoundtrip = documentControllerToBranchState(restored);
  assert.deepEqual(stateRoundtrip.sheets.metaById.Sheet1.view.colFormats, { "0": { bold: true } });
});

