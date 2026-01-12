import assert from "node:assert/strict";
import test from "node:test";

import { DocumentController } from "../../document/documentController.js";
import {
  applyBranchStateToDocumentController,
  documentControllerToBranchState,
} from "./branchStateAdapter.js";

/**
 * Attach a lightweight sheet-metadata shim to a DocumentController instance so we can
 * unit-test branching adapters before DocumentController's native sheet metadata
 * implementation lands (Task 201).
 *
 * The adapter under test should:
 * - read display names via `doc.getSheetMeta(sheetId).name`
 * - write display names by including `name` on the snapshot sheet objects passed to `applyState`
 *
 * @param {DocumentController} doc
 * @returns {{ metaById: Map<string, any> }}
 */
function attachSheetMetadataShim(doc) {
  /** @type {Map<string, any>} */
  const metaById = new Map();
  /** @type {string[] | null} */
  let lastSheetOrder = null;

  // Read path used by the adapter.
  doc.getSheetMeta = (sheetId) => metaById.get(sheetId) ?? null;

  // Test-only helper for seeding initial metadata.
  doc.setSheetMeta = (sheetId, meta) => {
    const prev = metaById.get(sheetId) ?? { id: sheetId };
    metaById.set(sheetId, { ...prev, ...meta, id: sheetId });
  };

  const decode = (bytes) => {
    if (typeof TextDecoder !== "undefined") return new TextDecoder().decode(bytes);
    // eslint-disable-next-line no-undef
    return Buffer.from(bytes).toString("utf8");
  };

  const originalApplyState = doc.applyState.bind(doc);
  doc.applyState = (snapshot) => {
    // Capture names/order from the snapshot passed by the adapter.
    const parsed = JSON.parse(decode(snapshot));
    const sheets = Array.isArray(parsed?.sheets) ? parsed.sheets : [];
    lastSheetOrder = Array.isArray(parsed?.sheetOrder) ? parsed.sheetOrder.map((id) => String(id)) : null;
    metaById.clear();
    for (const sheet of sheets) {
      const id = sheet?.id;
      if (typeof id !== "string" || id.length === 0) continue;
      const name = sheet?.name;
      const visibility = sheet?.visibility;
      const rawTabColor = sheet?.tabColor;
      const tabColor =
        rawTabColor && typeof rawTabColor === "object" && typeof rawTabColor.rgb === "string"
          ? rawTabColor.rgb
          : rawTabColor;
      metaById.set(id, {
        id,
        name: typeof name === "string" && name.length > 0 ? name : id,
        ...(visibility ? { visibility } : {}),
        ...(tabColor !== undefined ? { tabColor } : {}),
      });
    }
    originalApplyState(snapshot);
  };

  return { metaById, getLastSheetOrder: () => lastSheetOrder };
}

test("documentControllerToBranchState/applyBranchStateToDocumentController: preserves sheet order", () => {
  const doc = new DocumentController();

  // Create sheets in a non-lexicographic order.
  doc.setCellValue("SheetB", "A1", 1);
  doc.setCellValue("SheetA", "A1", 1);
  doc.setCellValue("SheetC", "A1", 1);

  const state = documentControllerToBranchState(doc);
  assert.deepEqual(state.sheets.order, ["SheetB", "SheetA", "SheetC"]);

  const restored = new DocumentController();
  const { getLastSheetOrder } = attachSheetMetadataShim(restored);
  applyBranchStateToDocumentController(restored, state);
  assert.deepEqual(getLastSheetOrder(), ["SheetB", "SheetA", "SheetC"]);
  assert.deepEqual(restored.getSheetIds(), ["SheetB", "SheetA", "SheetC"]);

  const roundTrip = documentControllerToBranchState(restored);
  assert.deepEqual(roundTrip.sheets.order, ["SheetB", "SheetA", "SheetC"]);
});

test("documentControllerToBranchState/applyBranchStateToDocumentController: round-trips sheet display names", () => {
  const doc = new DocumentController();
  attachSheetMetadataShim(doc);

  doc.setCellValue("Sheet1", "A1", 1);
  doc.setCellValue("Sheet2", "A1", 2);
  doc.setSheetMeta("Sheet1", { name: "Summary" });
  doc.setSheetMeta("Sheet2", { name: "Data" });

  const state = documentControllerToBranchState(doc);
  assert.equal(state.sheets.metaById.Sheet1?.name, "Summary");
  assert.equal(state.sheets.metaById.Sheet2?.name, "Data");

  const restored = new DocumentController();
  const { metaById } = attachSheetMetadataShim(restored);

  applyBranchStateToDocumentController(restored, state);
  assert.deepEqual(metaById.get("Sheet1")?.name, "Summary");
  assert.deepEqual(metaById.get("Sheet2")?.name, "Data");

  const roundTrip = documentControllerToBranchState(restored);
  assert.equal(roundTrip.sheets.metaById.Sheet1?.name, "Summary");
  assert.equal(roundTrip.sheets.metaById.Sheet2?.name, "Data");
});

test("documentControllerToBranchState/applyBranchStateToDocumentController: includes sheet visibility + tabColor when present", () => {
  const doc = new DocumentController();
  attachSheetMetadataShim(doc);

  doc.setCellValue("Sheet1", "A1", 1);
  doc.setSheetMeta("Sheet1", { name: "Summary", visibility: "hidden", tabColor: "FF00FF00" });
  doc.setCellValue("Sheet2", "A1", 2);
  doc.setSheetMeta("Sheet2", { name: "Data", visibility: "visible", tabColor: null });

  const state = documentControllerToBranchState(doc);
  assert.equal(state.sheets.metaById.Sheet1?.visibility, "hidden");
  assert.equal(state.sheets.metaById.Sheet1?.tabColor, "FF00FF00");
  assert.equal(state.sheets.metaById.Sheet2?.visibility, "visible");
  assert.equal(state.sheets.metaById.Sheet2?.tabColor, null);

  const restored = new DocumentController();
  const { metaById } = attachSheetMetadataShim(restored);

  applyBranchStateToDocumentController(restored, state);
  assert.equal(metaById.get("Sheet1")?.visibility, "hidden");
  assert.equal(metaById.get("Sheet1")?.tabColor, "FF00FF00");
  assert.equal(metaById.get("Sheet2")?.visibility, "visible");
  assert.equal(metaById.get("Sheet2")?.tabColor, null);

  const roundTrip = documentControllerToBranchState(restored);
  assert.equal(roundTrip.sheets.metaById.Sheet1?.visibility, "hidden");
  assert.equal(roundTrip.sheets.metaById.Sheet1?.tabColor, "FF00FF00");
  assert.equal(roundTrip.sheets.metaById.Sheet2?.visibility, "visible");
  assert.equal(roundTrip.sheets.metaById.Sheet2?.tabColor, null);
});
