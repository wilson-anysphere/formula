import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { DocumentBranchingWorkflow } from "../apps/desktop/src/versioning/branching/documentBranchingWorkflow.js";
import { BranchService } from "../packages/versioning/branches/src/BranchService.js";
import { InMemoryBranchStore } from "../packages/versioning/branches/src/store/InMemoryBranchStore.js";

test("DocumentController + BranchService: checkout/merge mutate the live workbook", async () => {
  const actor = { userId: "u1", role: "owner" };

  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  doc.setCellFormula("Sheet1", "B1", "=A1*2");
  doc.setRangeFormat("Sheet1", "A1", { font: { bold: true }, meta: { foo: "bar", nested: { n: 1 } } });
  doc.setFrozen("Sheet1", 2, 1);
  doc.setCellValue("Sheet2", "A1", "s2");

  const store = new InMemoryBranchStore();
  const branchService = new BranchService({ docId: "doc1", store });
  // Start with explicit sheet ids (including an empty sheet) to exercise empty-sheet handling.
  await branchService.init(actor, { sheets: { Sheet1: {}, Sheet2: {}, EmptySheet: {} } });

  const workflow = new DocumentBranchingWorkflow({ doc, branchService });
  await workflow.commitCurrentState(actor, "initial");

  await branchService.createBranch(actor, { name: "feature" });

  await workflow.checkoutIntoDoc(actor, "feature");
  assert.deepEqual(new Set(doc.getSheetIds()), new Set(["Sheet1", "Sheet2", "EmptySheet"]));
  assert.deepEqual(doc.getSheetView("Sheet1"), { frozenRows: 2, frozenCols: 1 });

  // Change frozen panes on the feature branch to ensure view state is committed and merged.
  doc.setFrozen("Sheet1", 3, 0);
  doc.setCellValue("Sheet1", "A1", 10);
  doc.setCellValue("Sheet1", "C1", 99);
  await workflow.commitCurrentState(actor, "feature edits");

  await workflow.checkoutIntoDoc(actor, "main");
  // C1 only existed on the feature branch, so checkout must delete it locally.
  assert.equal(doc.getCell("Sheet1", "C1").value, null);

  doc.setCellValue("Sheet1", "A1", 5);
  doc.setCellValue("Sheet1", "D1", 123);
  await workflow.commitCurrentState(actor, "main edits");

  const preview = await branchService.previewMerge(actor, { sourceBranch: "feature" });
  assert.equal(preview.conflicts.length, 1);
  assert.deepEqual(preview.conflicts[0], {
    type: "cell",
    sheetId: "Sheet1",
    cell: "A1",
    reason: "content",
    base: { value: 1, format: { font: { bold: true }, meta: { foo: "bar", nested: { n: 1 } } } },
    ours: { value: 5, format: { font: { bold: true }, meta: { foo: "bar", nested: { n: 1 } } } },
    theirs: { value: 10, format: { font: { bold: true }, meta: { foo: "bar", nested: { n: 1 } } } },
  });

  await workflow.mergeIntoDoc(actor, "feature", [{ conflictIndex: 0, choice: "theirs" }]);

  assert.equal(doc.getCell("Sheet1", "A1").value, 10);
  assert.equal(doc.getCell("Sheet1", "B1").formula, "=A1*2");
  assert.equal(doc.getCell("Sheet1", "C1").value, 99);
  assert.equal(doc.getCell("Sheet1", "D1").value, 123);
  assert.equal(doc.getCell("Sheet2", "A1").value, "s2");
  assert.deepEqual(doc.getSheetView("Sheet1"), { frozenRows: 3, frozenCols: 0 });

  const a1 = doc.getCell("Sheet1", "A1");
  assert.deepEqual(doc.styleTable.get(a1.styleId), {
    font: { bold: true },
    meta: { foo: "bar", nested: { n: 1 } },
  });
});

test("DocumentController + BranchService: format-only conflicts round-trip through adapter", async () => {
  const actor = { userId: "u1", role: "owner" };

  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  doc.setRangeFormat("Sheet1", "A1", { numberFormat: "currency" });

  const store = new InMemoryBranchStore();
  const branchService = new BranchService({ docId: "doc-format", store });
  await branchService.init(actor, { sheets: { Sheet1: {} } });

  const workflow = new DocumentBranchingWorkflow({ doc, branchService });
  await workflow.commitCurrentState(actor, "base");

  await branchService.createBranch(actor, { name: "fmt" });

  await workflow.checkoutIntoDoc(actor, "fmt");
  doc.setRangeFormat("Sheet1", "A1", { numberFormat: "percent" });
  await workflow.commitCurrentState(actor, "fmt: percent");

  await workflow.checkoutIntoDoc(actor, "main");
  doc.setRangeFormat("Sheet1", "A1", { numberFormat: "accounting" });
  await workflow.commitCurrentState(actor, "main: accounting");

  const preview = await branchService.previewMerge(actor, { sourceBranch: "fmt" });
  assert.equal(preview.conflicts.length, 1);
  assert.equal(preview.conflicts[0]?.reason, "format");

  await workflow.mergeIntoDoc(actor, "fmt", [{ conflictIndex: 0, choice: "theirs" }]);

  const a1 = doc.getCell("Sheet1", "A1");
  assert.equal(a1.value, 1);
  assert.deepEqual(doc.styleTable.get(a1.styleId), { numberFormat: "percent" });
});

test("DocumentController + BranchService: delete-vs-edit conflicts clear cells when resolved as delete", async () => {
  const actor = { userId: "u1", role: "owner" };

  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);

  const store = new InMemoryBranchStore();
  const branchService = new BranchService({ docId: "doc-delete", store });
  await branchService.init(actor, { sheets: { Sheet1: {} } });

  const workflow = new DocumentBranchingWorkflow({ doc, branchService });
  await workflow.commitCurrentState(actor, "base");

  await branchService.createBranch(actor, { name: "del" });

  await workflow.checkoutIntoDoc(actor, "del");
  doc.setCellValue("Sheet1", "A1", null);
  await workflow.commitCurrentState(actor, "delete A1");

  await workflow.checkoutIntoDoc(actor, "main");
  doc.setCellValue("Sheet1", "A1", 2);
  await workflow.commitCurrentState(actor, "edit A1");

  const preview = await branchService.previewMerge(actor, { sourceBranch: "del" });
  assert.equal(preview.conflicts.length, 1);
  assert.equal(preview.conflicts[0]?.reason, "delete-vs-edit");

  await workflow.mergeIntoDoc(actor, "del", [{ conflictIndex: 0, choice: "theirs" }]);
  assert.equal(doc.getCell("Sheet1", "A1").value, null);
  assert.equal(doc.getCell("Sheet1", "A1").formula, null);
});

test("DocumentController + BranchService: masked cells do not overwrite encrypted branch state", async () => {
  const actor = { userId: "u1", role: "owner" };

  const doc = new DocumentController();

  const store = new InMemoryBranchStore();
  const branchService = new BranchService({ docId: "doc-encrypted", store });

  const enc = {
    v: 1,
    alg: "AES-256-GCM",
    keyId: "k1",
    ivBase64: "iv",
    tagBase64: "tag",
    ciphertextBase64: "ct",
  };

  await branchService.init(actor, {
    schemaVersion: 1,
    sheets: { order: ["Sheet1"], metaById: { Sheet1: { id: "Sheet1", name: "Sheet1" } } },
    cells: { Sheet1: { A1: { enc } } },
    namedRanges: {},
    comments: {},
  });

  const workflow = new DocumentBranchingWorkflow({ doc, branchService });
  await workflow.checkoutIntoDoc(actor, "main");

  assert.equal(doc.getCell("Sheet1", "A1").value, "###");
  assert.equal(doc.getCell("Sheet1", "A1").formula, null);

  await workflow.commitCurrentState(actor, "commit masked snapshot");
  const state = await branchService.getCurrentState();
  assert.deepEqual(state.cells.Sheet1.A1, { enc });
});
