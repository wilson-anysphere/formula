import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";
import { getCellPermissions } from "../../../../../packages/collab/permissions/index.js";

function makeGuard({ role, userId, restrictions }) {
  return (cell) =>
    getCellPermissions({
      role,
      userId,
      restrictions,
      cell
    }).canEdit;
}

test("DocumentController can block edits via canEditCell guard (viewer vs editor)", () => {
  const restrictions = [];

  const viewerDoc = new DocumentController({
    canEditCell: makeGuard({ role: "viewer", userId: "u-viewer", restrictions })
  });
  viewerDoc.setCellValue("Sheet1", { row: 0, col: 0 }, "blocked");
  assert.equal(viewerDoc.getCell("Sheet1", { row: 0, col: 0 }).value, null);

  const editorDoc = new DocumentController({
    canEditCell: makeGuard({ role: "editor", userId: "u-editor", restrictions })
  });
  editorDoc.setCellValue("Sheet1", { row: 0, col: 0 }, "allowed");
  assert.equal(editorDoc.getCell("Sheet1", { row: 0, col: 0 }).value, "allowed");
});

test("DocumentController respects range-level edit allowlists", () => {
  const restrictions = [
    {
      range: { sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
      editAllowlist: ["u-owner"]
    }
  ];

  const editorDoc = new DocumentController({
    canEditCell: makeGuard({ role: "editor", userId: "u-editor", restrictions })
  });
  editorDoc.setCellValue("Sheet1", { row: 0, col: 0 }, "nope");
  assert.equal(editorDoc.getCell("Sheet1", { row: 0, col: 0 }).value, null);

  const ownerDoc = new DocumentController({
    canEditCell: makeGuard({ role: "owner", userId: "u-owner", restrictions })
  });
  ownerDoc.setCellValue("Sheet1", { row: 0, col: 0 }, "ok");
  assert.equal(ownerDoc.getCell("Sheet1", { row: 0, col: 0 }).value, "ok");
});

test("DocumentController allows sheet view mutations even when cell edits are blocked (freeze panes + row/col sizes)", () => {
  const restrictions = [];

  const viewerDoc = new DocumentController({
    canEditCell: makeGuard({ role: "viewer", userId: "u-viewer", restrictions })
  });

  viewerDoc.setFrozen("Sheet1", 2, 1);
  viewerDoc.setColWidth("Sheet1", 0, 120);
  viewerDoc.setRowHeight("Sheet1", 0, 40);

  const viewerView = viewerDoc.getSheetView("Sheet1");
  assert.equal(viewerView.frozenRows, 2);
  assert.equal(viewerView.frozenCols, 1);
  assert.equal(viewerView.colWidths?.["0"], 120);
  assert.equal(viewerView.rowHeights?.["0"], 40);

  const editorDoc = new DocumentController({
    canEditCell: makeGuard({ role: "editor", userId: "u-editor", restrictions })
  });
  editorDoc.setFrozen("Sheet1", 2, 1);
  editorDoc.setColWidth("Sheet1", 0, 120);
  editorDoc.setRowHeight("Sheet1", 0, 40);

  const view = editorDoc.getSheetView("Sheet1");
  assert.equal(view.frozenRows, 2);
  assert.equal(view.frozenCols, 1);
  assert.equal(view.colWidths?.["0"], 120);
  assert.equal(view.rowHeights?.["0"], 40);
});

test("DocumentController allows formatting defaults via setRangeFormat even when cell edits are blocked (full column)", () => {
  const restrictions = [];
  const viewerDoc = new DocumentController({
    canEditCell: makeGuard({ role: "viewer", userId: "u-viewer", restrictions })
  });

  const ok = viewerDoc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });
  assert.equal(ok, true);
  assert.equal(viewerDoc.getCellFormat("Sheet1", "A1")?.font?.bold, true);
});

test("DocumentController allows sheet view mutations when at least some cells are editable (partial range restrictions)", () => {
  const restrictions = [
    {
      // Restrict a small top-left range; editor can still edit the rest of the sheet.
      range: { sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 1, endCol: 1 },
      editAllowlist: ["u-owner"],
    },
  ];

  const editorDoc = new DocumentController({
    canEditCell: makeGuard({ role: "editor", userId: "u-editor", restrictions }),
  });

  editorDoc.setFrozen("Sheet1", 2, 1);
  editorDoc.setColWidth("Sheet1", 0, 120);
  editorDoc.setRowHeight("Sheet1", 0, 40);

  const view = editorDoc.getSheetView("Sheet1");
  assert.equal(view.frozenRows, 2);
  assert.equal(view.frozenCols, 1);
  assert.equal(view.colWidths?.["0"], 120);
  assert.equal(view.rowHeights?.["0"], 40);
});
