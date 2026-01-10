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
