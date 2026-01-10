import test from "node:test";
import assert from "node:assert/strict";

import {
  getCellPermissions,
  maskCellUpdatesForUser,
  roleCanComment,
  roleCanEdit,
  roleCanRead
} from "./index.js";

test("document roles: viewer can read, commenter can comment, editor can edit", () => {
  assert.equal(roleCanRead("viewer"), true);
  assert.equal(roleCanEdit("viewer"), false);
  assert.equal(roleCanComment("viewer"), false);

  assert.equal(roleCanRead("commenter"), true);
  assert.equal(roleCanEdit("commenter"), false);
  assert.equal(roleCanComment("commenter"), true);

  assert.equal(roleCanRead("editor"), true);
  assert.equal(roleCanEdit("editor"), true);
  assert.equal(roleCanComment("editor"), true);
});

test("range permissions: read allowlist masks values; edit allowlist blocks edits", () => {
  const restrictions = [
    {
      range: { sheetName: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
      readAllowlist: ["u-editor"]
    },
    {
      range: { sheetName: "Sheet1", startRow: 0, startCol: 1, endRow: 0, endCol: 1 },
      editAllowlist: ["u-owner"]
    }
  ];

  const readDenied = getCellPermissions({
    role: "viewer",
    restrictions,
    userId: "u-viewer",
    cell: { sheetId: "Sheet1", row: 0, col: 0 }
  });
  assert.deepEqual(readDenied, { canRead: false, canEdit: false });

  const readAllowed = getCellPermissions({
    role: "viewer",
    restrictions,
    userId: "u-editor",
    cell: { sheetId: "Sheet1", row: 0, col: 0 }
  });
  assert.deepEqual(readAllowed, { canRead: true, canEdit: false });

  const editDenied = getCellPermissions({
    role: "editor",
    restrictions,
    userId: "u-editor",
    cell: { sheetId: "Sheet1", row: 0, col: 1 }
  });
  assert.equal(editDenied.canRead, true);
  assert.equal(editDenied.canEdit, false);

  const editAllowed = getCellPermissions({
    role: "editor",
    restrictions,
    userId: "u-owner",
    cell: { sheetId: "Sheet1", row: 0, col: 1 }
  });
  assert.equal(editAllowed.canRead, true);
  assert.equal(editAllowed.canEdit, true);

  const masked = maskCellUpdatesForUser({
    role: "viewer",
    restrictions,
    userId: "u-viewer",
    updates: [{ cell: { sheetId: "Sheet1", row: 0, col: 0 }, value: "secret" }]
  });
  assert.equal(masked[0].value, "###");
});

