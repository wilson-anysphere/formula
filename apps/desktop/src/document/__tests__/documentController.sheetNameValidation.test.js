import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("DocumentController.renameSheet rejects leading/trailing apostrophes (Excel rule)", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);

  assert.throws(
    () => doc.renameSheet("Sheet1", "'Budget"),
    /apostrophe/i,
    "expected leading apostrophe to be rejected",
  );

  assert.throws(
    () => doc.renameSheet("Sheet1", "Budget'"),
    /apostrophe/i,
    "expected trailing apostrophe to be rejected",
  );
});

test("DocumentController.renameSheet rejects Unicode case-insensitive duplicates", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  doc.setCellValue("Sheet2", "A1", 2);

  doc.renameSheet("Sheet1", "é");

  assert.throws(
    () => doc.renameSheet("Sheet2", "É"),
    /already exists/i,
    "expected Unicode-case-folding duplicate to be rejected",
  );
});

