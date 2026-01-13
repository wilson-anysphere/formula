import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createCollabSession } from "../src/index.ts";
import { EncryptedRangeManager, createEncryptionPolicyFromDoc } from "../../encrypted-ranges/src/index.ts";

test("CollabSession encryption policy: shouldEncryptCell from shared encryptedRanges blocks plaintext writes without keys", async () => {
  const docId = "collab-session-encrypted-ranges-policy-test-doc";
  const doc = new Y.Doc({ guid: docId });

  // Define a protected/encrypted range in the shared workbook metadata.
  const ranges = new EncryptedRangeManager({ doc });
  ranges.add({ sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId: "k1" });

  const policy = createEncryptionPolicyFromDoc(doc);

  // Simulate a client that knows the policy (ranges + key ids) but does *not* have
  // the actual encryption keys. This client must refuse plaintext writes.
  const session = createCollabSession({
    doc,
    encryption: {
      keyForCell: () => null,
      shouldEncryptCell: policy.shouldEncryptCell,
    },
  });

  assert.equal(session.canEditCell({ sheetId: "Sheet1", row: 0, col: 0 }), false);
  assert.equal(await session.safeSetCellValue("Sheet1:0:0", "plaintext"), false);
  assert.equal(session.cells.has("Sheet1:0:0"), false);
  await assert.rejects(session.setCellValue("Sheet1:0:0", "plaintext"));

  // Cells outside the range should remain editable (no keys needed).
  assert.equal(await session.safeSetCellValue("Sheet1:0:1", "ok"), true);
  assert.equal((await session.getCell("Sheet1:0:1"))?.value, "ok");

  session.destroy();
  doc.destroy();
});
