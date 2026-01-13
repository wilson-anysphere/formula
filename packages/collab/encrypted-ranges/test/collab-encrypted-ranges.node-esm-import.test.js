import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when TypeScript execution isn't available.
import { createEncryptionPolicyFromDoc as createFromTs } from "../src/index.ts";

test("collab-encrypted-ranges is importable under Node ESM when executing TS sources directly", async () => {
  const mod = await import("@formula/collab-encrypted-ranges");

  assert.equal(typeof mod.EncryptedRangeManager, "function");
  assert.equal(typeof mod.createEncryptionPolicyFromDoc, "function");
  assert.equal(typeof createFromTs, "function");

  const doc = new Y.Doc();
  const mgr = new mod.EncryptedRangeManager({ doc });
  mgr.add({ sheetId: "Sheet1", startRow: 0, startCol: 0, endRow: 0, endCol: 0, keyId: "k1" });

  const policy = mod.createEncryptionPolicyFromDoc(doc);
  assert.equal(policy.shouldEncryptCell({ sheetId: "Sheet1", row: 0, col: 0 }), true);
});

