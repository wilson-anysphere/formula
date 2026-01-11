import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { YjsBranchStore } from "../packages/versioning/branches/src/store/YjsBranchStore.js";

test("YjsBranchStore.ensureDocument rejects docId mismatch for the same Y.Doc", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsBranchStore({ ydoc });
  const actor = { userId: "u1", role: "owner" };

  await store.ensureDocument("doc1", actor, { sheets: {} });

  await assert.rejects(store.ensureDocument("doc2", actor, { sheets: {} }), {
    message: /docId mismatch/,
  });
});

