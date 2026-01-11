import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { YjsBranchStore } from "../packages/versioning/branches/src/store/YjsBranchStore.js";

test("YjsBranchStore.setCurrentBranchName rejects unknown branches", async () => {
  const ydoc = new Y.Doc();
  const store = new YjsBranchStore({ ydoc });
  const docId = "doc1";
  const actor = { userId: "u1", role: "owner" };

  await store.ensureDocument(docId, actor, { sheets: {} });

  await assert.rejects(store.setCurrentBranchName(docId, "ghost"), {
    message: "Branch not found: ghost",
  });
});

