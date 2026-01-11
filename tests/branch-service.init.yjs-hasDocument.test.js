import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { BranchService } from "../packages/versioning/branches/src/BranchService.js";
import { YjsBranchStore } from "../packages/versioning/branches/src/store/YjsBranchStore.js";

test("BranchService.init uses store.hasDocument so existing Yjs docs don't require owner/admin", async () => {
  const docId = "doc1";
  const owner = { userId: "u-owner", role: "owner" };
  const viewer = { userId: "u-viewer", role: "viewer" };

  const ydoc = new Y.Doc();
  const store = new YjsBranchStore({ ydoc });

  const serviceOwner = new BranchService({ docId, store });
  await serviceOwner.init(owner, { sheets: {} });

  // Corrupt the doc by removing main; ensureDocument should repair without requiring an owner/admin.
  ydoc.getMap("branching:branches").delete("main");

  const serviceViewer = new BranchService({ docId, store });
  await serviceViewer.init(viewer, { sheets: {} });

  assert.ok(await store.getBranch(docId, "main"));
});

test("BranchService.init still requires owner/admin for brand new Yjs docs", async () => {
  const docId = "doc2";
  const viewer = { userId: "u-viewer", role: "viewer" };

  const ydoc = new Y.Doc();
  const store = new YjsBranchStore({ ydoc });
  const service = new BranchService({ docId, store });

  await assert.rejects(service.init(viewer, { sheets: {} }), {
    message: "init requires owner/admin permissions (role=viewer)",
  });
});

