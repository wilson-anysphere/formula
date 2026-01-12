import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when `--experimental-strip-types` is not available.
import { getCommentsRoot as getCommentsRootFromTs } from "../../../packages/collab/comments/src/index.ts";

test("collab-comments is importable under Node ESM when executing TS sources (strip-types)", async () => {
  const mod = await import("@formula/collab-comments");
  const Y = await import("yjs");

  assert.equal(typeof mod.getCommentsRoot, "function");
  assert.equal(typeof mod.createYComment, "function");
  assert.equal(typeof getCommentsRootFromTs, "function");

  const doc = new Y.Doc();
  const root = mod.getCommentsRoot(doc);
  assert.equal(root.kind, "map");
  assert.ok(root.map);
  assert.equal(typeof root.map.get, "function");
});

