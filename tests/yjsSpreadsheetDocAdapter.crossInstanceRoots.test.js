import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { createYjsSpreadsheetDocAdapter } from "../packages/versioning/src/yjs/yjsSpreadsheetDocAdapter.js";

function requireYjsCjs() {
  const require = createRequire(import.meta.url);
  const prevError = console.error;
  console.error = (...args) => {
    if (typeof args[0] === "string" && args[0].startsWith("Yjs was already imported.")) return;
    prevError(...args);
  };
  try {
    // eslint-disable-next-line import/no-named-as-default-member
    return require("yjs");
  } finally {
    console.error = prevError;
  }
}

test("Yjs doc adapter: encode/apply work when roots were created by a different Yjs instance (CJS applyUpdate)", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Apply an update using the CJS build of Yjs. This can result in root types
  // (and nested Y.Maps) that fail `instanceof` checks against the ESM build.
  const remote = new Ycjs.Doc();
  const comments = remote.getMap("comments");
  remote.transact(() => {
    const comment = new Ycjs.Map();
    comment.set("id", "c1");
    comment.set("content", "Hello");
    comments.set("c1", comment);
  });
  const update = Ycjs.encodeStateAsUpdate(remote);
  Ycjs.applyUpdate(doc, update);

  // Create an excluded root so the adapter takes the "filtered snapshot" path
  // (clone roots into a new doc instead of encoding the whole doc).
  doc.getMap("versions").set("v-local", 1);

  const adapter = createYjsSpreadsheetDocAdapter(doc, { excludeRoots: ["versions", "versionsMeta"] });

  const snapshot = adapter.encodeState();
  const replay = new Y.Doc();
  Y.applyUpdate(replay, snapshot);

  assert.equal(replay.share.has("versions"), false, "expected versions root to be excluded");
  assert.equal(replay.getMap("comments").get("c1")?.get("content"), "Hello");

  // Mutate and restore; this exercises root access in applyState().
  const commentsRoot = /** @type {any} */ (doc.share.get("comments"));
  assert.ok(commentsRoot && typeof commentsRoot.get === "function", "expected comments root to exist");
  commentsRoot.get("c1")?.set("content", "Changed");
  adapter.applyState(snapshot);
  const restoredRoot = /** @type {any} */ (doc.share.get("comments"));
  assert.equal(restoredRoot?.get?.("c1")?.get?.("content"), "Hello");
});
