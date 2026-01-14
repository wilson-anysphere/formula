import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { REMOTE_ORIGIN, createCollabUndoService, createUndoService } from "@formula/collab-undo";

import { createCommentManagerForDoc, createCommentManagerForSession, createYComment, getCommentsRoot } from "../../comments/src/manager.ts";
import { createCollabSession } from "../src/index.ts";

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

/**
 * @param {Y.Doc} docA
 * @param {Y.Doc} docB
 */
function connectDocs(docA, docB) {
  const forwardA = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docB, update, REMOTE_ORIGIN);
  };
  const forwardB = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docA, update, REMOTE_ORIGIN);
  };

  docA.on("update", forwardA);
  docB.on("update", forwardB);

  Y.applyUpdate(docA, Y.encodeStateAsUpdate(docB), REMOTE_ORIGIN);
  Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA), REMOTE_ORIGIN);

  return () => {
    docA.off("update", forwardA);
    docB.off("update", forwardB);
  };
}

test("CollabSession undo captures comment edits when comments root is created lazily (in-memory sync)", () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();

  // Regression: callers often don't create `comments` until after the session is constructed.
  assert.equal(docA.share.get("comments"), undefined);
  assert.equal(docB.share.get("comments"), undefined);

  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA, undo: {} });
  const sessionB = createCollabSession({ doc: docB, undo: {} });
  sessionA.setPermissions({ role: "editor", userId: "u-a", rangeRestrictions: [] });
  sessionB.setPermissions({ role: "editor", userId: "u-b", rangeRestrictions: [] });

  sessionA.setPermissions({ role: "editor", rangeRestrictions: [], userId: "u-a" });
  sessionB.setPermissions({ role: "editor", rangeRestrictions: [], userId: "u-b" });

  assert.ok(docA.share.get("comments") instanceof Y.Map);
  assert.ok(docB.share.get("comments") instanceof Y.Map);

  const commentsA = createCommentManagerForSession(sessionA);
  const commentsB = createCommentManagerForSession(sessionB);

  const commentIdA = commentsA.addComment({
    id: "c_a",
    cellRef: "Sheet1:0:0",
    kind: "threaded",
    content: "from-a",
    author: { id: "u-a", name: "User A" },
    now: 1,
  });
  sessionA.undo?.stopCapturing();

  const commentIdB = commentsB.addComment({
    id: "c_b",
    cellRef: "Sheet1:0:1",
    kind: "threaded",
    content: "from-b",
    author: { id: "u-b", name: "User B" },
    now: 2,
  });

  commentsA.setCommentContent({ commentId: commentIdA, content: "from-a (edited)", now: 3 });

  const getContent = (mgr, id) => mgr.listAll().find((c) => c.id === id)?.content ?? null;

  assert.equal(getContent(commentsA, commentIdA), "from-a (edited)");
  assert.equal(getContent(commentsB, commentIdA), "from-a (edited)");
  assert.equal(getContent(commentsA, commentIdB), "from-b");
  assert.equal(getContent(commentsB, commentIdB), "from-b");

  assert.equal(sessionA.undo?.canUndo(), true);
  sessionA.undo?.undo();

  assert.equal(getContent(commentsA, commentIdA), "from-a");
  assert.equal(getContent(commentsB, commentIdA), "from-a");
  assert.equal(getContent(commentsA, commentIdB), "from-b");
  assert.equal(getContent(commentsB, commentIdB), "from-b");

  assert.equal(sessionA.undo?.canUndo(), true);
  sessionA.undo?.undo();

  assert.equal(getContent(commentsA, commentIdA), null);
  assert.equal(getContent(commentsB, commentIdA), null);
  assert.equal(getContent(commentsA, commentIdB), "from-b");
  assert.equal(getContent(commentsB, commentIdB), "from-b");

  assert.equal(sessionA.undo?.canRedo(), true);
  sessionA.undo?.redo();

  assert.equal(getContent(commentsA, commentIdA), "from-a");
  assert.equal(getContent(commentsB, commentIdA), "from-a");
  assert.equal(getContent(commentsA, commentIdB), "from-b");
  assert.equal(getContent(commentsB, commentIdB), "from-b");

  assert.equal(sessionA.undo?.canRedo(), true);
  sessionA.undo?.redo();

  assert.equal(getContent(commentsA, commentIdA), "from-a (edited)");
  assert.equal(getContent(commentsB, commentIdA), "from-a (edited)");
  assert.equal(getContent(commentsA, commentIdB), "from-b");
  assert.equal(getContent(commentsB, commentIdB), "from-b");

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession undo does not clobber legacy Array-backed comments root (in-memory sync)", () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();

  const legacyRoot = docA.getArray("comments");
  legacyRoot.push([
    createYComment({
      id: "c_legacy",
      cellRef: "Sheet1:0:0",
      kind: "threaded",
      content: "legacy",
      author: { id: "u_legacy", name: "Legacy User" },
      now: 1,
    }),
  ]);

  // Legacy doc already has the array root. The receiving doc intentionally does
  // not instantiate it yet, so `doc.share.get("comments")` may be an
  // AbstractType placeholder after sync.
  assert.ok(docA.share.get("comments") instanceof Y.Array);
  assert.equal(docB.share.get("comments"), undefined);

  const disconnect = connectDocs(docA, docB);

  assert.ok(docB.share.get("comments"));

  const sessionA = createCollabSession({ doc: docA, undo: {} });
  const sessionB = createCollabSession({ doc: docB, undo: {} });
  sessionA.setPermissions({ role: "editor", userId: "u-a", rangeRestrictions: [] });
  sessionB.setPermissions({ role: "editor", userId: "u-b", rangeRestrictions: [] });

  sessionA.setPermissions({ role: "editor", rangeRestrictions: [], userId: "u_legacy" });
  sessionB.setPermissions({ role: "editor", rangeRestrictions: [], userId: "u_legacy_remote" });

  assert.ok(docA.share.get("comments") instanceof Y.Array);
  assert.ok(docB.share.get("comments") instanceof Y.Array);

  const commentsA = createCommentManagerForSession(sessionA);
  const commentsB = createCommentManagerForSession(sessionB);

  const getContent = (mgr) => mgr.listAll().find((c) => c.id === "c_legacy")?.content ?? null;

  assert.equal(getContent(commentsA), "legacy");
  assert.equal(getContent(commentsB), "legacy");

  commentsA.setCommentContent({ commentId: "c_legacy", content: "legacy (edited)", now: 2 });

  assert.equal(getContent(commentsA), "legacy (edited)");
  assert.equal(getContent(commentsB), "legacy (edited)");

  assert.equal(sessionA.undo?.canUndo(), true);
  sessionA.undo?.undo();

  assert.equal(getContent(commentsA), "legacy");
  assert.equal(getContent(commentsB), "legacy");

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession can optionally migrate legacy Array-backed comments root to canonical Map schema", async () => {
  const doc = new Y.Doc();

  const legacyRoot = doc.getArray("comments");
  legacyRoot.push([
    createYComment({
      id: "c_legacy",
      cellRef: "Sheet1:0:0",
      kind: "threaded",
      content: "legacy",
      author: { id: "u_legacy", name: "Legacy User" },
      now: 1,
    }),
  ]);

  assert.ok(doc.share.get("comments") instanceof Y.Array);

  const session = createCollabSession({
    doc,
    undo: {},
    comments: { migrateLegacyArrayToMap: true },
  });

  session.setPermissions({ role: "editor", rangeRestrictions: [], userId: "u_legacy" });

  // Migration is scheduled after hydration in a microtask.
  await new Promise((resolve) => queueMicrotask(resolve));

  assert.ok(doc.share.get("comments") instanceof Y.Map);

  const comments = createCommentManagerForSession(session);
  const getContent = () => comments.listAll().find((c) => c.id === "c_legacy")?.content ?? null;

  assert.equal(getContent(), "legacy");

  comments.setCommentContent({ commentId: "c_legacy", content: "legacy (edited)", now: 2 });
  assert.equal(getContent(), "legacy (edited)");

  // Migration replaces the comments root; ensure it remains undoable.
  assert.equal(session.undo?.canUndo(), true);
  session.undo?.undo();
  assert.equal(getContent(), "legacy");

  session.destroy();
  doc.destroy();
});

test("CollabSession comment migration is gated by comment permissions (viewer does not migrate)", async () => {
  const doc = new Y.Doc();

  const legacyRoot = doc.getArray("comments");
  legacyRoot.push([
    createYComment({
      id: "c_legacy",
      cellRef: "Sheet1:0:0",
      kind: "threaded",
      content: "legacy",
      author: { id: "u_legacy", name: "Legacy User" },
      now: 1,
    }),
  ]);

  assert.ok(doc.share.get("comments") instanceof Y.Array);

  const session = createCollabSession({
    doc,
    undo: {},
    comments: { migrateLegacyArrayToMap: true },
  });

  session.setPermissions({ role: "viewer", rangeRestrictions: [], userId: "u_viewer" });

  // Migration is scheduled after hydration in a microtask, but should be skipped for viewers.
  await new Promise((resolve) => queueMicrotask(resolve));
  assert.ok(doc.share.get("comments") instanceof Y.Array);

  // If permissions change to a role that can comment, migration should run.
  session.setPermissions({ role: "commenter", rangeRestrictions: [], userId: "u_viewer" });
  await new Promise((resolve) => queueMicrotask(resolve));

  assert.ok(doc.share.get("comments") instanceof Y.Map);

  const comments = createCommentManagerForSession(session);
  const getContent = () => comments.listAll().find((c) => c.id === "c_legacy")?.content ?? null;
  assert.equal(getContent(), "legacy");

  session.destroy();
  doc.destroy();
});

test("CollabSession undo captures comment edits when comments root was created by a different Yjs instance (CJS applyUpdate)", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const comments = remote.getMap("comments");
  remote.transact(() => {
    const comment = new Ycjs.Map();
    comment.set("id", "c1");
    comment.set("cellRef", "Sheet1:0:0");
    comment.set("kind", "threaded");
    comment.set("authorId", "u1");
    comment.set("authorName", "Alice");
    comment.set("createdAt", 1);
    comment.set("updatedAt", 1);
    comment.set("resolved", false);
    comment.set("content", "from-cjs");
    comment.set("mentions", []);
    comment.set("replies", new Ycjs.Array());
    comments.set("c1", comment);
  });

  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Apply update via the CJS build to simulate y-websocket applying updates.
  Ycjs.applyUpdate(doc, update, REMOTE_ORIGIN);

  // Ensure the root exists but isn't necessarily `instanceof Y.Map` in this module.
  assert.ok(doc.share.get("comments"));

  const session = createCollabSession({ doc, undo: {} });
  session.setPermissions({ role: "editor", userId: "u1", rangeRestrictions: [] });
  const commentsMgr = createCommentManagerForSession(session);

  assert.deepEqual(commentsMgr.listAll().map((c) => c.id), ["c1"]);
  assert.equal(commentsMgr.listAll()[0]?.content, "from-cjs");

  commentsMgr.setCommentContent({ commentId: "c1", content: "edited", now: 2 });
  assert.equal(commentsMgr.listAll()[0]?.content, "edited");
  assert.equal(session.undo?.canUndo(), true);

  session.undo?.undo();
  assert.equal(commentsMgr.listAll()[0]?.content, "from-cjs");

  session.destroy();
  doc.destroy();
});

test("CollabSession undo captures comment edits when foreign Yjs constructors are renamed", () => {
  const Ycjs = requireYjsCjs();
  const remote = new Ycjs.Doc();
  const comments = remote.getMap("comments");
  remote.transact(() => {
    const comment = new Ycjs.Map();
    comment.set("id", "c1");
    comment.set("cellRef", "Sheet1:0:0");
    comment.set("kind", "threaded");
    comment.set("authorId", "u1");
    comment.set("authorName", "Alice");
    comment.set("createdAt", 1);
    comment.set("updatedAt", 1);
    comment.set("resolved", false);
    comment.set("content", "from-cjs");
    comment.set("mentions", []);
    comment.set("replies", new Ycjs.Array());
    comments.set("c1", comment);
  });

  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Pre-create the root in the ESM instance so applying the update only introduces
  // foreign nested maps/arrays (not a foreign root placeholder).
  doc.getMap("comments");
  // Apply update via the CJS build to simulate y-websocket applying updates.
  Ycjs.applyUpdate(doc, update, REMOTE_ORIGIN);

  const root = doc.getMap("comments");
  const foreignComment = root.get("c1");
  assert.ok(foreignComment);
  assert.equal(
    foreignComment instanceof Y.Map,
    false,
    "expected nested comment to be a foreign Y.Map instance (not instanceof the ESM Y.Map)"
  );

  // Simulate bundler-renamed constructors without mutating global `yjs` state
  // (which can cause cross-test interference under concurrency).
  class RenamedMap extends foreignComment.constructor {}
  Object.setPrototypeOf(foreignComment, RenamedMap.prototype);

  const replies = foreignComment.get("replies");
  if (replies) {
    class RenamedArray extends replies.constructor {}
    Object.setPrototypeOf(replies, RenamedArray.prototype);
  }

  const session = createCollabSession({ doc, undo: {} });
  session.setPermissions({ role: "editor", userId: "u1", rangeRestrictions: [] });
  const commentsMgr = createCommentManagerForSession(session);

  assert.deepEqual(commentsMgr.listAll().map((c) => c.id), ["c1"]);
  assert.equal(commentsMgr.listAll()[0]?.content, "from-cjs");

  commentsMgr.setCommentContent({ commentId: "c1", content: "edited", now: 2 });
  assert.equal(commentsMgr.listAll()[0]?.content, "edited");
  assert.equal(session.undo?.canUndo(), true);

  session.undo?.undo();
  assert.equal(commentsMgr.listAll()[0]?.content, "from-cjs");

  session.destroy();
  doc.destroy();
});

test("CollabSession undo captures comment edits when comments root is a legacy array created by a different Yjs instance (CJS applyUpdate)", () => {
  const require = createRequire(import.meta.url);
  // eslint-disable-next-line import/no-named-as-default-member
  const Ycjs = require("yjs");

  const remote = new Ycjs.Doc();
  const comments = remote.getArray("comments");
  remote.transact(() => {
    const comment = new Ycjs.Map();
    comment.set("id", "c1");
    comment.set("cellRef", "Sheet1:0:0");
    comment.set("kind", "threaded");
    comment.set("authorId", "u1");
    comment.set("authorName", "Alice");
    comment.set("createdAt", 1);
    comment.set("updatedAt", 1);
    comment.set("resolved", false);
    comment.set("content", "from-cjs");
    comment.set("mentions", []);
    comment.set("replies", new Ycjs.Array());
    comments.push([comment]);
  });

  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Apply update via the CJS build to simulate y-websocket applying updates.
  Ycjs.applyUpdate(doc, update, REMOTE_ORIGIN);

  assert.ok(doc.share.get("comments"));

  const session = createCollabSession({ doc, undo: {} });
  session.setPermissions({ role: "editor", userId: "u1", rangeRestrictions: [] });
  const commentsMgr = createCommentManagerForSession(session);

  assert.deepEqual(commentsMgr.listAll().map((c) => c.id), ["c1"]);
  assert.equal(commentsMgr.listAll()[0]?.content, "from-cjs");

  commentsMgr.setCommentContent({ commentId: "c1", content: "edited", now: 2 });
  assert.equal(commentsMgr.listAll()[0]?.content, "edited");
  assert.equal(session.undo?.canUndo(), true);

  session.undo?.undo();
  assert.equal(commentsMgr.listAll()[0]?.content, "from-cjs");

  session.destroy();
  doc.destroy();
});

test("Binder-origin collaborative undo captures comment add/edit when using a tracked transact wrapper (desktop-style)", () => {
  const doc = new Y.Doc();

  // Desktop's binder-origin undo service tracks only transactions that run with
  // the binder origin. If comment edits do not use that transact wrapper they
  // won't be captured by UndoManager.
  const binderOrigin = { type: "document-controller:binder" };

  // Ensure the comments root exists *before* creating the UndoManager so it's
  // included in its scope (mirrors desktop's undo scope construction).
  const commentsRoot = doc.getMap("comments");

  const undo = createCollabUndoService({
    doc,
    scope: commentsRoot,
    origin: binderOrigin,
  });

  const comments = createCommentManagerForDoc({ doc, transact: undo.transact });

  const commentId = comments.addComment({
    id: "c1",
    cellRef: "Sheet1:0:0",
    kind: "threaded",
    content: "Hello",
    author: { id: "u1", name: "Alice" },
    now: 1,
  });

  // Split the add/edit into separate undo stack entries.
  undo.stopCapturing();

  comments.setCommentContent({ commentId, content: "Hello (edited)", now: 2 });
  undo.stopCapturing();

  comments.addReply({
    commentId,
    id: "r1",
    content: "First reply",
    author: { id: "u1", name: "Alice" },
    now: 3,
  });
  undo.stopCapturing();

  comments.setReplyContent({ commentId, replyId: "r1", content: "First reply (edited)", now: 4 });
  undo.stopCapturing();

  comments.setResolved({ commentId, resolved: true, now: 5 });

  const get = () => comments.listAll().find((c) => c.id === commentId) ?? null;

  assert.equal(get()?.content ?? null, "Hello (edited)");
  assert.equal(get()?.replies.length ?? 0, 1);
  assert.equal(get()?.replies[0]?.content ?? null, "First reply (edited)");
  assert.equal(get()?.resolved ?? null, true);
  assert.equal(undo.canUndo(), true);

  // Undo resolve.
  undo.undo();
  assert.equal(get()?.resolved ?? null, false);

  // Undo reply edit.
  assert.equal(undo.canUndo(), true);
  undo.undo();
  assert.equal(get()?.replies.length ?? 0, 1);
  assert.equal(get()?.replies[0]?.content ?? null, "First reply");

  // Undo reply add.
  assert.equal(undo.canUndo(), true);
  undo.undo();
  assert.equal(get()?.replies.length ?? 0, 0);

  // Undo edit.
  assert.equal(undo.canUndo(), true);
  undo.undo();
  assert.equal(get()?.content ?? null, "Hello");

  // Undo add.
  assert.equal(undo.canUndo(), true);
  undo.undo();
  assert.equal(get(), null);

  // Redo add + edit + reply + resolve.
  assert.equal(undo.canRedo(), true);
  undo.redo();
  assert.equal(get()?.content ?? null, "Hello");

  assert.equal(undo.canRedo(), true);
  undo.redo();
  assert.equal(get()?.content ?? null, "Hello (edited)");

  assert.equal(undo.canRedo(), true);
  undo.redo();
  assert.equal(get()?.replies.length ?? 0, 1);
  assert.equal(get()?.replies[0]?.content ?? null, "First reply");

  assert.equal(undo.canRedo(), true);
  undo.redo();
  assert.equal(get()?.replies.length ?? 0, 1);
  assert.equal(get()?.replies[0]?.content ?? null, "First reply (edited)");

  assert.equal(undo.canRedo(), true);
  undo.redo();
  assert.equal(get()?.resolved ?? null, true);

  doc.destroy();
});

test("Binder-origin undo captures comment edits when comments root is added to UndoManager scope lazily (desktop-style)", () => {
  const doc = new Y.Doc();

  const binderOrigin = { type: "document-controller:binder" };

  // Desktop creates the UndoManager before the comments root exists in many cases.
  // Start with a scope that does *not* include comments.
  const cellsRoot = doc.getMap("cells");
  const undoService = createUndoService({ mode: "collab", doc, scope: [cellsRoot], origin: binderOrigin });

  // Find the underlying UndoManager (exposed via localOrigins).
  const undoManager = Array.from(undoService.localOrigins ?? []).find((origin) => origin instanceof Y.UndoManager);
  assert.ok(undoManager);

  let commentsScopeAdded = false;
  const ensureCommentsScope = () => {
    if (commentsScopeAdded) return;
    const root = getCommentsRoot(doc);
    undoManager.addToScope(root.kind === "map" ? root.map : root.array);
    commentsScopeAdded = true;
  };

  const comments = createCommentManagerForDoc({
    doc,
    transact: (fn) => {
      ensureCommentsScope();
      undoService.transact(fn);
    },
  });

  const commentId = comments.addComment({
    id: "c1",
    cellRef: "Sheet1:0:0",
    kind: "threaded",
    content: "hello",
    author: { id: "u1", name: "Alice" },
    now: 1,
  });
  undoService.stopCapturing();

  comments.setCommentContent({ commentId, content: "hello (edited)", now: 2 });
  assert.equal(comments.listAll().find((c) => c.id === commentId)?.content ?? null, "hello (edited)");

  assert.equal(undoService.canUndo(), true);
  undoService.undo();
  assert.equal(comments.listAll().find((c) => c.id === commentId)?.content ?? null, "hello");

  assert.equal(undoService.canUndo(), true);
  undoService.undo();
  assert.equal(comments.listAll().find((c) => c.id === commentId)?.content ?? null, null);

  doc.destroy();
});

test("Binder-origin undo captures comment edits when legacy Array-backed comments root is added to UndoManager scope lazily (desktop-style)", () => {
  const doc = new Y.Doc();

  const binderOrigin = { type: "document-controller:binder" };

  // Start with an undo scope that does not include comments.
  const cellsRoot = doc.getMap("cells");
  const undoService = createUndoService({ mode: "collab", doc, scope: [cellsRoot], origin: binderOrigin });

  const undoManager = Array.from(undoService.localOrigins ?? []).find((origin) => origin instanceof Y.UndoManager);
  assert.ok(undoManager);

  // Simulate a legacy doc that already uses an Array-backed comments root, but
  // where the UndoManager was created before the comments root was added to its scope.
  doc.getArray("comments"); // force legacy array schema

  let commentsScopeAdded = false;
  const ensureCommentsScope = () => {
    if (commentsScopeAdded) return;
    const root = getCommentsRoot(doc);
    undoManager.addToScope(root.kind === "map" ? root.map : root.array);
    commentsScopeAdded = true;
  };

  const comments = createCommentManagerForDoc({
    doc,
    transact: (fn) => {
      ensureCommentsScope();
      undoService.transact(fn);
    },
  });

  const commentId = comments.addComment({
    id: "c1",
    cellRef: "Sheet1:0:0",
    kind: "threaded",
    content: "hello",
    author: { id: "u1", name: "Alice" },
    now: 1,
  });
  undoService.stopCapturing();

  comments.setCommentContent({ commentId, content: "hello (edited)", now: 2 });
  assert.equal(comments.listAll().find((c) => c.id === commentId)?.content ?? null, "hello (edited)");

  assert.equal(undoService.canUndo(), true);
  undoService.undo();
  assert.equal(comments.listAll().find((c) => c.id === commentId)?.content ?? null, "hello");

  assert.equal(undoService.canUndo(), true);
  undoService.undo();
  assert.equal(comments.listAll().find((c) => c.id === commentId)?.content ?? null, null);

  doc.destroy();
});

test("Binder-origin undo captures comment edits when nested comment maps were created by a different Yjs instance (CJS applyUpdate)", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const comments = remote.getMap("comments");
  remote.transact(() => {
    const comment = new Ycjs.Map();
    comment.set("id", "c1");
    comment.set("cellRef", "Sheet1:0:0");
    comment.set("kind", "threaded");
    comment.set("authorId", "u1");
    comment.set("authorName", "Alice");
    comment.set("createdAt", 1);
    comment.set("updatedAt", 1);
    comment.set("resolved", false);
    comment.set("content", "from-cjs");
    comment.set("mentions", []);
    comment.set("replies", new Ycjs.Array());
    comments.set("c1", comment);
  });

  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Ensure the root exists in the ESM build so the update only introduces foreign
  // nested comment maps (not a foreign `comments` root).
  const commentsRoot = doc.getMap("comments");
  Ycjs.applyUpdate(doc, update, REMOTE_ORIGIN);

  // Verify the nested comment map is foreign.
  const nested = commentsRoot.get("c1");
  assert.ok(nested);
  assert.equal(nested instanceof Y.Map, false);

  const binderOrigin = { type: "document-controller:binder" };
  const undoService = createUndoService({ mode: "collab", doc, scope: [commentsRoot], origin: binderOrigin });
  const mgr = createCommentManagerForDoc({ doc, transact: undoService.transact });

  assert.equal(mgr.listAll()[0]?.content, "from-cjs");

  mgr.setCommentContent({ commentId: "c1", content: "edited", now: 2 });
  assert.equal(mgr.listAll()[0]?.content, "edited");
  assert.equal(undoService.canUndo(), true);

  undoService.undo();
  assert.equal(mgr.listAll()[0]?.content, "from-cjs");

  doc.destroy();
});

test("Binder-origin undo captures comment edits when comments root itself was created by a different Yjs instance (CJS applyUpdate)", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const comments = remote.getMap("comments");
  remote.transact(() => {
    const comment = new Ycjs.Map();
    comment.set("id", "c1");
    comment.set("cellRef", "Sheet1:0:0");
    comment.set("kind", "threaded");
    comment.set("authorId", "u1");
    comment.set("authorName", "Alice");
    comment.set("createdAt", 1);
    comment.set("updatedAt", 1);
    comment.set("resolved", false);
    comment.set("content", "from-cjs");
    comment.set("mentions", []);
    comment.set("replies", new Ycjs.Array());
    comments.set("c1", comment);
  });
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  // Apply update via the CJS build to simulate y-websocket applying updates.
  Ycjs.applyUpdate(doc, update, REMOTE_ORIGIN);

  // Root exists but may be a generic placeholder type after applying updates via
  // a different Yjs module instance (CJS vs ESM).
  const existing = doc.share.get("comments");
  assert.ok(existing);
  assert.equal(existing instanceof Y.Map, false);

  const root = getCommentsRoot(doc);
  if (root.kind !== "map") {
    throw new Error("Expected canonical comments root schema in this test");
  }
  const nested = root.map.get("c1");
  assert.ok(nested);
  // The comments root can be instantiated locally via `getCommentsRoot`, but nested
  // comment objects may still be created by a different Yjs module instance when
  // updates are applied via the CJS build (e.g. y-websocket).
  assert.equal(nested instanceof Y.Map, false);

  const binderOrigin = { type: "document-controller:binder" };
  const undoService = createUndoService({ mode: "collab", doc, scope: [root.map], origin: binderOrigin });
  const mgr = createCommentManagerForDoc({ doc, transact: undoService.transact });

  assert.equal(mgr.listAll()[0]?.content ?? null, "from-cjs");
  mgr.setCommentContent({ commentId: "c1", content: "edited", now: 2 });
  assert.equal(mgr.listAll()[0]?.content ?? null, "edited");
  assert.equal(undoService.canUndo(), true);

  undoService.undo();
  assert.equal(mgr.listAll()[0]?.content ?? null, "from-cjs");

  doc.destroy();
});

test("CollabSession undo captures comment edits when comments root is a foreign AbstractType placeholder (CJS Doc.get)", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();
  // Simulate another Yjs module instance calling `Doc.get(name)` (defaulting to
  // AbstractType) on this doc. This creates a foreign `AbstractType` placeholder
  // that would cause `doc.getMap("comments")` to throw from the ESM build.
  Ycjs.Doc.prototype.get.call(doc, "comments");

  assert.ok(doc.share.get("comments"));
  // Note: other tests in this file create UndoManagers that patch foreign Yjs
  // prototype chains so foreign types pass `instanceof Y.AbstractType` checks.
  // We therefore can't reliably assert `instanceof Y.AbstractType === false`
  // here. Instead, assert the symptom we care about: `doc.getMap("comments")`
  // would throw "different constructor" without session normalization.
  assert.throws(
    () => doc.getMap("comments"),
    /Type with the name comments has already been defined with a different constructor/,
    "expected doc.getMap(\"comments\") to throw when the root is a foreign AbstractType placeholder"
  );

  const session = createCollabSession({ doc, undo: {} });
  session.setPermissions({ role: "editor", userId: "u1", rangeRestrictions: [] });

  // The session should normalize the placeholder so the canonical root is usable.
  assert.ok(doc.share.get("comments") instanceof Y.Map);

  const mgr = createCommentManagerForSession(session);

  const commentId = mgr.addComment({
    id: "c1",
    cellRef: "Sheet1:0:0",
    kind: "threaded",
    content: "Hello",
    author: { id: "u1", name: "Alice" },
    now: 1,
  });
  session.undo?.stopCapturing();

  mgr.setCommentContent({ commentId, content: "Hello (edited)", now: 2 });
  assert.equal(mgr.listAll()[0]?.content ?? null, "Hello (edited)");
  assert.equal(session.undo?.canUndo(), true);

  session.undo?.undo();
  assert.equal(mgr.listAll()[0]?.content ?? null, "Hello");

  session.destroy();
  doc.destroy();
});
