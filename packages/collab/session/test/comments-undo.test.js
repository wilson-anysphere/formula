import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { REMOTE_ORIGIN } from "@formula/collab-undo";

import { createCommentManagerForSession } from "../../comments/src/manager.ts";
import { createCollabSession } from "../src/index.ts";

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

test("CollabSession undo captures comment edits when comments root is a legacy array (in-memory sync)", () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();

  // Legacy docs stored comments as a Y.Array. The session must not clobber this
  // by eagerly calling `doc.getMap("comments")` when building the undo scope.
  docA.getArray("comments");
  docB.getArray("comments");
  assert.ok(docA.share.get("comments") instanceof Y.Array);
  assert.ok(docB.share.get("comments") instanceof Y.Array);

  const disconnect = connectDocs(docA, docB);

  const sessionA = createCollabSession({ doc: docA, undo: {} });
  const sessionB = createCollabSession({ doc: docB, undo: {} });

  assert.ok(docA.share.get("comments") instanceof Y.Array);
  assert.ok(docB.share.get("comments") instanceof Y.Array);

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

  sessionA.destroy();
  sessionB.destroy();
  disconnect();
  docA.destroy();
  docB.destroy();
});

test("CollabSession undo captures comment edits when comments root was created by a different Yjs instance (CJS applyUpdate)", () => {
  const require = createRequire(import.meta.url);
  // eslint-disable-next-line import/no-named-as-default-member
  const Ycjs = require("yjs");

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
