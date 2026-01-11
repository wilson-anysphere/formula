import test from "node:test";
import assert from "node:assert/strict";

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

function ensureCommentsRoot(doc) {
  doc.getMap("comments");
}

test("CollabSession undo captures comment edits (in-memory sync)", () => {
  const docA = new Y.Doc();
  const docB = new Y.Doc();

  ensureCommentsRoot(docA);
  ensureCommentsRoot(docB);

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
