import { createRequire } from "node:module";
import * as Y from "yjs";
import { describe, expect, it } from "vitest";

import { CommentManager, createCommentManagerForDoc } from "../src/manager";

function requireYjsCjs(): typeof import("yjs") {
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

describe("CommentManager permissions guard", () => {
  it("rejects addComment when canComment=false (and does not instantiate the root)", () => {
    const doc = new Y.Doc();
    const mgr = new CommentManager(doc, { canComment: () => false });

    expect(() =>
      mgr.addComment({
        cellRef: "A1",
        kind: "threaded",
        content: "Hello",
        author: { id: "u1", name: "Alice" },
      }),
    ).toThrowError("Permission denied: cannot comment");

    // Guard should fire before `getCommentsRoot()` so viewers don't accidentally
    // create the comments root in pre-hydration docs.
    expect(doc.share.get("comments")).toBe(undefined);
  });

  it("fails closed when canComment throws (and does not instantiate the root)", () => {
    const doc = new Y.Doc();
    const mgr = new CommentManager(doc, {
      canComment: () => {
        throw new Error("boom");
      },
    });

    expect(() =>
      mgr.addComment({
        cellRef: "A1",
        kind: "threaded",
        content: "Hello",
        author: { id: "u1", name: "Alice" },
      }),
    ).toThrowError("Permission denied: cannot comment");

    expect(doc.share.get("comments")).toBe(undefined);
  });

  it("rejects all mutating operations but still allows reading existing threads", () => {
    const doc = new Y.Doc();
    const writer = new CommentManager(doc);
    const commentId = writer.addComment({
      cellRef: "A1",
      kind: "threaded",
      content: "Hello",
      author: { id: "u1", name: "Alice" },
    });
    const replyId = writer.addReply({
      commentId,
      content: "Reply",
      author: { id: "u2", name: "Bob" },
    });
    const snapshot = writer.listAll();

    const viewer = new CommentManager(doc, { canComment: () => false });

    // Read paths remain available.
    expect(viewer.listAll()).toEqual(snapshot);
    expect(viewer.listForCell("A1")).toEqual(snapshot);

    // Mutations are blocked.
    expect(() =>
      viewer.addReply({
        commentId,
        content: "nope",
        author: { id: "u3", name: "Mallory" },
      }),
    ).toThrowError("Permission denied: cannot comment");
    expect(() => viewer.setResolved({ commentId, resolved: true })).toThrowError("Permission denied: cannot comment");
    expect(() =>
      viewer.setCommentContent({
        commentId,
        content: "edit",
      }),
    ).toThrowError("Permission denied: cannot comment");
    expect(() =>
      viewer.setReplyContent({
        commentId,
        replyId,
        content: "edit",
      }),
    ).toThrowError("Permission denied: cannot comment");

    // Ensure no changes were applied.
    expect(writer.listAll()).toEqual(snapshot);
  });

  it("propagates canComment through createCommentManagerForDoc", () => {
    const doc = new Y.Doc();
    const mgr = createCommentManagerForDoc({
      doc,
      transact: (fn) => doc.transact(fn),
      canComment: () => false,
    });

    expect(() =>
      mgr.addComment({
        cellRef: "A1",
        kind: "threaded",
        content: "Hello",
        author: { id: "u1", name: "Alice" },
      }),
    ).toThrowError("Permission denied: cannot comment");
  });

  it("does not normalize foreign nested comment maps when canComment=false", () => {
    const Ycjs = requireYjsCjs();

    const remote = new Ycjs.Doc();
    const comments = remote.getMap("comments");
    remote.transact(() => {
      const comment = new Ycjs.Map();
      comment.set("id", "c1");
      comment.set("cellRef", "A1");
      comment.set("kind", "threaded");
      comment.set("authorId", "u1");
      comment.set("authorName", "Alice");
      comment.set("createdAt", 1);
      comment.set("updatedAt", 1);
      comment.set("resolved", false);
      comment.set("content", "Hello");
      comment.set("mentions", []);
      comment.set("replies", new Ycjs.Array());
      comments.set("c1", comment);
    });

    const update = Ycjs.encodeStateAsUpdate(remote);

    const doc = new Y.Doc();
    // Force the comments root to be created by the *local* Yjs instance so that
    // applying a foreign update can result in foreign nested values under a local root.
    doc.getMap("comments");
    Ycjs.applyUpdate(doc, update);

    const root = doc.getMap("comments");
    const before = root.get("c1");
    expect(before).toBeTruthy();
    // The nested comment map should be from the foreign (CJS) Yjs build.
    expect(before).not.toBeInstanceOf(Y.Map);

    let updates = 0;
    doc.on("update", () => {
      updates += 1;
    });

    // Viewers should not cause normalization transactions that create Yjs updates.
    createCommentManagerForDoc({
      doc,
      transact: (fn) => doc.transact(fn),
      canComment: () => false,
    });

    expect(updates).toBe(0);
    expect(root.get("c1")).not.toBeInstanceOf(Y.Map);

    // With comment permissions, normalization should run (so foreign maps become local).
    createCommentManagerForDoc({ doc, transact: (fn) => doc.transact(fn), canComment: () => true });
    expect(root.get("c1")).toBeInstanceOf(Y.Map);
    expect(updates).toBeGreaterThan(0);
  });
});
