import * as Y from "yjs";
import { describe, expect, it } from "vitest";

import { CommentManager, createCommentManagerForDoc } from "../src/manager";

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
});

