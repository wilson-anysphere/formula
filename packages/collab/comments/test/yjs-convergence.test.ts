import * as Y from "yjs";
import { describe, expect, it } from "vitest";

import { CommentManager } from "../src/manager";

function syncDocs(a: Y.Doc, b: Y.Doc): void {
  const updateA = Y.encodeStateAsUpdate(a);
  const updateB = Y.encodeStateAsUpdate(b);
  Y.applyUpdate(a, updateB);
  Y.applyUpdate(b, updateA);
}

describe("collab comments", () => {
  it("converges when resolving and replying concurrently", () => {
    const doc1 = new Y.Doc();
    const doc2 = new Y.Doc();

    const mgr1 = new CommentManager(doc1);
    const mgr2 = new CommentManager(doc2);

    const commentId = mgr1.addComment({
      cellRef: "A1",
      kind: "threaded",
      content: "Root",
      author: { id: "u1", name: "Alice" },
      id: "c1",
      now: 1,
    });

    syncDocs(doc1, doc2);

    mgr1.addReply({
      commentId,
      content: "Reply",
      author: { id: "u1", name: "Alice" },
      id: "r1",
      now: 2,
    });
    mgr2.setResolved({ commentId, resolved: true, now: 3 });

    syncDocs(doc1, doc2);

    expect(mgr1.listAll()).toEqual(mgr2.listAll());
    expect(mgr1.listAll()[0]?.resolved).toBe(true);
    expect(mgr1.listAll()[0]?.replies).toHaveLength(1);
  });
});

