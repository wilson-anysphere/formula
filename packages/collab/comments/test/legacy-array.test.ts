import * as Y from "yjs";
import { describe, expect, it } from "vitest";

import { CommentManager } from "../src/manager";
import { createYComment, getCommentsMap, migrateCommentsArrayToMap } from "../src/yjs";

function syncDocs(a: Y.Doc, b: Y.Doc): void {
  const updateA = Y.encodeStateAsUpdate(a);
  const updateB = Y.encodeStateAsUpdate(b);
  Y.applyUpdate(a, updateB);
  Y.applyUpdate(b, updateA);
}

describe("collab comments legacy array schema", () => {
  it("CommentManager works on legacy array docs without clobbering", () => {
    const doc = new Y.Doc();
    const comments = doc.getArray<Y.Map<unknown>>("comments");
    comments.push([
      createYComment({
        id: "c1",
        cellRef: "A1",
        kind: "threaded",
        author: { id: "u1", name: "Alice" },
        now: 2,
        content: "First pushed",
      }),
    ]);
    comments.push([
      createYComment({
        id: "c0",
        cellRef: "A1",
        kind: "threaded",
        author: { id: "u1", name: "Alice" },
        now: 1,
        content: "Second pushed but earlier timestamp",
      }),
    ]);

    expect(doc.share.get("comments")).toBeInstanceOf(Y.Array);

    const mgr = new CommentManager(doc);
    const listed = mgr.listAll();
    expect(listed.map((c) => c.id)).toEqual(["c0", "c1"]);
    expect(doc.share.get("comments")).toBeInstanceOf(Y.Array);

    mgr.addReply({
      commentId: "c1",
      content: "Reply",
      author: { id: "u1", name: "Alice" },
      id: "r1",
      now: 3,
    });
    mgr.setResolved({ commentId: "c1", resolved: true, now: 4 });
    mgr.setReplyContent({ commentId: "c1", replyId: "r1", content: "Edited reply", now: 5 });

    const updated = mgr.listAll().find((c) => c.id === "c1");
    expect(updated?.resolved).toBe(true);
    expect(updated?.replies).toHaveLength(1);
    expect(updated?.replies[0]?.content).toBe("Edited reply");
    expect(doc.share.get("comments")).toBeInstanceOf(Y.Array);
  });

  it("does not clobber legacy array comments when the root is an uninstantiated placeholder", () => {
    const source = new Y.Doc();
    const legacy = source.getArray<Y.Map<unknown>>("comments");
    legacy.push([
      createYComment({
        id: "c1",
        cellRef: "A1",
        kind: "threaded",
        author: { id: "u1", name: "Alice" },
        now: 1,
        content: "Hello",
      }),
    ]);

    const snapshot = Y.encodeStateAsUpdate(source);

    const target = new Y.Doc();
    Y.applyUpdate(target, snapshot);

    // Root exists but is not yet instantiated: this is the dangerous case where
    // calling `doc.getMap("comments")` would silently choose the wrong
    // constructor and make the legacy array content inaccessible.
    expect(target.share.get("comments")).toBeInstanceOf(Y.AbstractType);
    expect(target.share.get("comments")).not.toBeInstanceOf(Y.Array);
    expect(target.share.get("comments")).not.toBeInstanceOf(Y.Map);

    const mgr = new CommentManager(target);
    expect(mgr.listAll().map((c) => c.id)).toEqual(["c1"]);

    // Ensure we instantiated as an Array, not a Map.
    expect(target.share.get("comments")).toBeInstanceOf(Y.Array);
  });

  it("recovers legacy array comments after the root was instantiated as a Map (clobbered schema)", () => {
    const source = new Y.Doc();
    const legacy = source.getArray<Y.Map<unknown>>("comments");
    legacy.push([
      createYComment({
        id: "c1",
        cellRef: "A1",
        kind: "threaded",
        author: { id: "u1", name: "Alice" },
        now: 1,
        content: "Legacy",
      }),
    ]);

    const snapshot = Y.encodeStateAsUpdate(source);

    const target = new Y.Doc();
    Y.applyUpdate(target, snapshot);

    // Simulate the old bug: choosing the wrong constructor while the root is
    // still a placeholder.
    target.getMap("comments");
    expect(target.share.get("comments")).toBeInstanceOf(Y.Map);

    const mgr = new CommentManager(target);
    expect(mgr.listAll().map((c) => c.id)).toEqual(["c1"]);

    mgr.addReply({
      commentId: "c1",
      content: "Reply",
      author: { id: "u1", name: "Alice" },
      id: "r1",
      now: 2,
    });

    // Adding a new comment should use the canonical map schema even though the
    // legacy comment is stored as an array item inside the map root.
    mgr.addComment({
      cellRef: "A2",
      kind: "threaded",
      content: "New",
      author: { id: "u1", name: "Alice" },
      id: "c2",
      now: 3,
    });

    expect(mgr.listAll().map((c) => c.id)).toEqual(["c1", "c2"]);

    // Migration should normalize the legacy list item into a proper map entry.
    expect(migrateCommentsArrayToMap(target)).toBe(true);
    const map = getCommentsMap(target);
    expect(map.size).toBe(2);
    expect(map.has("c1")).toBe(true);
    expect(map.has("c2")).toBe(true);
  });

  it("converges when resolving and replying concurrently (legacy array root)", () => {
    const doc1 = new Y.Doc();
    const doc2 = new Y.Doc();

    // Force legacy schema on both docs.
    doc1.getArray("comments");
    doc2.getArray("comments");

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
    expect(doc1.share.get("comments")).toBeInstanceOf(Y.Array);
    expect(doc2.share.get("comments")).toBeInstanceOf(Y.Array);
  });

  it("migrates legacy array comments to a map without data loss", () => {
    const doc = new Y.Doc();
    doc.getArray("comments");

    const mgr = new CommentManager(doc);
    const commentId = mgr.addComment({
      cellRef: "A1",
      kind: "threaded",
      content: "Root",
      author: { id: "u1", name: "Alice" },
      id: "c1",
      now: 1,
    });
    mgr.addReply({
      commentId,
      content: "Reply",
      author: { id: "u1", name: "Alice" },
      id: "r1",
      now: 2,
    });
    mgr.setResolved({ commentId, resolved: true, now: 3 });

    const before = mgr.listAll();

    expect(migrateCommentsArrayToMap(doc)).toBe(true);
    expect(doc.share.get("comments")).toBeInstanceOf(Y.Map);

    const map = getCommentsMap(doc);
    expect(map.size).toBe(1);
    expect(map.get("c1")).toBeInstanceOf(Y.Map);

    const after = new CommentManager(doc).listAll();
    expect(after).toEqual(before);
  });
});
