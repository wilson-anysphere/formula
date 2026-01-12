import { createRequire } from "node:module";

import * as Y from "yjs";
import { describe, expect, it } from "vitest";

import { CommentManager } from "../src/manager";
import { getCommentsRoot, migrateCommentsArrayToMap } from "../src/yjs";

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

describe("collab comments cross-instance Yjs roots (ESM doc + CJS applyUpdate)", () => {
  it("reads and mutates a comments map root created by a different Yjs instance", () => {
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
    // Apply using the CJS build; this can produce roots/nested types that fail
    // `instanceof` checks against the ESM build.
    Ycjs.applyUpdate(doc, update);

    const mgr = new CommentManager(doc);
    expect(mgr.listAll().map((c) => c.id)).toEqual(["c1"]);
    expect(mgr.listAll()[0]?.content).toBe("Hello");

    mgr.addReply({
      commentId: "c1",
      content: "Reply",
      author: { id: "u1", name: "Alice" },
      id: "r1",
      now: 2,
    });
    mgr.setResolved({ commentId: "c1", resolved: true, now: 3 });

    const updated = mgr.listAll()[0]!;
    expect(updated.resolved).toBe(true);
    expect(updated.replies).toHaveLength(1);
    expect(updated.replies[0]?.content).toBe("Reply");
  });

  it("reads and appends to a legacy comments array root created by a different Yjs instance", () => {
    const Ycjs = requireYjsCjs();

    const remote = new Ycjs.Doc();
    const comments = remote.getArray("comments");

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
      comments.push([comment]);
    });

    const update = Ycjs.encodeStateAsUpdate(remote);

    const doc = new Y.Doc();
    Ycjs.applyUpdate(doc, update);

    const mgr = new CommentManager(doc);
    expect(mgr.listAll().map((c) => c.id)).toEqual(["c1"]);

    mgr.addComment({
      cellRef: "A2",
      kind: "threaded",
      content: "Added locally",
      author: { id: "u1", name: "Alice" },
      id: "c2",
      now: 2,
    });

    expect(mgr.listAll().map((c) => c.id).sort()).toEqual(["c1", "c2"]);
  });

  it("can write to a doc created by a different Yjs module instance (CJS Doc)", () => {
    const Ycjs = requireYjsCjs();

    // Doc + root types are from the CJS build.
    const doc = new Ycjs.Doc();

    // CommentManager uses the ESM build.
    const mgr = new CommentManager(doc);

    mgr.addComment({
      cellRef: "A1",
      kind: "threaded",
      content: "Hello",
      author: { id: "u1", name: "Alice" },
      id: "c1",
      now: 1,
    });

    mgr.addReply({
      commentId: "c1",
      content: "Reply",
      author: { id: "u1", name: "Alice" },
      id: "r1",
      now: 2,
    });

    mgr.setResolved({ commentId: "c1", resolved: true, now: 3 });

    const listed = mgr.listAll();
    expect(listed).toHaveLength(1);
    expect(listed[0]?.id).toBe("c1");
    expect(listed[0]?.resolved).toBe(true);
    expect(listed[0]?.replies).toHaveLength(1);
    expect(listed[0]?.replies[0]?.content).toBe("Reply");
  });

  it("can write + migrate a legacy comments array root on a CJS Doc", () => {
    const Ycjs = requireYjsCjs();

    const doc = new Ycjs.Doc();
    doc.getArray("comments"); // force legacy schema

    const mgr = new CommentManager(doc);

    mgr.addComment({
      cellRef: "A1",
      kind: "threaded",
      content: "Hello",
      author: { id: "u1", name: "Alice" },
      id: "c1",
      now: 1,
    });
    mgr.addReply({
      commentId: "c1",
      content: "Reply",
      author: { id: "u1", name: "Alice" },
      id: "r1",
      now: 2,
    });

    expect(getCommentsRoot(doc).kind).toBe("array");
    expect(mgr.listAll()[0]?.replies).toHaveLength(1);

    expect(migrateCommentsArrayToMap(doc)).toBe(true);
    expect(getCommentsRoot(doc).kind).toBe("map");

    // No data loss after migration.
    const listed = mgr.listAll();
    expect(listed).toHaveLength(1);
    expect(listed[0]?.id).toBe("c1");
    expect(listed[0]?.replies).toHaveLength(1);
    expect(listed[0]?.replies[0]?.content).toBe("Reply");
  });

  it("migrates a legacy comments array root containing foreign maps (CJS applyUpdate into ESM Doc)", () => {
    const Ycjs = requireYjsCjs();

    const remote = new Ycjs.Doc();
    const comments = remote.getArray("comments");
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

      const replies = new Ycjs.Array();
      const reply = new Ycjs.Map();
      reply.set("id", "r1");
      reply.set("authorId", "u1");
      reply.set("authorName", "Alice");
      reply.set("createdAt", 1);
      reply.set("updatedAt", 1);
      reply.set("content", "Reply");
      reply.set("mentions", []);
      replies.push([reply]);
      comment.set("replies", replies);

      comments.push([comment]);
    });

    const update = Ycjs.encodeStateAsUpdate(remote);

    const doc = new Y.Doc();
    // Apply update with the CJS build so nested Y.Maps are from the CJS instance.
    Ycjs.applyUpdate(doc, update);

    const mgr = new CommentManager(doc);
    expect(mgr.listAll().map((c) => c.id)).toEqual(["c1"]);
    expect(mgr.listAll()[0]?.replies).toHaveLength(1);

    expect(migrateCommentsArrayToMap(doc)).toBe(true);
    expect(getCommentsRoot(doc).kind).toBe("map");

    // After migration, comments should be stored as *local* Y.Maps/Y.Arrays.
    const root = doc.getMap("comments");
    const c1 = root.get("c1");
    expect(c1).toBeInstanceOf(Y.Map);
    expect(c1?.get("replies")).toBeInstanceOf(Y.Array);
  });
});
