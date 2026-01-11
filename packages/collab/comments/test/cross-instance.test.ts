import { createRequire } from "node:module";

import * as Y from "yjs";
import { describe, expect, it } from "vitest";

import { CommentManager } from "../src/manager";

describe("collab comments cross-instance Yjs roots (ESM doc + CJS applyUpdate)", () => {
  it("reads and mutates a comments map root created by a different Yjs instance", () => {
    const require = createRequire(import.meta.url);
    // eslint-disable-next-line import/no-named-as-default-member
    const Ycjs = require("yjs");

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
    const require = createRequire(import.meta.url);
    // eslint-disable-next-line import/no-named-as-default-member
    const Ycjs = require("yjs");

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
});

