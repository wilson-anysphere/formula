import * as Y from "yjs";
import { describe, expect, it } from "vitest";

import { bindDocToStorage } from "../src/persistence";
import { CommentManager } from "../src/manager";

class MemoryStorage {
  private map = new Map<string, string>();

  getItem(key: string): string | null {
    return this.map.get(key) ?? null;
  }

  setItem(key: string, value: string): void {
    this.map.set(key, value);
  }

  removeItem(key: string): void {
    this.map.delete(key);
  }
}

describe("collab comments persistence", () => {
  it("persists ydoc state to storage and restores it", () => {
    const storage = new MemoryStorage();

    const doc1 = new Y.Doc();
    const stop = bindDocToStorage(doc1, storage, "doc");

    const mgr1 = new CommentManager(doc1);
    mgr1.addComment({
      cellRef: "A1",
      kind: "threaded",
      content: "Persist me",
      author: { id: "u1", name: "Alice" },
      id: "c1",
      now: 1,
    });

    stop();

    const doc2 = new Y.Doc();
    bindDocToStorage(doc2, storage, "doc")();
    const mgr2 = new CommentManager(doc2);

    expect(mgr2.listAll()).toEqual(mgr1.listAll());
  });
});

