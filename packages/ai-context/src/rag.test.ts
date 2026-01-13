import { describe, expect, it } from "vitest";

import { InMemoryVectorStore } from "./rag.js";

describe("InMemoryVectorStore.search", () => {
  it("returns deterministic id ordering when similarity scores tie", async () => {
    const store = new InMemoryVectorStore();

    const embedding = [1, 0, 0];

    // Insert out of order to ensure the sort tie-breaker (id) is what drives the
    // final result ordering when similarity scores are identical.
    await store.add([
      { id: "chunk-b", embedding, metadata: null, text: "b" },
      { id: "chunk-a", embedding, metadata: null, text: "a" },
      { id: "chunk-c", embedding, metadata: null, text: "c" },
    ]);

    const results = await store.search([1, 0, 0], 10);

    expect(results.map((r) => r.item.id)).toEqual(["chunk-a", "chunk-b", "chunk-c"]);
  });
});

