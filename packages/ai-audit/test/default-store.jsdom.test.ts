// @vitest-environment jsdom
import { describe, expect, it } from "vitest";

import { BoundedAIAuditStore } from "../src/bounded-store.js";
import { createDefaultAIAuditStore } from "../src/index.js";
import { LocalStorageAIAuditStore } from "../src/local-storage-store.js";

function unwrap(store: unknown): unknown {
  if (!(store instanceof BoundedAIAuditStore)) return store;
  // `store` is a private field but is present at runtime.
  return (store as any).store;
}

describe("createDefaultAIAuditStore (jsdom)", () => {
  it("chooses LocalStorageAIAuditStore when localStorage is available", async () => {
    const store = await createDefaultAIAuditStore();
    expect(store).toBeInstanceOf(BoundedAIAuditStore);
    expect(unwrap(store)).toBeInstanceOf(LocalStorageAIAuditStore);
  });
});
