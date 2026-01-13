import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { YjsVersionStore } from "./yjsVersionStore.js";

describe("YjsVersionStore: pruneIncompleteVersions()", () => {
  it("removes stale incomplete streamed records from both versions and versionsMeta.order", async () => {
    const doc = new Y.Doc();
    const store = new YjsVersionStore({
      doc,
      writeMode: "stream",
      chunkSize: 1024,
      maxChunksPerTransaction: 2,
    });

    await store.saveVersion({
      id: "complete",
      kind: "snapshot",
      timestampMs: Date.now(),
      userId: null,
      userName: null,
      description: null,
      checkpointName: null,
      checkpointLocked: null,
      checkpointAnnotations: null,
      snapshot: new Uint8Array([1, 2, 3]),
    });

    // Insert an incomplete record (mimics a client crash mid-stream).
    doc.transact(() => {
      const versions = doc.getMap("versions");
      const meta = doc.getMap("versionsMeta");
      let order = meta.get("order") as any;
      if (!(order instanceof Y.Array)) {
        order = new Y.Array<string>();
        meta.set("order", order);
      }

      const record = new Y.Map<any>();
      record.set("schemaVersion", 1);
      record.set("id", "incomplete");
      record.set("kind", "snapshot");
      record.set("timestampMs", Date.now());
      record.set("userId", null);
      record.set("userName", null);
      record.set("description", null);
      record.set("checkpointName", null);
      record.set("checkpointLocked", null);
      record.set("checkpointAnnotations", null);
      record.set("compression", "none");
      record.set("snapshotEncoding", "chunks");
      record.set("snapshotChunkCountExpected", 2);
      record.set("snapshotComplete", false);
      const chunks = new Y.Array<Uint8Array>();
      chunks.push([new Uint8Array([1, 2, 3])]);
      record.set("snapshotChunks", chunks);
      versions.set("incomplete", record);
      order.push(["incomplete"]);
    }, "test");

    const listed = await store.listVersions();
    expect(listed.map((v) => v.id)).toEqual(["complete"]);

    await store.pruneIncompleteVersions({ olderThanMs: 0 });

    expect(doc.getMap("versions").get("incomplete")).toBeUndefined();
    expect(doc.getMap("versions").get("complete")).toBeDefined();

    const order = doc.getMap("versionsMeta").get("order") as any;
    expect(order?.toArray?.()).toEqual(["complete"]);
  });

  it("does not prune recent incomplete streamed records by default (avoids racing slow streams)", async () => {
    const doc = new Y.Doc();
    const store = new YjsVersionStore({
      doc,
      writeMode: "stream",
      chunkSize: 1024,
      maxChunksPerTransaction: 2,
    });

    doc.transact(() => {
      const versions = doc.getMap("versions");
      const meta = doc.getMap("versionsMeta");
      const order = new Y.Array<string>();
      meta.set("order", order);

      const record = new Y.Map<any>();
      record.set("schemaVersion", 1);
      record.set("id", "incomplete");
      record.set("kind", "snapshot");
      record.set("timestampMs", Date.now());
      record.set("createdAtMs", Date.now());
      record.set("compression", "none");
      record.set("snapshotEncoding", "chunks");
      record.set("snapshotChunkCountExpected", 2);
      record.set("snapshotComplete", false);
      const chunks = new Y.Array<Uint8Array>();
      chunks.push([new Uint8Array([1, 2, 3])]);
      record.set("snapshotChunks", chunks);
      versions.set("incomplete", record);
      order.push(["incomplete"]);
    }, "test");

    // listVersions() should not return the incomplete record, but should also not
    // delete it because it's too new to be considered stale under the default
    // policy.
    expect(await store.listVersions()).toEqual([]);
    expect(doc.getMap("versions").get("incomplete")).toBeDefined();
    expect((doc.getMap("versionsMeta").get("order") as any)?.toArray?.()).toEqual(["incomplete"]);
  });

  it("prefers createdAtMs over timestampMs when checking staleness", async () => {
    const doc = new Y.Doc();
    const store = new YjsVersionStore({
      doc,
      writeMode: "stream",
      chunkSize: 1024,
      maxChunksPerTransaction: 2,
    });

    doc.transact(() => {
      const versions = doc.getMap("versions");
      const meta = doc.getMap("versionsMeta");
      const order = new Y.Array<string>();
      meta.set("order", order);

      const record = new Y.Map<any>();
      record.set("schemaVersion", 1);
      record.set("id", "incomplete");
      record.set("kind", "snapshot");
      // Deliberately "ancient" timestamp (could be user-supplied); createdAtMs
      // reflects when the stream began.
      record.set("timestampMs", 1);
      record.set("createdAtMs", Date.now());
      record.set("compression", "none");
      record.set("snapshotEncoding", "chunks");
      record.set("snapshotChunkCountExpected", 2);
      record.set("snapshotComplete", false);
      const chunks = new Y.Array<Uint8Array>();
      chunks.push([new Uint8Array([1, 2, 3])]);
      record.set("snapshotChunks", chunks);
      versions.set("incomplete", record);
      order.push(["incomplete"]);
    }, "test");

    await store.pruneIncompleteVersions();
    expect(doc.getMap("versions").get("incomplete")).toBeDefined();
    expect((doc.getMap("versionsMeta").get("order") as any)?.toArray?.()).toEqual(["incomplete"]);
  });

  it("finalizes records that have all expected chunks even if snapshotComplete=false", async () => {
    const doc = new Y.Doc();
    const store = new YjsVersionStore({
      doc,
      writeMode: "stream",
      chunkSize: 1024,
      maxChunksPerTransaction: 2,
    });

    doc.transact(() => {
      const versions = doc.getMap("versions");
      const meta = doc.getMap("versionsMeta");
      const order = new Y.Array<string>();
      meta.set("order", order);

      const record = new Y.Map<any>();
      record.set("schemaVersion", 1);
      record.set("id", "recoverable");
      record.set("kind", "snapshot");
      record.set("timestampMs", 1);
      record.set("createdAtMs", 1);
      record.set("compression", "none");
      record.set("snapshotEncoding", "chunks");
      record.set("snapshotChunkCountExpected", 1);
      record.set("snapshotComplete", false);
      const chunks = new Y.Array<Uint8Array>();
      chunks.push([new Uint8Array([4, 5, 6])]);
      record.set("snapshotChunks", chunks);
      versions.set("recoverable", record);
      order.push(["recoverable"]);
    }, "test");

    expect(await store.getVersion("recoverable")).toBeNull();

    await store.pruneIncompleteVersions({ olderThanMs: 0 });

    const recovered = await store.getVersion("recoverable");
    expect(recovered?.id).toBe("recoverable");
    expect(Array.from(recovered!.snapshot)).toEqual([4, 5, 6]);
    expect((doc.getMap("versions").get("recoverable") as any)?.get?.("snapshotComplete")).toBe(true);
    expect((doc.getMap("versionsMeta").get("order") as any)?.toArray?.()).toEqual(["recoverable"]);
  });

  it("does not finalize (and instead prunes) malformed records missing required metadata", async () => {
    const doc = new Y.Doc();
    const store = new YjsVersionStore({
      doc,
      writeMode: "stream",
      chunkSize: 1024,
      maxChunksPerTransaction: 2,
    });

    doc.transact(() => {
      const versions = doc.getMap("versions");
      const meta = doc.getMap("versionsMeta");
      const order = new Y.Array<string>();
      meta.set("order", order);

      const record = new Y.Map<any>();
      record.set("schemaVersion", 1);
      record.set("id", "broken");
      record.set("kind", "snapshot");
      // Invalid type; getVersion would throw if snapshotComplete flipped true.
      record.set("timestampMs", "not-a-number");
      record.set("compression", "none");
      record.set("snapshotEncoding", "chunks");
      record.set("snapshotChunkCountExpected", 1);
      record.set("snapshotComplete", false);
      const chunks = new Y.Array<Uint8Array>();
      chunks.push([new Uint8Array([7, 8, 9])]);
      record.set("snapshotChunks", chunks);
      versions.set("broken", record);
      order.push(["broken"]);
    }, "test");

    await expect(store.getVersion("broken")).resolves.toBeNull();
    await store.pruneIncompleteVersions({ olderThanMs: 0 });
    expect(doc.getMap("versions").get("broken")).toBeUndefined();
    expect((doc.getMap("versionsMeta").get("order") as any)?.toArray?.()).toEqual([]);
  });

  it("treats future timestamps as 'now' so they can still be pruned", async () => {
    const doc = new Y.Doc();
    const store = new YjsVersionStore({
      doc,
      writeMode: "stream",
      chunkSize: 1024,
      maxChunksPerTransaction: 2,
    });

    doc.transact(() => {
      const versions = doc.getMap("versions");
      const meta = doc.getMap("versionsMeta");
      const order = new Y.Array<string>();
      meta.set("order", order);

      const record = new Y.Map<any>();
      record.set("schemaVersion", 1);
      record.set("id", "future");
      record.set("kind", "snapshot");
      record.set("timestampMs", Date.now() + 60_000);
      record.set("compression", "none");
      record.set("snapshotEncoding", "chunks");
      record.set("snapshotChunkCountExpected", 2);
      record.set("snapshotComplete", false);
      const chunks = new Y.Array<Uint8Array>();
      chunks.push([new Uint8Array([1])]);
      record.set("snapshotChunks", chunks);
      versions.set("future", record);
      order.push(["future"]);
    }, "test");

    await store.pruneIncompleteVersions({ olderThanMs: 0 });
    expect(doc.getMap("versions").get("future")).toBeUndefined();
    expect((doc.getMap("versionsMeta").get("order") as any)?.toArray?.()).toEqual([]);
  });

  it("prunes stale incomplete chunk records even when snapshotComplete is unset", async () => {
    const doc = new Y.Doc();
    const store = new YjsVersionStore({
      doc,
      writeMode: "stream",
      chunkSize: 1024,
      maxChunksPerTransaction: 2,
    });

    doc.transact(() => {
      const versions = doc.getMap("versions");
      const meta = doc.getMap("versionsMeta");
      const order = new Y.Array<string>();
      meta.set("order", order);

      const record = new Y.Map<any>();
      record.set("schemaVersion", 1);
      record.set("id", "missingChunks");
      record.set("kind", "snapshot");
      record.set("timestampMs", 1);
      record.set("compression", "none");
      record.set("snapshotEncoding", "chunks");
      record.set("snapshotChunkCountExpected", 2);
      // NOTE: intentionally do not set snapshotComplete.
      const chunks = new Y.Array<Uint8Array>();
      chunks.push([new Uint8Array([1, 2, 3])]);
      record.set("snapshotChunks", chunks);
      versions.set("missingChunks", record);
      order.push(["missingChunks"]);
    }, "test");

    await expect(store.getVersion("missingChunks")).resolves.toBeNull();
    await store.pruneIncompleteVersions({ olderThanMs: 0 });
    expect(doc.getMap("versions").get("missingChunks")).toBeUndefined();
    expect((doc.getMap("versionsMeta").get("order") as any)?.toArray?.()).toEqual([]);
  });

  it("listVersions() opportunistically prunes stale incomplete records using the default policy", async () => {
    const doc = new Y.Doc();
    const store = new YjsVersionStore({
      doc,
      writeMode: "stream",
      chunkSize: 1024,
      maxChunksPerTransaction: 2,
    });

    doc.transact(() => {
      const versions = doc.getMap("versions");
      const meta = doc.getMap("versionsMeta");
      const order = new Y.Array<string>();
      meta.set("order", order);

      const record = new Y.Map<any>();
      record.set("schemaVersion", 1);
      record.set("id", "staleIncomplete");
      record.set("kind", "snapshot");
      // Extremely old timestamps; should be pruned by default (10 minute) threshold.
      record.set("timestampMs", 1);
      record.set("createdAtMs", 1);
      record.set("compression", "none");
      record.set("snapshotEncoding", "chunks");
      record.set("snapshotChunkCountExpected", 2);
      record.set("snapshotComplete", false);
      const chunks = new Y.Array<Uint8Array>();
      chunks.push([new Uint8Array([1])]);
      record.set("snapshotChunks", chunks);

      versions.set("staleIncomplete", record);
      order.push(["staleIncomplete"]);
    }, "test");

    expect(await store.listVersions()).toEqual([]);
    expect(doc.getMap("versions").get("staleIncomplete")).toBeUndefined();
    expect((doc.getMap("versionsMeta").get("order") as any)?.toArray?.()).toEqual([]);
  });

  it("prunes stale records missing snapshotChunks when snapshotEncoding is chunks", async () => {
    const doc = new Y.Doc();
    const store = new YjsVersionStore({
      doc,
      writeMode: "stream",
      chunkSize: 1024,
      maxChunksPerTransaction: 2,
    });

    doc.transact(() => {
      const versions = doc.getMap("versions");
      const meta = doc.getMap("versionsMeta");
      const order = new Y.Array<string>();
      meta.set("order", order);

      const record = new Y.Map<any>();
      record.set("schemaVersion", 1);
      record.set("id", "missingChunksArr");
      record.set("kind", "snapshot");
      record.set("timestampMs", 1);
      record.set("compression", "none");
      record.set("snapshotEncoding", "chunks");
      record.set("snapshotChunkCountExpected", 1);
      // NOTE: intentionally omit snapshotChunks and snapshotComplete.
      versions.set("missingChunksArr", record);
      order.push(["missingChunksArr"]);
    }, "test");

    await expect(store.getVersion("missingChunksArr")).resolves.toBeNull();
    await store.pruneIncompleteVersions({ olderThanMs: 0 });
    expect(doc.getMap("versions").get("missingChunksArr")).toBeUndefined();
    expect((doc.getMap("versionsMeta").get("order") as any)?.toArray?.()).toEqual([]);
  });

  it("uses incompleteSinceMs to avoid clock skew preventing pruning", async () => {
    const doc = new Y.Doc();
    const store = new YjsVersionStore({
      doc,
      writeMode: "stream",
      chunkSize: 1024,
      maxChunksPerTransaction: 2,
    });

    const future = Date.now() + 365 * 24 * 60 * 60 * 1000;
    doc.transact(() => {
      const versions = doc.getMap("versions");
      const meta = doc.getMap("versionsMeta");
      const order = new Y.Array<string>();
      meta.set("order", order);

      const record = new Y.Map<any>();
      record.set("schemaVersion", 1);
      record.set("id", "skewed");
      record.set("kind", "snapshot");
      // Writer clock is far in the future.
      record.set("timestampMs", future);
      record.set("createdAtMs", future);
      record.set("compression", "none");
      record.set("snapshotEncoding", "chunks");
      record.set("snapshotChunkCountExpected", 2);
      record.set("snapshotComplete", false);
      const chunks = new Y.Array<Uint8Array>();
      chunks.push([new Uint8Array([1])]);
      record.set("snapshotChunks", chunks);
      versions.set("skewed", record);
      order.push(["skewed"]);
    }, "test");

    // First pass should mark incompleteSinceMs (based on local time), not prune.
    await store.pruneIncompleteVersions({ olderThanMs: 10_000 });
    const raw = doc.getMap("versions").get("skewed") as any;
    expect(raw).toBeDefined();
    expect(typeof raw?.get?.("incompleteSinceMs")).toBe("number");
    expect(raw.get("incompleteSinceMs")).toBeLessThanOrEqual(Date.now());

    // If we later decide it's old based on incompleteSinceMs, it should prune even though createdAtMs/timestampMs are future.
    doc.transact(() => {
      const r = doc.getMap("versions").get("skewed") as any;
      r?.set?.("incompleteSinceMs", 1);
    }, "test");
    await store.pruneIncompleteVersions({ olderThanMs: 1 });
    expect(doc.getMap("versions").get("skewed")).toBeUndefined();
    expect((doc.getMap("versionsMeta").get("order") as any)?.toArray?.()).toEqual([]);
  });
});
