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
});
