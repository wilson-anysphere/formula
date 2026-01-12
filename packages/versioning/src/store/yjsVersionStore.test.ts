import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { YjsVersionStore } from "./yjsVersionStore.js";

describe("YjsVersionStore (streaming mode)", () => {
  it("streams large snapshots across multiple Yjs transactions and round-trips bytes", async () => {
    const doc = new Y.Doc();
    let updateCount = 0;
    doc.on("update", () => {
      updateCount += 1;
    });

    const store = new YjsVersionStore({
      doc,
      writeMode: "stream",
      chunkSize: 1024,
      maxChunksPerTransaction: 2,
    });

    const snapshot = new Uint8Array(10_000);
    for (let i = 0; i < snapshot.length; i += 1) snapshot[i] = i % 256;

    await store.saveVersion({
      id: "v1",
      kind: "snapshot",
      timestampMs: 1,
      userId: null,
      userName: null,
      description: null,
      checkpointName: null,
      checkpointLocked: null,
      checkpointAnnotations: null,
      snapshot,
    });

    const roundTrip = await store.getVersion("v1");
    expect(roundTrip).not.toBeNull();
    expect(Buffer.from(roundTrip!.snapshot)).toEqual(Buffer.from(snapshot));

    const raw = doc.getMap("versions").get("v1") as any;
    expect(raw?.get?.("snapshotEncoding")).toBe("chunks");
    expect(raw?.get?.("snapshotComplete")).toBe(true);
    expect(updateCount).toBeGreaterThan(1);
  });

  it("filters incomplete streamed versions from listVersions()", async () => {
    const doc = new Y.Doc();
    const store = new YjsVersionStore({
      doc,
      writeMode: "stream",
      chunkSize: 1024,
      maxChunksPerTransaction: 2,
    });

    // Create a partially-written version record (snapshotComplete=false and only
    // a subset of chunks appended).
    doc.transact(() => {
      const versions = doc.getMap("versions");
      const meta = doc.getMap("versionsMeta");
      const order = new Y.Array<string>();
      meta.set("order", order);
      const record = new Y.Map<any>();
      record.set("schemaVersion", 1);
      record.set("id", "incomplete");
      record.set("kind", "snapshot");
      record.set("timestampMs", 1);
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

    expect(await store.getVersion("incomplete")).toBeNull();
    expect(await store.listVersions()).toEqual([]);
  });
});

