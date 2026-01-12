import assert from "node:assert/strict";
import test from "node:test";

import { Y } from "./yjs-interop.ts";
import { inspectReservedRootUpdate } from "../src/reservedRootInspector.js";

function keyPathIncludes(
  path: Array<string | number>,
  required: string[]
): boolean {
  let i = 0;
  for (const part of path) {
    if (part === required[i]) i += 1;
    if (i >= required.length) return true;
  }
  return required.length === 0;
}

test("reservedRootInspector: versionsMeta.order array mutation is attributed to versionsMeta.order", () => {
  const serverDoc = new Y.Doc();
  serverDoc.transact(() => {
    const meta = serverDoc.getMap("versionsMeta");
    const order = new Y.Array<string>();
    meta.set("order", order);
  });

  const clientDoc = new Y.Doc();
  Y.applyUpdate(clientDoc, Y.encodeStateAsUpdate(serverDoc));

  const order = clientDoc.getMap("versionsMeta").get("order") as any;
  assert.ok(order, "expected versionsMeta.order to exist on client");

  clientDoc.transact(() => {
    order.push(["v1"]);
  });

  const update = Y.encodeStateAsUpdate(clientDoc, Y.encodeStateVector(serverDoc));
  const hits = inspectReservedRootUpdate({
    baseDoc: serverDoc,
    update,
    reservedRoots: ["versionsMeta"],
  });

  assert.ok(
    hits.some((hit) => hit.root === "versionsMeta" && keyPathIncludes(hit.keyPath, ["order"])),
    `expected an inspector hit for versionsMeta.order; got: ${JSON.stringify(hits)}`
  );
});

test("reservedRootInspector: versions[v1].snapshotChunks array mutation is attributed to versions.v1.snapshotChunks", () => {
  const serverDoc = new Y.Doc();
  serverDoc.transact(() => {
    const versions = serverDoc.getMap("versions");
    const record = new Y.Map<any>();
    const snapshotChunks = new Y.Array<Uint8Array>();
    record.set("snapshotChunks", snapshotChunks);
    versions.set("v1", record);
  });

  const clientDoc = new Y.Doc();
  Y.applyUpdate(clientDoc, Y.encodeStateAsUpdate(serverDoc));

  const record = clientDoc.getMap("versions").get("v1") as any;
  assert.ok(record && typeof record.get === "function", "expected versions.v1 record to exist on client");
  const snapshotChunks = record.get("snapshotChunks") as any;
  assert.ok(snapshotChunks, "expected versions.v1.snapshotChunks to exist on client");

  clientDoc.transact(() => {
    snapshotChunks.push([new Uint8Array([1, 2, 3])]);
  });

  const update = Y.encodeStateAsUpdate(clientDoc, Y.encodeStateVector(serverDoc));
  const hits = inspectReservedRootUpdate({
    baseDoc: serverDoc,
    update,
    reservedRoots: ["versions"],
  });

  assert.ok(
    hits.some(
      (hit) => hit.root === "versions" && keyPathIncludes(hit.keyPath, ["v1", "snapshotChunks"])
    ),
    `expected an inspector hit for versions.v1.snapshotChunks; got: ${JSON.stringify(hits)}`
  );
});

