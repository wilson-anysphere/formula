import assert from "node:assert/strict";
import test from "node:test";

import { inspectUpdate } from "../src/yjsUpdateInspection.js";
import { Y } from "./yjs-interop.ts";

const reservedRootNames = new Set<string>(["versions", "versionsMeta"]);
const reservedRootPrefixes = ["branching:"];

test("yjs update inspection: direct root write flags reserved root", () => {
  const serverDoc = new Y.Doc();
  const clientDoc = new Y.Doc();

  clientDoc.getMap("versions").set("v1", new Y.Map());
  const update = Y.encodeStateAsUpdate(clientDoc);

  const res = inspectUpdate({
    ydoc: serverDoc,
    update,
    reservedRootNames,
    reservedRootPrefixes,
    maxTouches: 10,
  });

  assert.equal(res.touchesReserved, true);
  assert.ok(
    res.touches.some(
      (t) => t.kind === "insert" && t.root === "versions" && t.keyPath.includes("v1")
    )
  );

  serverDoc.destroy();
  clientDoc.destroy();
});

test("yjs update inspection: nested write resolves root + keyPath", () => {
  const serverDoc = new Y.Doc();
  const versions = serverDoc.getMap("versions");
  const v1 = new Y.Map();
  versions.set("v1", v1);

  const clientDoc = new Y.Doc();
  Y.applyUpdate(clientDoc, Y.encodeStateAsUpdate(serverDoc));

  const v1Client = clientDoc.getMap("versions").get("v1") as Y.Map<unknown>;
  v1Client.set("checkpointLocked", true);

  const update = Y.encodeStateAsUpdate(clientDoc, Y.encodeStateVector(serverDoc));

  const res = inspectUpdate({
    ydoc: serverDoc,
    update,
    reservedRootNames,
    reservedRootPrefixes,
    maxTouches: 10,
  });

  assert.equal(res.touchesReserved, true);
  assert.ok(
    res.touches.some(
      (t) =>
        t.kind === "insert" &&
        t.root === "versions" &&
        t.keyPath.length >= 2 &&
        t.keyPath[0] === "v1" &&
        t.keyPath[1] === "checkpointLocked"
    )
  );

  serverDoc.destroy();
  clientDoc.destroy();
});

test("yjs update inspection: parent-info copy case is resolved", () => {
  const serverDoc = new Y.Doc();
  const clientDoc = new Y.Doc();
  const versions = clientDoc.getMap("versions");

  clientDoc.transact(() => {
    versions.set("v1", 1);
    versions.set("v1", 2);
  });

  const update = Y.encodeStateAsUpdate(clientDoc);

  // Ensure the test actually exercises the "parent info omitted" encoding.
  const decoded = Y.decodeUpdate(update);
  const sawParentOmitted = decoded.structs.some(
    (s) => s.constructor.name === "Item" && (s as any).parent == null
  );
  assert.equal(sawParentOmitted, true);

  const res = inspectUpdate({
    ydoc: serverDoc,
    update,
    reservedRootNames,
    reservedRootPrefixes,
    maxTouches: 10,
  });

  assert.equal(res.touchesReserved, true);
  assert.ok(res.touches.some((t) => t.kind === "insert" && t.root === "versions" && t.keyPath[0] === "v1"));

  serverDoc.destroy();
  clientDoc.destroy();
});

test("yjs update inspection: delete-only update (delete set) is inspected", () => {
  const serverDoc = new Y.Doc();
  serverDoc.getMap("versionsMeta").set("order", "abc");

  const clientDoc = new Y.Doc();
  Y.applyUpdate(clientDoc, Y.encodeStateAsUpdate(serverDoc));

  clientDoc.getMap("versionsMeta").delete("order");

  const update = Y.encodeStateAsUpdate(clientDoc, Y.encodeStateVector(serverDoc));
  const decoded = Y.decodeUpdate(update);
  assert.equal(decoded.structs.length, 0);
  assert.ok(decoded.ds.clients.size > 0);

  const res = inspectUpdate({
    ydoc: serverDoc,
    update,
    reservedRootNames,
    reservedRootPrefixes,
    maxTouches: 10,
  });

  assert.equal(res.touchesReserved, true);
  assert.ok(
    res.touches.some(
      (t) => t.kind === "delete" && t.root === "versionsMeta" && t.keyPath[0] === "order"
    )
  );

  serverDoc.destroy();
  clientDoc.destroy();
});

test("yjs update inspection: reserved root prefix is matched", () => {
  const serverDoc = new Y.Doc();
  const clientDoc = new Y.Doc();

  clientDoc.getMap("branching:main").set("x", 1);
  const update = Y.encodeStateAsUpdate(clientDoc);

  const res = inspectUpdate({
    ydoc: serverDoc,
    update,
    reservedRootNames,
    reservedRootPrefixes,
    maxTouches: 10,
  });

  assert.equal(res.touchesReserved, true);
  assert.ok(res.touches.some((t) => t.kind === "insert" && t.root === "branching:main"));

  serverDoc.destroy();
  clientDoc.destroy();
});

