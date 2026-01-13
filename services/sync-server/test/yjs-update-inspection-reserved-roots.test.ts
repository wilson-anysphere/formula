import assert from "node:assert/strict";
import test from "node:test";

import { inspectUpdate, inspectUpdateForReservedRootGuard } from "../src/yjsUpdateInspection.js";
import { Y } from "./yjs-interop.ts";

const reservedRootNames = new Set<string>(["versions", "versionsMeta"]);
const reservedRootPrefixes = ["branching:"];

function isYjsItemStruct(value: unknown): value is { parent: unknown; parentSub: unknown } {
  if (!value || typeof value !== "object") return false;
  const maybe = value as any;
  // Yjs internal `Item` structs have these fields (see yjs/src/structs/Item).
  if (!("id" in maybe)) return false;
  if (typeof maybe.length !== "number") return false;
  if (!("content" in maybe)) return false;
  if (!("parent" in maybe)) return false;
  if (!("parentSub" in maybe)) return false;
  if (typeof maybe.content?.getContent !== "function") return false;
  return true;
}

function setDeterministicClientId(doc: Y.Doc, clientId: number): void {
  // Yjs generates random clientIDs; for regression tests we'd rather have stable,
  // deterministic struct IDs across runs.
  (doc as any).clientID = clientId;
}

function createBaselineReservedRootDoc(): Y.Doc {
  const doc = new Y.Doc();
  setDeterministicClientId(doc, 1);

  doc.transact(() => {
    const versions = doc.getMap("versions");

    const makeRecord = (id: string) => {
      const record = new Y.Map<any>();
      record.set("schemaVersion", 1);
      record.set("id", id);
      record.set("checkpointLocked", false);
      record.set("snapshotChunks", new Y.Array<Uint8Array>());
      return record;
    };

    versions.set("v1", makeRecord("v1"));
    versions.set("v2", makeRecord("v2"));

    const meta = doc.getMap("versionsMeta");
    const order = new Y.Array<string>();
    order.push(["v1", "v2"]);
    meta.set("order", order);
  }, "baseline-reserved-roots");

  return doc;
}

function collectKeyPathsFromObserveDeep(params: {
  baselineUpdate: Uint8Array;
  update: Uint8Array;
}): Map<string, Set<string>> {
  const shadow = new Y.Doc();
  Y.applyUpdate(shadow, params.baselineUpdate);

  const touchedByRoot = new Map<string, Set<string>>([
    ["versions", new Set<string>()],
    ["versionsMeta", new Set<string>()],
  ]);

  const makeObserver =
    (root: string) =>
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (events: any[]) => {
      const out = touchedByRoot.get(root);
      if (!out) return;
      for (const event of events) {
        const rawPath = event?.path;
        const path =
          Array.isArray(rawPath) && rawPath.length > 0
            ? rawPath.filter((p: unknown): p is string => typeof p === "string")
            : [];

        const keys = event?.changes?.keys;
        if (keys) {
          let addedKey = false;
          if (typeof keys.entries === "function") {
            for (const [key] of keys.entries()) {
              if (typeof key === "string") {
                out.add(JSON.stringify([...path, key]));
                addedKey = true;
              }
            }
          } else if (typeof keys.keys === "function") {
            for (const key of keys.keys()) {
              if (typeof key === "string") {
                out.add(JSON.stringify([...path, key]));
                addedKey = true;
              }
            }
          }
          if (addedKey) continue;
        }

        // Array/text updates don't have `changes.keys`, but still represent a touch
        // of the map key that holds the nested type.
        if (path.length > 0) {
          out.add(JSON.stringify(path));
        }
      }
    };

  const versions = shadow.getMap("versions");
  const meta = shadow.getMap("versionsMeta");
  const versionsObserver = makeObserver("versions");
  const metaObserver = makeObserver("versionsMeta");

  versions.observeDeep(versionsObserver);
  meta.observeDeep(metaObserver);

  try {
    Y.applyUpdate(shadow, params.update);
  } finally {
    versions.unobserveDeep(versionsObserver);
    meta.unobserveDeep(metaObserver);
    shadow.destroy();
  }

  return touchedByRoot;
}

function collectKeyPathsFromInspector(params: {
  serverDoc: Y.Doc;
  update: Uint8Array;
}): Map<string, Set<string>> {
  const res = inspectUpdate({
    ydoc: params.serverDoc,
    update: params.update,
    reservedRootNames,
    reservedRootPrefixes,
    maxTouches: 1000,
  });

  assert.equal(
    res.unknownReason,
    undefined,
    `inspector unexpectedly failed closed: ${res.unknownReason}`
  );

  const out = new Map<string, Set<string>>();
  for (const touch of res.touches) {
    const existing = out.get(touch.root);
    const set = existing ?? new Set<string>();
    set.add(JSON.stringify(touch.keyPath));
    if (!existing) out.set(touch.root, set);
  }
  return out;
}

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

test("yjs update inspection: matches observeDeep ground truth for reserved roots", () => {
  const serverDoc = createBaselineReservedRootDoc();
  const baselineUpdate = Y.encodeStateAsUpdate(serverDoc);
  const serverStateVector = Y.encodeStateVector(serverDoc);

  const makeUpdate = (
    clientId: number,
    mutate: (clientDoc: Y.Doc) => void
  ): Uint8Array => {
    const clientDoc = new Y.Doc();
    setDeterministicClientId(clientDoc, clientId);
    Y.applyUpdate(clientDoc, baselineUpdate);
    mutate(clientDoc);
    const update = Y.encodeStateAsUpdate(clientDoc, serverStateVector);
    clientDoc.destroy();
    return update;
  };

  const scenarios: Array<{ name: string; update: Uint8Array }> = [
    {
      name: "overwrite same map key twice (origin-copy)",
      update: makeUpdate(2, (doc) => {
        const versions = doc.getMap("versions");
        doc.transact(() => {
          versions.set("v3", 1);
          versions.set("v3", 2);
        }, "origin-copy");

         // Ensure this scenario actually exercises the tricky "parent omitted, copy from origin" encoding.
         const update = Y.encodeStateAsUpdate(doc, serverStateVector);
         const decoded = Y.decodeUpdate(update);
         const sawParentOmitted = decoded.structs.some(
          (s) => isYjsItemStruct(s) && (s as any).parent == null
         );
         assert.equal(sawParentOmitted, true);
       }),
     },
    {
      name: "mutate nested record field",
      update: makeUpdate(3, (doc) => {
        const v1 = doc.getMap("versions").get("v1") as any;
        assert.ok(v1 && typeof v1.set === "function");
        doc.transact(() => {
          v1.set("checkpointLocked", true);
        }, "nested-record-field");
      }),
    },
    {
      name: "mutate nested array (order.push)",
      update: makeUpdate(4, (doc) => {
        const order = doc.getMap("versionsMeta").get("order") as any;
        assert.ok(order && typeof order.push === "function");
        doc.transact(() => {
          order.push(["v3"]);
        }, "order-push");

         // Ensure this scenario actually encodes leaf items with parentSub=null (array insertions).
         const update = Y.encodeStateAsUpdate(doc, serverStateVector);
         const decoded = Y.decodeUpdate(update);
         const sawArrayLeaf = decoded.structs.some(
          (s) => isYjsItemStruct(s) && (s as any).parentSub === null
         );
         assert.equal(sawArrayLeaf, true);
       }),
     },
    {
      name: "delete a key (delete set)",
      update: makeUpdate(5, (doc) => {
        doc.transact(() => {
          doc.getMap("versionsMeta").delete("order");
        }, "delete-order-key");
      }),
    },
    {
      name: "merged update touches both versions + versionsMeta",
      update: (() => {
        const u1 = makeUpdate(6, (doc) => {
          const v2 = doc.getMap("versions").get("v2") as any;
          assert.ok(v2 && typeof v2.set === "function");
          doc.transact(() => {
            v2.set("checkpointLocked", true);
          }, "u1-versions");
        });

        const u2 = makeUpdate(7, (doc) => {
          const order = doc.getMap("versionsMeta").get("order") as any;
          assert.ok(order && typeof order.push === "function");
          doc.transact(() => {
            order.push(["v3"]);
          }, "u2-versionsMeta");
        });

        return Y.mergeUpdates([u1, u2]);
      })(),
    },
  ];

  for (const scenario of scenarios) {
    const groundTruth = collectKeyPathsFromObserveDeep({
      baselineUpdate,
      update: scenario.update,
    });
    const inspected = collectKeyPathsFromInspector({
      serverDoc,
      update: scenario.update,
    });

    const rootsFromGroundTruth = Array.from(groundTruth.entries())
      .filter(([_root, paths]) => paths.size > 0)
      .map(([root]) => root)
      .sort();
    const rootsFromInspector = Array.from(inspected.keys()).sort();

    assert.deepEqual(
      rootsFromInspector,
      rootsFromGroundTruth,
      `${scenario.name}: reserved roots touched mismatch`
    );

    for (const root of rootsFromGroundTruth) {
      const expected = Array.from(groundTruth.get(root) ?? []).sort();
      const actual = Array.from(inspected.get(root) ?? []).sort();
      assert.deepEqual(
        actual,
        expected,
        `${scenario.name}: keyPaths mismatch for ${root}`
      );
    }
  }

  serverDoc.destroy();
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

test("yjs update inspection: versionsMeta.order array mutation resolves to versionsMeta.order", () => {
  const serverDoc = new Y.Doc();
  serverDoc.transact(() => {
    const meta = serverDoc.getMap("versionsMeta");
    const order = new Y.Array<string>();
    meta.set("order", order);
  });

  const clientDoc = new Y.Doc();
  Y.applyUpdate(clientDoc, Y.encodeStateAsUpdate(serverDoc));

  const orderClient = clientDoc.getMap("versionsMeta").get("order") as any;
  assert.ok(orderClient, "expected versionsMeta.order to exist on client");

  clientDoc.transact(() => {
    orderClient.push(["v1"]);
  });

  const update = Y.encodeStateAsUpdate(clientDoc, Y.encodeStateVector(serverDoc));

  // Ensure the mutation itself is encoded as sequence items with parentSub=null
  // (Y.Array inserts). The inspector must walk up via the array type's insertion item
  // to recover the `order` map key.
  const decoded = Y.decodeUpdate(update);
  const sawArrayLeaf = decoded.structs.some(
    (s) => isYjsItemStruct(s) && (s as any).parentSub === null
  );
  assert.equal(sawArrayLeaf, true);

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
        t.root === "versionsMeta" &&
        t.keyPath.length >= 1 &&
        t.keyPath[0] === "order"
    ),
    `expected touch for versionsMeta.order, got: ${JSON.stringify(res.touches)}`
  );

  serverDoc.destroy();
  clientDoc.destroy();
});

test("yjs update inspection: versions[v1].snapshotChunks array mutation resolves to versions.v1.snapshotChunks", () => {
  const serverDoc = new Y.Doc();
  serverDoc.transact(() => {
    const versions = serverDoc.getMap("versions");
    const record = new Y.Map<any>();
    record.set("snapshotChunks", new Y.Array<Uint8Array>());
    versions.set("v1", record);
  });

  const clientDoc = new Y.Doc();
  Y.applyUpdate(clientDoc, Y.encodeStateAsUpdate(serverDoc));

  const recordClient = clientDoc.getMap("versions").get("v1") as any;
  assert.ok(recordClient && typeof recordClient.get === "function", "expected versions.v1 record");
  const snapshotChunks = recordClient.get("snapshotChunks") as any;
  assert.ok(snapshotChunks, "expected versions.v1.snapshotChunks array");

  clientDoc.transact(() => {
    snapshotChunks.push([new Uint8Array([1, 2, 3])]);
  });

  const update = Y.encodeStateAsUpdate(clientDoc, Y.encodeStateVector(serverDoc));

  // Ensure the update contains leaf array items (parentSub=null). The inspector must
  // attribute them to `versions.v1.snapshotChunks` via the nested type insertion chain.
  const decoded = Y.decodeUpdate(update);
  const sawArrayLeaf = decoded.structs.some(
    (s) => isYjsItemStruct(s) && (s as any).parentSub === null
  );
  assert.equal(sawArrayLeaf, true);

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
        t.keyPath[1] === "snapshotChunks"
    ),
    `expected touch for versions.v1.snapshotChunks, got: ${JSON.stringify(res.touches)}`
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
    (s) => isYjsItemStruct(s) && (s as any).parent == null
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

test("yjs update inspection: parent-info copy from store (origin points into server store) is resolved", () => {
  const serverDoc = new Y.Doc();
  serverDoc.getMap("versions").set("v1", 1);

  const clientDoc = new Y.Doc();
  Y.applyUpdate(clientDoc, Y.encodeStateAsUpdate(serverDoc));

  clientDoc.getMap("versions").set("v1", 2);

  const update = Y.encodeStateAsUpdate(clientDoc, Y.encodeStateVector(serverDoc));

  const decoded = Y.decodeUpdate(update);
  assert.equal(decoded.structs.length, 1);
  assert.equal((decoded.structs[0] as any).parent, null);
  assert.equal(Buffer.from(update).includes(Buffer.from("versions")), false);

  const res = inspectUpdate({
    ydoc: serverDoc,
    update,
    reservedRootNames,
    reservedRootPrefixes,
    maxTouches: 10,
  });

  assert.equal(res.touchesReserved, true);
  assert.ok(
    res.touches.some((t) => t.root === "versions" && t.keyPath.length >= 1 && t.keyPath[0] === "v1")
  );

  serverDoc.destroy();
  clientDoc.destroy();
});

test("yjs update inspection: origin id inside a struct range is resolved (binary search)", () => {
  const serverDoc = new Y.Doc();
  const txt = new Y.Text();
  txt.insert(0, "hello"); // length 5 (one Item struct range)
  serverDoc.getMap("versions").set("v1", txt);

  const clientDoc = new Y.Doc();
  Y.applyUpdate(clientDoc, Y.encodeStateAsUpdate(serverDoc));

  const txtClient = clientDoc.getMap("versions").get("v1") as Y.Text;
  txtClient.insert(txtClient.length, "!");

  const update = Y.encodeStateAsUpdate(clientDoc, Y.encodeStateVector(serverDoc));

  // The origin id typically points to the last character ID (inside the existing struct range),
  // and the update does not include the root string.
  assert.equal(Buffer.from(update).includes(Buffer.from("versions")), false);

  const res = inspectUpdate({
    ydoc: serverDoc,
    update,
    reservedRootNames,
    reservedRootPrefixes,
    maxTouches: 10,
  });

  assert.equal(res.touchesReserved, true);
  assert.ok(
    res.touches.some((t) => t.root === "versions" && t.keyPath.length >= 1 && t.keyPath[0] === "v1")
  );

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

test("yjs update inspection: prefix match works even when root name is only derivable from store", () => {
  const serverDoc = new Y.Doc();
  serverDoc.getMap("branching:main").set("x", 1);

  const clientDoc = new Y.Doc();
  Y.applyUpdate(clientDoc, Y.encodeStateAsUpdate(serverDoc));

  clientDoc.getMap("branching:main").set("x", 2);
  const update = Y.encodeStateAsUpdate(clientDoc, Y.encodeStateVector(serverDoc));
  assert.equal(Buffer.from(update).includes(Buffer.from("branching:main")), false);

  const res = inspectUpdate({
    ydoc: serverDoc,
    update,
    reservedRootNames,
    reservedRootPrefixes,
    maxTouches: 10,
  });

  assert.equal(res.touchesReserved, true);
  assert.ok(res.touches.some((t) => t.root === "branching:main"));

  serverDoc.destroy();
  clientDoc.destroy();
});

test("yjs update inspection: decodes v2 updates via fallback", () => {
  const serverDoc = new Y.Doc();
  const clientDoc = new Y.Doc();

  clientDoc.getMap("versions").set("v1", "one");
  const updateV2 = Y.encodeStateAsUpdateV2(clientDoc);

  // `decodeUpdate` (v1) can decode v2 updates as a no-op; ensure the inspector still detects touches.
  const res = inspectUpdate({
    ydoc: serverDoc,
    update: updateV2,
    reservedRootNames,
    reservedRootPrefixes,
    maxTouches: 10,
  });

  assert.equal(res.touchesReserved, true);
  assert.ok(res.touches.some((t) => t.root === "versions" && t.keyPath[0] === "v1"));

  serverDoc.destroy();
  clientDoc.destroy();
});

test("yjs update inspection (optimized): matches inspectUpdate root + keyPath on rejection", () => {
  const serverDoc = new Y.Doc();
  serverDoc.getMap("versions").set("v1", new Y.Map());

  const attackerDoc = new Y.Doc();
  Y.applyUpdate(attackerDoc, Y.encodeStateAsUpdate(serverDoc));
  const recordClient = attackerDoc.getMap("versions").get("v1") as any;
  assert.ok(recordClient && typeof recordClient.set === "function");
  recordClient.set("checkpointLocked", true);

  const update = Y.encodeStateAsUpdate(attackerDoc, Y.encodeStateVector(serverDoc));

  const full = inspectUpdate({
    ydoc: serverDoc,
    update,
    reservedRootNames,
    reservedRootPrefixes,
    maxTouches: 1,
  });
  const optimized = inspectUpdateForReservedRootGuard({
    ydoc: serverDoc,
    update,
    reservedRootNames,
    reservedRootPrefixes,
    maxTouches: 1,
  });

  assert.equal(full.touchesReserved, true);
  assert.equal(optimized.touchesReserved, true);
  assert.equal(full.unknownReason, optimized.unknownReason);

  const fullTouch = full.touches[0];
  const optTouch = optimized.touches[0];
  assert.ok(fullTouch);
  assert.ok(optTouch);
  assert.equal(optTouch.root, fullTouch.root);
  assert.deepEqual(optTouch.keyPath, fullTouch.keyPath);
  assert.equal(optTouch.kind, fullTouch.kind);

  serverDoc.destroy();
  attackerDoc.destroy();
});
