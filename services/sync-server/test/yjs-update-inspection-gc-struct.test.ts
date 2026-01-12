import assert from "node:assert/strict";
import test from "node:test";

import { inspectUpdate } from "../src/yjsUpdateInspection.js";
import { Y } from "./yjs-interop.ts";

test("yjs update inspection: GC detection does not rely on ctor.name", () => {
  const serverDoc = new Y.Doc();

  // Create a GC struct, but simulate bundler-renamed constructors by swapping the instance prototype.
  const gc = new (Y as any).GC(Y.createID(1, 0), 1);
  class RenamedGC extends gc.constructor {}
  Object.setPrototypeOf(gc, RenamedGC.prototype);
  assert.notEqual((gc as any).constructor?.name, "GC");

  // Create an Item-like struct with missing parent info that references the GC struct via origin.
  const item = {
    id: { client: 1, clock: 1 },
    length: 1,
    parent: null,
    parentSub: null,
    origin: (gc as any).id,
    rightOrigin: null,
  };

  const decoded = { structs: [item, gc], ds: { clients: new Map() } };

  // Monkey patch `decodeUpdate` so we can feed custom struct instances into the inspector.
  const originalDecodeUpdate = (Y as any).decodeUpdate;
  try {
    (Y as any).decodeUpdate = () => decoded;

    const res = inspectUpdate({
      ydoc: serverDoc,
      update: new Uint8Array([1, 2, 3]),
      reservedRootNames: new Set(),
      reservedRootPrefixes: [],
      maxTouches: 10,
    });

    assert.equal(res.touchesReserved, true);
    assert.equal(res.unknownReason, "origin_or_right_origin_is_gc");
    assert.equal(res.touches[0]?.kind, "gc");
  } finally {
    (Y as any).decodeUpdate = originalDecodeUpdate;
    serverDoc.destroy();
  }
});

