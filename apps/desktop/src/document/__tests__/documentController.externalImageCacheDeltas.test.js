import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("applyExternalImageCacheDeltas updates cache, emits change.imageDeltas, and does not create undo history or mark dirty", () => {
  const doc = new DocumentController();

  const beforeDepth = doc.getStackDepths();
  assert.equal(doc.isDirty, false);

  /** @type {any | null} */
  let lastChange = null;
  doc.on("change", (payload) => {
    lastChange = payload;
  });

  doc.applyExternalImageCacheDeltas(
    [
      {
        imageId: "img1",
        before: null,
        after: { bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" },
      },
    ],
    { source: "hydration" },
  );

  const image = doc.getImage("img1");
  assert.ok(image);
  assert.equal(image?.mimeType, "image/png");
  assert.deepEqual(Array.from(image?.bytes ?? []), [1, 2, 3]);

  assert.deepEqual(doc.getStackDepths(), beforeDepth, "cache deltas should not create undo history");
  assert.equal(doc.isDirty, false, "cache deltas should not mark the document dirty by default");

  assert.ok(lastChange, "expected a change event");
  assert.equal(lastChange.source, "hydration");
  assert.deepEqual(lastChange.imageDeltas, [{ imageId: "img1", before: null, after: { mimeType: "image/png", byteLength: 3 } }]);

  const snapshot = JSON.parse(new TextDecoder().decode(doc.encodeState()));
  assert.equal(snapshot.images, undefined, "cached images should not be serialized into snapshots");
});

test("applyExternalImageCacheDeltas respects markDirty=true", () => {
  const doc = new DocumentController();
  doc.markSaved();
  assert.equal(doc.isDirty, false);

  doc.applyExternalImageCacheDeltas(
    [
      {
        imageId: "img1",
        before: null,
        after: { bytes: new Uint8Array([9]), mimeType: "image/png" },
      },
    ],
    { source: "hydration", markDirty: true },
  );

  assert.ok(doc.getImage("img1"));
  assert.equal(doc.isDirty, true);
});

test("applyExternalImageCacheDeltas ignores no-op deltas", () => {
  const doc = new DocumentController();
  doc.applyExternalImageCacheDeltas(
    [
      {
        imageId: "img1",
        before: null,
        after: { bytes: new Uint8Array([1, 2]), mimeType: "image/png" },
      },
    ],
    { source: "seed" },
  );
  doc.markSaved();
  assert.equal(doc.isDirty, false);

  let changeCount = 0;
  doc.on("change", () => {
    changeCount += 1;
  });

  doc.applyExternalImageCacheDeltas(
    [
      {
        imageId: "img1",
        before: { bytes: new Uint8Array([1, 2]), mimeType: "image/png" },
        after: { bytes: new Uint8Array([1, 2]), mimeType: "image/png" },
      },
    ],
    { source: "collab" },
  );

  assert.equal(changeCount, 0);
  assert.equal(doc.isDirty, false);
});

