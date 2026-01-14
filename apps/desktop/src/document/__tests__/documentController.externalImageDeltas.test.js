import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("applyExternalImageDeltas updates image store, emits change.imageDeltas, and does not create undo history", () => {
  const doc = new DocumentController();

  const beforeDepth = doc.getStackDepths();
  assert.equal(doc.isDirty, false);

  /** @type {any | null} */
  let lastChange = null;
  doc.on("change", (payload) => {
    lastChange = payload;
  });

  doc.applyExternalImageDeltas(
    [
      {
        imageId: "img1",
        before: null,
        after: { bytes: new Uint8Array([1, 2, 3]), mimeType: " image/png " },
      },
    ],
    { source: "hydration" },
  );

  const image = doc.getImage("img1");
  assert.ok(image);
  assert.equal(image?.mimeType, "image/png");
  assert.deepEqual(Array.from(image?.bytes ?? []), [1, 2, 3]);

  assert.deepEqual(doc.getStackDepths(), beforeDepth, "external deltas should not create undo history");
  assert.equal(doc.isDirty, true, "external deltas should mark the document dirty by default");

  assert.ok(lastChange, "expected a change event");
  assert.deepEqual(lastChange.imageDeltas, [{ imageId: "img1", before: null, after: { mimeType: "image/png", byteLength: 3 } }]);
});

test("applyExternalImageDeltas respects markDirty=false", () => {
  const doc = new DocumentController();
  doc.markSaved();
  assert.equal(doc.isDirty, false);

  doc.applyExternalImageDeltas(
    [
      {
        imageId: "img1",
        before: null,
        after: { bytes: new Uint8Array([9]), mimeType: "image/png" },
      },
    ],
    { source: "hydration", markDirty: false },
  );

  assert.ok(doc.getImage("img1"));
  assert.equal(doc.isDirty, false);
});

test("applyExternalImageDeltas ignores no-op deltas", () => {
  const doc = new DocumentController();
  doc.setImage("img1", { bytes: new Uint8Array([1, 2]), mimeType: "image/png" });
  doc.markSaved();
  assert.equal(doc.isDirty, false);

  let changeCount = 0;
  doc.on("change", () => {
    changeCount += 1;
  });

  doc.applyExternalImageDeltas(
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

test("applyExternalImageDeltas ignores invalid `after` payloads (does not delete existing images)", () => {
  const doc = new DocumentController();
  doc.setImage("img1", { bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" });
  doc.markSaved();
  assert.equal(doc.isDirty, false);

  let changeCount = 0;
  doc.on("change", () => {
    changeCount += 1;
  });

  doc.applyExternalImageDeltas(
    [
      {
        imageId: "img1",
        before: { bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" },
        // Invalid bytes payload; should be ignored rather than treated as a delete.
        after: { bytes: "not-bytes", mimeType: "image/png" },
      },
    ],
    { source: "collab" },
  );

  assert.equal(changeCount, 0);
  assert.equal(doc.isDirty, false);
  assert.deepEqual(Array.from(doc.getImage("img1")?.bytes ?? []), [1, 2, 3]);
});
