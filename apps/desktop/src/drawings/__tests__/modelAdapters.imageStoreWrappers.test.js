import assert from "node:assert/strict";
import test from "node:test";

import { convertModelImageStoreToUiImageStore } from "../modelAdapters.ts";

test("convertModelImageStoreToUiImageStore tolerates singleton-wrapped byte payloads (interop)", () => {
  const modelImages = {
    images: {
      // Wrapper around a number[].
      "wrapped-array.png": { bytes: { 0: [1, 2, 3] }, content_type: "image/png" },
      // Wrapper around a base64 string.
      "wrapped-b64.png": { bytes: { 0: "AQID" }, content_type: "image/png" }, // [1,2,3]
      // Wrapper around a Node Buffer JSON representation.
      "wrapped-buf.png": { bytes: { 0: { type: "Buffer", data: [9, 10, 11] } }, content_type: "image/png" },
      // Wrapper around a numeric-key object encoding.
      "wrapped-obj.png": { bytes: { 0: { 0: 7, 1: 8 } }, content_type: "image/png" },
    },
  };

  const store = convertModelImageStoreToUiImageStore(modelImages);
  assert.deepEqual(store.get("wrapped-array.png")?.bytes, new Uint8Array([1, 2, 3]));
  assert.deepEqual(store.get("wrapped-b64.png")?.bytes, new Uint8Array([1, 2, 3]));
  assert.deepEqual(store.get("wrapped-buf.png")?.bytes, new Uint8Array([9, 10, 11]));
  assert.deepEqual(store.get("wrapped-obj.png")?.bytes, new Uint8Array([7, 8]));
});

