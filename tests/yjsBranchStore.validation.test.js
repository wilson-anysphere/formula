import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { YjsBranchStore } from "../packages/versioning/branches/src/store/YjsBranchStore.js";

test("YjsBranchStore: rejects invalid chunkSize/maxChunksPerTransaction values", () => {
  const ydoc = new Y.Doc();

  // chunkSize
  assert.throws(() => new YjsBranchStore({ ydoc, chunkSize: 0 }), /chunkSize/i);
  assert.throws(() => new YjsBranchStore({ ydoc, chunkSize: -1 }), /chunkSize/i);
  // @ts-expect-error - runtime validation
  assert.throws(() => new YjsBranchStore({ ydoc, chunkSize: Number.NaN }), /chunkSize/i);
  assert.throws(() => new YjsBranchStore({ ydoc, chunkSize: 1.5 }), /chunkSize/i);
  // @ts-expect-error - runtime validation
  assert.throws(() => new YjsBranchStore({ ydoc, chunkSize: Number.POSITIVE_INFINITY }), /chunkSize/i);

  // maxChunksPerTransaction
  assert.throws(() => new YjsBranchStore({ ydoc, maxChunksPerTransaction: 0 }), /maxChunksPerTransaction/i);
  assert.throws(() => new YjsBranchStore({ ydoc, maxChunksPerTransaction: -1 }), /maxChunksPerTransaction/i);
  // @ts-expect-error - runtime validation
  assert.throws(
    () => new YjsBranchStore({ ydoc, maxChunksPerTransaction: Number.NaN }),
    /maxChunksPerTransaction/i,
  );
  assert.throws(() => new YjsBranchStore({ ydoc, maxChunksPerTransaction: 1.5 }), /maxChunksPerTransaction/i);
  // @ts-expect-error - runtime validation
  assert.throws(
    () => new YjsBranchStore({ ydoc, maxChunksPerTransaction: Number.POSITIVE_INFINITY }),
    /maxChunksPerTransaction/i,
  );
});

