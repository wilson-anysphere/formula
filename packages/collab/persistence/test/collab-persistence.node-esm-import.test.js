import assert from "node:assert/strict";
import test from "node:test";

// Include explicit `.ts` import specifiers so the repo's node:test runner can
// automatically skip this suite when TypeScript execution isn't available.
import { FileCollabPersistence as FileFromTs } from "../src/file.ts";
import { IndexedDbCollabPersistence as IndexeddbFromTs } from "../src/indexeddb.ts";
import "../src/index.ts";

test("collab-persistence TS sources are importable under Node ESM when executing TS sources directly", async () => {
  const root = await import("@formula/collab-persistence");
  assert.ok(root && typeof root === "object");

  const file = await import("@formula/collab-persistence/file");
  const indexeddb = await import("@formula/collab-persistence/indexeddb");

  assert.equal(typeof file.FileCollabPersistence, "function");
  assert.equal(typeof indexeddb.IndexedDbCollabPersistence, "function");
  assert.equal(typeof FileFromTs, "function");
  assert.equal(typeof IndexeddbFromTs, "function");
});
