import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { fileURLToPath } from "node:url";

test("utils/hash.js is browser-safe (no node:crypto import)", async () => {
  const mod = await import("../src/utils/hash.js");
  assert.equal(typeof mod.contentHash, "function");

  const digest1 = await mod.contentHash("hello");
  const digest2 = await mod.contentHash("hello");
  assert.match(digest1, /^[0-9a-f]+$/);
  assert.ok(digest1.length === 64 || digest1.length === 16);
  assert.equal(digest1, digest2);

  const sourcePath = fileURLToPath(new URL("../src/utils/hash.js", import.meta.url));
  const source = await readFile(sourcePath, "utf8");
  assert.equal(source.includes("node:crypto"), false);
});
