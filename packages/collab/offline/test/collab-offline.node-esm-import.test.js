import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when `--experimental-strip-types` is not available.
import { attachOfflinePersistence as attachFromTs } from "../src/index.node.ts";

test("collab-offline is importable under Node ESM when executing TS sources (strip-types)", async () => {
  const mod = await import("@formula/collab-offline");

  assert.equal(typeof mod.attachOfflinePersistence, "function");
  assert.equal(typeof attachFromTs, "function");
});

