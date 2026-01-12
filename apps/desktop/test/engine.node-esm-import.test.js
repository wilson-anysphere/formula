import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when `--experimental-strip-types` is not available.
import { createEngineClient as createEngineClientFromTs } from "../../../packages/engine/src/index.ts";

test("engine is importable under Node ESM when executing TS sources (strip-types)", async () => {
  const mod = await import("@formula/engine");

  assert.equal(typeof mod.createEngineClient, "function");
  assert.equal(typeof mod.EngineWorker, "function");
  assert.equal(typeof createEngineClientFromTs, "function");
});

