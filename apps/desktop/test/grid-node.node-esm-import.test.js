import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` import specifier so the repo's node:test runner can
// automatically skip this suite when `--experimental-strip-types` is not available.
import { CanvasGridRenderer as RendererFromTs } from "../../../packages/grid/src/node.ts";

test("grid/node is importable under Node ESM when executing TS sources (strip-types)", async () => {
  const mod = await import("@formula/grid/node");

  assert.equal(typeof mod.CanvasGridRenderer, "function");
  assert.equal(typeof mod.VirtualScrollManager, "function");
  assert.equal(typeof mod.DirtyRegionTracker, "function");

  // Sanity: the TS source import resolves and matches the package export type.
  assert.equal(typeof RendererFromTs, "function");
});

