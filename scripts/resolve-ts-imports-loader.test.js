import assert from "node:assert/strict";
import test from "node:test";

// Include an explicit `.ts` specifier so `scripts/run-node-tests.mjs` can skip this
// suite when running under a Node version that doesn't support `--experimental-strip-types`.
import { valueFromBar } from "./__fixtures__/resolve-ts-imports/foo.ts";

test("node:test runner resolves bundler-style './bar.js' specifiers to '.ts' sources", () => {
  assert.equal(valueFromBar(), 42);
});

