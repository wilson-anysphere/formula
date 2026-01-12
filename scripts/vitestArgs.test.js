import test from "node:test";
import assert from "node:assert/strict";

import { normalizeVitestArgs } from "./vitestArgs.mjs";

test("normalizeVitestArgs strips bare `--` delimiters", () => {
  assert.deepEqual(normalizeVitestArgs(["--"]), []);
  assert.deepEqual(normalizeVitestArgs(["foo", "--", "bar"]), ["foo", "bar"]);
  assert.deepEqual(normalizeVitestArgs(["--", "foo", "--", "bar", "--"]), ["foo", "bar"]);
});

test("normalizeVitestArgs rewrites `--silent` to an explicit boolean", () => {
  assert.deepEqual(normalizeVitestArgs(["--silent"]), ["--silent=true"]);
  assert.deepEqual(normalizeVitestArgs(["--silent", "apps/foo.test.ts"]), ["--silent=true", "apps/foo.test.ts"]);
  assert.deepEqual(normalizeVitestArgs(["--", "--silent", "apps/foo.test.ts"]), ["--silent=true", "apps/foo.test.ts"]);
});

test("normalizeVitestArgs preserves explicit `--silent=<bool>` forms", () => {
  assert.deepEqual(normalizeVitestArgs(["--silent=true", "apps/foo.test.ts"]), ["--silent=true", "apps/foo.test.ts"]);
  assert.deepEqual(normalizeVitestArgs(["--silent=false", "apps/foo.test.ts"]), ["--silent=false", "apps/foo.test.ts"]);
});

