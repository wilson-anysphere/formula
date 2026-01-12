import assert from "node:assert/strict";
import test from "node:test";

import { entryValue } from "./__tests__/fixtures/extensionless/entry.ts";

test("resolve-ts-imports-loader resolves extensionless TS imports (./foo -> ./foo.ts)", () => {
  assert.equal(entryValue, 42);
});

