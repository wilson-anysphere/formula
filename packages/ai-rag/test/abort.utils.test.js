import assert from "node:assert/strict";
import test from "node:test";

import { createAbortError, throwIfAborted } from "../src/utils/abort.js";

test("createAbortError returns an Error named AbortError", () => {
  assert.equal(createAbortError().name, "AbortError");
});

test("throwIfAborted throws an AbortError when signal is aborted", () => {
  const ac = new AbortController();
  ac.abort();

  assert.throws(() => throwIfAborted(ac.signal), { name: "AbortError" });
});

