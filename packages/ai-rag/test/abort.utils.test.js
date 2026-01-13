import assert from "node:assert/strict";
import test from "node:test";

import { awaitWithAbort, createAbortError, throwIfAborted } from "../src/utils/abort.js";

test("createAbortError returns an Error named AbortError", () => {
  assert.equal(createAbortError().name, "AbortError");
});

test("throwIfAborted throws an AbortError when signal is aborted", () => {
  const ac = new AbortController();
  ac.abort();

  assert.throws(() => throwIfAborted(ac.signal), { name: "AbortError" });
});

test("awaitWithAbort resolves the wrapped promise when not aborted", async () => {
  const ac = new AbortController();
  assert.equal(await awaitWithAbort(Promise.resolve(123), ac.signal), 123);
});

test("awaitWithAbort rejects with AbortError when signal is already aborted", async () => {
  const ac = new AbortController();
  ac.abort();

  await assert.rejects(awaitWithAbort(Promise.resolve(123), ac.signal), { name: "AbortError" });
});

test("awaitWithAbort rejects promptly when signal aborts before promise resolves", async () => {
  const ac = new AbortController();

  /** @type {(value: number) => void} */
  let resolve;
  const promise = new Promise((r) => {
    resolve = r;
  });

  const wrapped = awaitWithAbort(promise, ac.signal);
  ac.abort();
  resolve(123);

  await assert.rejects(wrapped, { name: "AbortError" });
});

test("awaitWithAbort forwards the underlying rejection when promise rejects", async () => {
  const ac = new AbortController();
  const error = new Error("boom");
  await assert.rejects(awaitWithAbort(Promise.reject(error), ac.signal), error);
});
