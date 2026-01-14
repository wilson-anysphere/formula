import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { readdir } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { stripComments } from "../../../apps/desktop/test/sourceTextUtils.js";

/**
 * @param {string} dir
 * @param {string[]} out
 */
async function collectJsFiles(dir, out) {
  const entries = await readdir(dir, { withFileTypes: true });
  for (const entry of entries) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await collectJsFiles(full, out);
      continue;
    }
    if (!entry.isFile()) continue;
    if (!entry.name.endsWith(".js")) continue;
    out.push(full);
  }
}

test("utils/hash.js is browser-safe (no Node builtin crypto import)", async () => {
  const mod = await import("../src/utils/hash.js");
  assert.equal(typeof mod.contentHash, "function");

  const digest1 = await mod.contentHash("hello");
  const digest2 = await mod.contentHash("hello");
  assert.match(digest1, /^[0-9a-f]+$/);
  assert.ok(digest1.length === 64 || digest1.length === 16);
  assert.equal(digest1, digest2);

  // When WebCrypto is available *and succeeds* we should get a real SHA-256 digest.
  // Some runtimes expose crypto.subtle but throw at runtime (in which case
  // contentHash intentionally falls back).
  if (globalThis.crypto?.subtle && digest1.length === 64) {
    assert.equal(
      digest1,
      "2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824"
    );
  }
 
  const sourcePath = fileURLToPath(new URL("../src/utils/hash.js", import.meta.url));
  const source = stripComments(await readFile(sourcePath, "utf8"));
  const nodeCryptoSpecifier = ["node", "crypto"].join(":");
  assert.equal(source.includes(nodeCryptoSpecifier), false);

  // Ensure we didn't accidentally re-introduce a Node crypto import elsewhere in ai-rag.
  // Avoid using `new URL()` with `import.meta.url` on the src directory here so
  // `scripts/run-node-tests.mjs` doesn't treat it as a module dependency on
  // `src/index.js` (which conditionally depends on external packages like `sql.js`).
  const srcDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../src");
  /** @type {string[]} */
  const files = [];
  await collectJsFiles(srcDir, files);
  for (const file of files) {
    const contents = stripComments(await readFile(file, "utf8"));
    assert.equal(
      contents.includes(nodeCryptoSpecifier),
      false,
      `Unexpected Node crypto import specifier in ${path.relative(srcDir, file)}`
    );
  }
});

test("contentHash falls back when WebCrypto is unavailable", async () => {
  const mod = await import("../src/utils/hash.js");
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, "crypto");
  if (descriptor && descriptor.configurable !== true) {
    // Can't override in this runtime; just ensure it works.
    const digest = await mod.contentHash("hello");
    assert.match(digest, /^[0-9a-f]+$/);
    assert.ok(digest.length === 64 || digest.length === 16);
    return;
  }

  try {
    Object.defineProperty(globalThis, "crypto", { value: undefined, configurable: true });
    const digest = await mod.contentHash("hello");
    assert.match(digest, /^[0-9a-f]{16}$/);
  } finally {
    if (descriptor) Object.defineProperty(globalThis, "crypto", descriptor);
    else delete globalThis.crypto;
  }
});

test("contentHash falls back when WebCrypto digest throws", async () => {
  const mod = await import("../src/utils/hash.js");
  const descriptor = Object.getOwnPropertyDescriptor(globalThis, "crypto");
  if (descriptor && descriptor.configurable !== true) {
    // Can't override in this runtime; skip.
    return;
  }

  try {
    Object.defineProperty(globalThis, "crypto", {
      value: {
        subtle: {
          async digest() {
            throw new Error("boom");
          },
        },
      },
      configurable: true,
    });

    const digest = await mod.contentHash("hello");
    assert.match(digest, /^[0-9a-f]{16}$/);
  } finally {
    if (descriptor) Object.defineProperty(globalThis, "crypto", descriptor);
    else delete globalThis.crypto;
  }
});

test("contentHash works when TextEncoder is unavailable", async () => {
  const modNormal = await import(`../src/utils/hash.js?normal-text-encoder=${Date.now()}`);
  const surrogateDigest = await modNormal.contentHash("\uD800");

  const descriptor = Object.getOwnPropertyDescriptor(globalThis, "TextEncoder");
  if (descriptor && descriptor.configurable !== true) {
    // Can't override in this runtime; skip.
    return;
  }

  try {
    Object.defineProperty(globalThis, "TextEncoder", { value: undefined, configurable: true, writable: true });
    const mod = await import(`../src/utils/hash.js?no-text-encoder=${Date.now()}`);
    const digest = await mod.contentHash("hello");
    const surrogateDigestNoEncoder = await mod.contentHash("\uD800");
    assert.equal(surrogateDigestNoEncoder, surrogateDigest);
    if (globalThis.crypto?.subtle) {
      assert.equal(
        digest,
        "2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824"
      );
    } else {
      assert.match(digest, /^[0-9a-f]{16}$/);
    }
  } finally {
    if (descriptor) Object.defineProperty(globalThis, "TextEncoder", descriptor);
    else delete globalThis.TextEncoder;
  }
});
