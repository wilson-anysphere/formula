const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs/promises");
const path = require("node:path");
const vm = require("node:vm");

test("shared/extension-manifest can execute in a browser-like environment (no require/module)", async () => {
  const semverCode = await fs.readFile(path.join(__dirname, "..", "semver-range", "index.js"), "utf8");
  const manifestCode = await fs.readFile(path.join(__dirname, "index.js"), "utf8");

  // Simulate a browser module environment: no `require`, no `module`, but `globalThis` exists.
  /** @type {Record<string, any>} */
  const sandbox = {};
  sandbox.globalThis = sandbox;

  function runModuleLike(code, filename) {
    // Both `shared/semver-range/index.js` and `shared/extension-manifest/index.js` are authored to
    // work as CommonJS in Node and as ESM in browsers. When we run them via `vm`, we wrap the
    // source to simulate module scoping so top-level `const` declarations don't collide.
    vm.runInNewContext(`(() => {\n${code}\n})();`, sandbox, { filename });
  }

  runModuleLike(semverCode, "shared/semver-range/index.js");
  const semverImpl = sandbox.__formula_semver_range__;
  assert.ok(semverImpl, "expected semver-range to register on globalThis");
  assert.equal(typeof semverImpl.satisfies, "function");

  runModuleLike(manifestCode, "shared/extension-manifest/index.js");

  const impl = sandbox.__formula_extension_manifest__;
  assert.ok(impl, "expected implementation to register on globalThis");
  assert.equal(typeof impl.validateExtensionManifest, "function");
  assert.equal(Object.prototype.toString.call(impl.VALID_PERMISSIONS), "[object Set]");
  assert.equal(typeof impl.ManifestError, "function");

  assert.throws(
    () => impl.validateExtensionManifest({ version: "1.0.0" }, { enforceEngine: false }),
    /must be a non-empty string/
  );
});
