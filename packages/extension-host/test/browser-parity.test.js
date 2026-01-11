const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { pathToFileURL } = require("node:url");

const nodeHostPkg = require("../src");
const nodeManifestPkg = require("../src/manifest");
const sharedManifestPkg = require("../../../shared/extension-manifest");

test("browser host parity: API_PERMISSIONS matches Node ExtensionHost", async () => {
  const browserModuleUrl = pathToFileURL(
    path.join(__dirname, "..", "src", "browser", "index.mjs")
  ).href;
  const browserPkg = await import(browserModuleUrl);

  const nodePerms = nodeHostPkg.API_PERMISSIONS;
  const browserPerms = browserPkg.API_PERMISSIONS;

  assert.deepEqual(Object.keys(browserPerms).sort(), Object.keys(nodePerms).sort());

  for (const key of Object.keys(nodePerms)) {
    assert.deepEqual(browserPerms[key], nodePerms[key], `Permission mapping drift for ${key}`);
  }
});

test("browser host parity: manifest validation is shared (no drift)", async () => {
  const browserManifestUrl = pathToFileURL(
    path.join(__dirname, "..", "src", "browser", "manifest.mjs")
  ).href;
  const browserManifestPkg = await import(browserManifestUrl);

  assert.strictEqual(nodeManifestPkg.VALID_PERMISSIONS, sharedManifestPkg.VALID_PERMISSIONS);
  assert.strictEqual(nodeManifestPkg.validateExtensionManifest, sharedManifestPkg.validateExtensionManifest);
  assert.strictEqual(nodeManifestPkg.ManifestError, sharedManifestPkg.ManifestError);

  assert.strictEqual(browserManifestPkg.VALID_PERMISSIONS, sharedManifestPkg.VALID_PERMISSIONS);
  assert.strictEqual(browserManifestPkg.validateExtensionManifest, sharedManifestPkg.validateExtensionManifest);
  assert.strictEqual(browserManifestPkg.ManifestError, sharedManifestPkg.ManifestError);
});
