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

test("permission schema: API_PERMISSIONS only uses known manifest permissions", () => {
  const unknown = new Set();
  for (const perms of Object.values(nodeHostPkg.API_PERMISSIONS)) {
    for (const perm of perms) {
      if (!nodeManifestPkg.VALID_PERMISSIONS.has(perm)) unknown.add(perm);
    }
  }
  assert.deepEqual([...unknown].sort(), []);
});

test("browser host parity: permission management methods are available on both hosts", async () => {
  const browserModuleUrl = pathToFileURL(
    path.join(__dirname, "..", "src", "browser", "index.mjs")
  ).href;
  const browserPkg = await import(browserModuleUrl);

  const browserProto = browserPkg.BrowserExtensionHost?.prototype;
  const nodeProto = nodeHostPkg.ExtensionHost?.prototype;
  assert.ok(browserProto, "Expected BrowserExtensionHost prototype");
  assert.ok(nodeProto, "Expected ExtensionHost prototype");

  const methods = [
    "getGrantedPermissions",
    "revokePermissions",
    "resetPermissions",
    "resetAllPermissions"
  ];

  for (const method of methods) {
    assert.equal(
      typeof browserProto[method],
      "function",
      `BrowserExtensionHost missing method: ${method}`
    );
    assert.equal(typeof nodeProto[method], "function", `ExtensionHost missing method: ${method}`);
  }
});
