const test = require("node:test");
const assert = require("node:assert/strict");

const { validateExtensionManifest, ManifestError } = require("../src/manifest");

test("manifest validation: required fields", () => {
  assert.throws(
    () => validateExtensionManifest({ version: "1.0.0" }, { engineVersion: "1.0.0", enforceEngine: true }),
    ManifestError
  );
});

test("manifest validation: invalid semver version rejected", () => {
  assert.throws(
    () =>
      validateExtensionManifest(
        {
          name: "x",
          version: "not-a-version",
          publisher: "p",
          main: "./dist/extension.js",
          engines: { formula: "^1.0.0" }
        },
        { engineVersion: "1.0.0", enforceEngine: true }
      ),
    /Invalid version/
  );
});

test("manifest validation: engine mismatch rejected", () => {
  assert.throws(
    () =>
      validateExtensionManifest(
        {
          name: "x",
          version: "1.0.0",
          publisher: "p",
          main: "./dist/extension.js",
          engines: { formula: "^1.0.0" }
        },
        { engineVersion: "2.0.0", enforceEngine: true }
      ),
    /engine mismatch/
  );
});

test("manifest validation: activation event must reference contributed command", () => {
  assert.throws(
    () =>
      validateExtensionManifest(
        {
          name: "x",
          version: "1.0.0",
          publisher: "p",
          main: "./dist/extension.js",
          engines: { formula: "^1.0.0" },
          activationEvents: ["onCommand:missing.command"],
          contributes: { commands: [] }
        },
        { engineVersion: "1.0.0", enforceEngine: true }
      ),
    /unknown command/
  );
});

test("manifest validation: activation event must reference contributed data connector", () => {
  assert.throws(
    () =>
      validateExtensionManifest(
        {
          name: "x",
          version: "1.0.0",
          publisher: "p",
          main: "./dist/extension.js",
          engines: { formula: "^1.0.0" },
          activationEvents: ["onDataConnector:missing.connector"],
          contributes: { dataConnectors: [] }
        },
        { engineVersion: "1.0.0", enforceEngine: true }
      ),
    /unknown data connector/
  );
});

test("manifest validation: invalid permission rejected", () => {
  assert.throws(
    () =>
      validateExtensionManifest(
        {
          name: "x",
          version: "1.0.0",
          publisher: "p",
          main: "./dist/extension.js",
          engines: { formula: "^1.0.0" },
          permissions: ["totally.not.real"]
        },
        { engineVersion: "1.0.0", enforceEngine: true }
      ),
    /Invalid permission/
  );
});

test("manifest validation: permission objects with a single key are allowed", () => {
  assert.doesNotThrow(() =>
    validateExtensionManifest(
      {
        name: "x",
        version: "1.0.0",
        publisher: "p",
        main: "./dist/extension.js",
        engines: { formula: "^1.0.0" },
        permissions: [{ network: { mode: "allowlist", hosts: ["example.com"] } }]
      },
      { engineVersion: "1.0.0", enforceEngine: true }
    )
  );
});

test("manifest validation: permission objects must have exactly one key", () => {
  assert.throws(
    () =>
      validateExtensionManifest(
        {
          name: "x",
          version: "1.0.0",
          publisher: "p",
          main: "./dist/extension.js",
          engines: { formula: "^1.0.0" },
          permissions: [{ network: true, storage: true }]
        },
        { engineVersion: "1.0.0", enforceEngine: true }
      ),
    /must be a permission string or an object with a single permission key/
  );
});

test("manifest validation: configuration must declare typed properties", () => {
  assert.throws(
    () =>
      validateExtensionManifest(
        {
          name: "x",
          version: "1.0.0",
          publisher: "p",
          main: "./dist/extension.js",
          engines: { formula: "^1.0.0" },
          contributes: {
            configuration: {
              title: "X",
              properties: {
                "x.setting": {
                  description: "missing type"
                }
              }
            }
          }
        },
        { engineVersion: "1.0.0", enforceEngine: true }
      ),
    /contributes\.configuration\.properties\.x\.setting\.type/
  );
});

test("manifest validation: module/browser entrypoints must be strings when present", () => {
  assert.throws(
    () =>
      validateExtensionManifest(
        {
          name: "x",
          version: "1.0.0",
          publisher: "p",
          main: "./dist/extension.js",
          module: { not: "a string" },
          engines: { formula: "^1.0.0" }
        },
        { engineVersion: "1.0.0", enforceEngine: true }
      ),
    /module must be a string/
  );

  assert.throws(
    () =>
      validateExtensionManifest(
        {
          name: "x",
          version: "1.0.0",
          publisher: "p",
          main: "./dist/extension.js",
          browser: { not: "a string" },
          engines: { formula: "^1.0.0" }
        },
        { engineVersion: "1.0.0", enforceEngine: true }
      ),
    /browser must be a string/
  );
});
