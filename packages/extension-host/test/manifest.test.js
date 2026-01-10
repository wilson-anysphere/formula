const test = require("node:test");
const assert = require("node:assert/strict");

const { validateExtensionManifest, ManifestError } = require("../src/manifest");

test("manifest validation: required fields", () => {
  assert.throws(
    () => validateExtensionManifest({ version: "1.0.0" }, { engineVersion: "1.0.0" }),
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
        { engineVersion: "1.0.0" }
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
        { engineVersion: "2.0.0" }
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
        { engineVersion: "1.0.0" }
      ),
    /unknown command/
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
        { engineVersion: "1.0.0" }
      ),
    /Invalid permission/
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
        { engineVersion: "1.0.0" }
      ),
    /contributes\.configuration\.properties\.x\.setting\.type/
  );
});
