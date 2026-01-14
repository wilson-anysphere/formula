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

test("manifest validation: engine supports compound semver ranges", () => {
  assert.doesNotThrow(() =>
    validateExtensionManifest(
      {
        name: "x",
        version: "1.0.0",
        publisher: "p",
        main: "./dist/extension.js",
        engines: { formula: ">=1.0.0 <2.0.0" }
      },
      { engineVersion: "1.5.0", enforceEngine: true }
    )
  );
});

test("manifest validation: engine supports OR (||) semver ranges", () => {
  assert.doesNotThrow(() =>
    validateExtensionManifest(
      {
        name: "x",
        version: "1.0.0",
        publisher: "p",
        main: "./dist/extension.js",
        engines: { formula: "<1.0.0 || >=2.0.0" }
      },
      { engineVersion: "2.1.0", enforceEngine: true }
    )
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

test("manifest validation: contributes.commands supports optional description + keywords", () => {
  assert.doesNotThrow(() =>
    validateExtensionManifest(
      {
        name: "x",
        version: "1.0.0",
        publisher: "p",
        main: "./dist/extension.js",
        engines: { formula: "^1.0.0" },
        activationEvents: ["onCommand:test.cmd"],
        permissions: ["ui.commands"],
        contributes: {
          commands: [
            {
              command: "test.cmd",
              title: "Test Command",
              description: "A helpful subtitle shown in the command palette",
              keywords: ["hello", "world"]
            }
          ]
        }
      },
      { engineVersion: "1.0.0", enforceEngine: true }
    )
  );
});

test("manifest validation: contributes.commands keywords must be a string array", () => {
  assert.throws(
    () =>
      validateExtensionManifest(
        {
          name: "x",
          version: "1.0.0",
          publisher: "p",
          main: "./dist/extension.js",
          engines: { formula: "^1.0.0" },
          activationEvents: ["onCommand:test.cmd"],
          permissions: ["ui.commands"],
          contributes: {
            commands: [
              {
                command: "test.cmd",
                title: "Test Command",
                // @ts-expect-error - intentional invalid manifest type
                keywords: "not-an-array"
              }
            ]
          }
        },
        { engineVersion: "1.0.0", enforceEngine: true }
      ),
    /contributes\.commands\[0\]\.keywords must be an array/
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

test("manifest validation: module/browser entrypoints must be non-empty strings when present", () => {
  assert.throws(
    () =>
      validateExtensionManifest(
        {
          name: "x",
          version: "1.0.0",
          publisher: "p",
          main: "./dist/extension.js",
          module: "   ",
          engines: { formula: "^1.0.0" }
        },
        { engineVersion: "1.0.0", enforceEngine: true }
      ),
    /module must be a non-empty string/
  );

  assert.throws(
    () =>
      validateExtensionManifest(
        {
          name: "x",
          version: "1.0.0",
          publisher: "p",
          main: "./dist/extension.js",
          browser: "",
          engines: { formula: "^1.0.0" }
        },
        { engineVersion: "1.0.0", enforceEngine: true }
      ),
    /browser must be a non-empty string/
  );
});

test("manifest validation: main entrypoint must be CommonJS (.js/.cjs)", () => {
  assert.throws(
    () =>
      validateExtensionManifest(
        {
          name: "x",
          version: "1.0.0",
          publisher: "p",
          main: "./dist/extension.mjs",
          engines: { formula: "^1.0.0" }
        },
        { engineVersion: "1.0.0", enforceEngine: true }
      ),
    /main entrypoint must end with/i
  );
});

test("manifest validation: trims canonical string fields (ids, entrypoints, activation events, permissions)", () => {
  const validated = validateExtensionManifest(
    {
      name: " x ",
      version: " 1.0.0 ",
      publisher: " p ",
      main: " ./dist/extension.js ",
      engines: { formula: " ^1.0.0 " },
      activationEvents: [" onCommand: test.cmd "],
      permissions: [" ui.commands "],
      contributes: {
        commands: [
          {
            command: " test.cmd ",
            title: " Test Command ",
            keywords: [" hello ", " world "]
          }
        ]
      }
    },
    { engineVersion: "1.0.0", enforceEngine: true }
  );

  assert.equal(validated.name, "x");
  assert.equal(validated.version, "1.0.0");
  assert.equal(validated.publisher, "p");
  assert.equal(validated.main, "./dist/extension.js");
  assert.deepEqual(validated.activationEvents, ["onCommand:test.cmd"]);
  assert.deepEqual(validated.permissions, ["ui.commands"]);
  assert.equal(validated.engines.formula, "^1.0.0");

  assert.equal(validated.contributes.commands.length, 1);
  assert.equal(validated.contributes.commands[0].command, "test.cmd");
  assert.equal(validated.contributes.commands[0].title, "Test Command");
  assert.deepEqual(validated.contributes.commands[0].keywords, ["hello", "world"]);
});

test("manifest validation: trims permission object keys (while preserving values)", () => {
  const validated = validateExtensionManifest(
    {
      name: "x",
      version: "1.0.0",
      publisher: "p",
      main: "./dist/extension.js",
      engines: { formula: "^1.0.0" },
      permissions: [{ " network ": { mode: "allowlist", hosts: ["example.com"] } }]
    },
    { enforceEngine: false }
  );

  assert.deepEqual(validated.permissions, [{ network: { mode: "allowlist", hosts: ["example.com"] } }]);
});
