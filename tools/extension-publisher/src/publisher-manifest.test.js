const test = require("node:test");
const assert = require("node:assert/strict");

const { validateManifest } = require("./publisher");

test("extension-publisher: validateManifest rejects invalid permissions", () => {
  assert.throws(
    () =>
      validateManifest({
        name: "x",
        version: "1.0.0",
        publisher: "p",
        main: "./dist/extension.js",
        engines: { formula: "^1.0.0" },
        permissions: ["totally.not.real"],
      }),
    /invalid permission/i
  );
});

test("extension-publisher: validateManifest rejects activationEvents referencing unknown commands", () => {
  assert.throws(
    () =>
      validateManifest({
        name: "x",
        version: "1.0.0",
        publisher: "p",
        main: "./dist/extension.js",
        engines: { formula: "^1.0.0" },
        activationEvents: ["onCommand:missing.command"],
        contributes: { commands: [] },
      }),
    /unknown command/i
  );
});

