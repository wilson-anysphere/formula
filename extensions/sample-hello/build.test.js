const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const fs = require("node:fs/promises");

const { checkExtension } = require("../../tools/extension-builder/src/builder");

test("sample-hello: dist entrypoints are in sync with src", async () => {
  await checkExtension(__dirname, { strict: true });

  // Smoke test: ensure both entrypoints are syntactically valid. We intentionally
  // don't execute them here because the extension runtime relies on the host worker
  // injecting @formula/extension-api.
  const cjsPath = path.join(__dirname, "dist", "extension.js");
  const esmPath = path.join(__dirname, "dist", "extension.mjs");
  for (const entrypoint of [cjsPath, esmPath]) {
    const result = spawnSync(process.execPath, ["--check", entrypoint], { encoding: "utf8" });
    assert.equal(
      result.status,
      0,
      result.stderr || result.stdout || `node --check failed for ${entrypoint}`
    );
  }

  const [cjsSource, esmSource] = await Promise.all([
    fs.readFile(cjsPath, "utf8"),
    fs.readFile(esmPath, "utf8")
  ]);

  assert.match(cjsSource, /require\(["']@formula\/extension-api["']\)/);
  assert.match(cjsSource, /module\.exports/);
  assert.match(cjsSource, /\bactivate\b/);

  assert.match(esmSource, /@formula\/extension-api/);
  assert.match(esmSource, /\bactivate\b/);
  assert.match(esmSource, /\bexport\s+default\b/);
});
