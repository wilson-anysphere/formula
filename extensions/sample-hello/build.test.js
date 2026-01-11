const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const fs = require("node:fs/promises");

const { build, DIST_ESM_PATH } = require("./build");

test("sample-hello: dist entrypoints are in sync with src", async () => {
  await build({ check: true });

  // Smoke test: ensure both entrypoints are syntactically valid. We intentionally
  // don't execute them here because the extension runtime relies on the host worker
  // injecting @formula/extension-api.
  const cjsPath = path.join(__dirname, "dist", "extension.js");
  for (const entrypoint of [cjsPath, DIST_ESM_PATH]) {
    const result = spawnSync(process.execPath, ["--check", entrypoint], { encoding: "utf8" });
    assert.equal(
      result.status,
      0,
      result.stderr || result.stdout || `node --check failed for ${entrypoint}`
    );
  }

  const [cjsSource, esmSource] = await Promise.all([
    fs.readFile(cjsPath, "utf8"),
    fs.readFile(DIST_ESM_PATH, "utf8")
  ]);

  assert.match(cjsSource, /require\(["']@formula\/extension-api["']\)/);
  assert.match(cjsSource, /module\.exports\s*=\s*\{/);
  assert.match(cjsSource, /\bactivate\b/);

  assert.match(esmSource, /import\s+\*\s+as\s+formula\s+from\s+["']@formula\/extension-api["']/);
  assert.match(esmSource, /export\s*\{\s*activate\s*\}\s*;/);
  assert.ok(!/module\.exports/.test(esmSource), "ESM build should not use module.exports");
});
