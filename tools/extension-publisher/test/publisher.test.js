const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const { packageExtension } = require("../src/publisher");

async function writeJson(filePath, value) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, JSON.stringify(value, null, 2), "utf8");
}

test("packageExtension rejects missing manifest main entrypoints", async (t) => {
  const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "formula-publish-missing-main-"));
  t.after(async () => {
    await fs.rm(tmp, { recursive: true, force: true });
  });

  await writeJson(path.join(tmp, "package.json"), {
    name: "missing-main",
    publisher: "publisher",
    version: "1.0.0",
    main: "./dist/extension.js",
    engines: { formula: "^1.0.0" }
  });

  await assert.rejects(
    () => packageExtension(tmp),
    /Did you forget to build the extension\?/i
  );
});

test("packageExtension rejects manifest main that escapes extensionDir", async (t) => {
  const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "formula-publish-escape-main-"));
  t.after(async () => {
    await fs.rm(tmp, { recursive: true, force: true });
  });

  await writeJson(path.join(tmp, "package.json"), {
    name: "escape-main",
    publisher: "publisher",
    version: "1.0.0",
    main: "../escape.js",
    engines: { formula: "^1.0.0" }
  });

  await assert.rejects(() => packageExtension(tmp), /resolve inside extensionDir/i);
});

