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

test("packageExtension rejects missing manifest module entrypoints", async (t) => {
  const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "formula-publish-missing-module-"));
  t.after(async () => {
    await fs.rm(tmp, { recursive: true, force: true });
  });

  await writeJson(path.join(tmp, "package.json"), {
    name: "missing-module",
    publisher: "publisher",
    version: "1.0.0",
    main: "./dist/extension.js",
    module: "./dist/extension.mjs",
    engines: { formula: "^1.0.0" }
  });

  await fs.mkdir(path.join(tmp, "dist"), { recursive: true });
  await fs.writeFile(path.join(tmp, "dist", "extension.js"), "module.exports = {};\n", "utf8");

  await assert.rejects(() => packageExtension(tmp), /module entrypoint is missing/i);
});

test("packageExtension rejects missing manifest browser entrypoints", async (t) => {
  const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "formula-publish-missing-browser-"));
  t.after(async () => {
    await fs.rm(tmp, { recursive: true, force: true });
  });

  await writeJson(path.join(tmp, "package.json"), {
    name: "missing-browser",
    publisher: "publisher",
    version: "1.0.0",
    main: "./dist/extension.js",
    browser: "./dist/extension.mjs",
    engines: { formula: "^1.0.0" }
  });

  await fs.mkdir(path.join(tmp, "dist"), { recursive: true });
  await fs.writeFile(path.join(tmp, "dist", "extension.js"), "module.exports = {};\n", "utf8");

  await assert.rejects(() => packageExtension(tmp), /browser entrypoint is missing/i);
});
