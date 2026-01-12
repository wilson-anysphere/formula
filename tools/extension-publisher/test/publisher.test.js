const test = require("node:test");
const assert = require("node:assert/strict");
const os = require("node:os");
const path = require("node:path");
const fs = require("node:fs/promises");

const crypto = require("node:crypto");

const { packageExtension, publishExtension } = require("../src/publisher");

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

test("publishExtension tolerates marketplaceUrl ending with /api (avoids /api/api/...)", async (t) => {
  const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "formula-publish-baseurl-"));
  t.after(async () => {
    await fs.rm(tmp, { recursive: true, force: true });
  });

  await writeJson(path.join(tmp, "package.json"), {
    name: "hello",
    publisher: "publisher",
    version: "1.0.0",
    main: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
  });
  await fs.mkdir(path.join(tmp, "dist"), { recursive: true });
  await fs.writeFile(path.join(tmp, "dist", "extension.js"), "module.exports = {};\n", "utf8");

  const { privateKey } = crypto.generateKeyPairSync("ed25519");
  const privateKeyPem = privateKey.export({ type: "pkcs8", format: "pem" });

  const calls = [];
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async (url) => {
    calls.push(String(url));
    return new Response(JSON.stringify({ id: "publisher.hello", version: "1.0.0" }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  };

  t.after(() => {
    globalThis.fetch = originalFetch;
  });

  const published = await publishExtension({
    extensionDir: tmp,
    marketplaceUrl: "https://marketplace.example.com/api",
    token: "publisher-token",
    privateKeyPemOrPath: privateKeyPem,
    formatVersion: 2,
  });

  assert.equal(published.id, "publisher.hello");
  assert.equal(published.version, "1.0.0");
  assert.ok(published.manifest && typeof published.manifest === "object");
  assert.equal(published.manifest.name, "hello");
  assert.equal(published.manifest.publisher, "publisher");
  assert.equal(published.manifest.version, "1.0.0");
  assert.equal(published.manifest.main, "./dist/extension.js");
  assert.deepEqual(calls, ["https://marketplace.example.com/api/publish-bin"]);
});

test("publishExtension strips query/hash from marketplaceUrl", async (t) => {
  const tmp = await fs.mkdtemp(path.join(os.tmpdir(), "formula-publish-baseurl-query-"));
  t.after(async () => {
    await fs.rm(tmp, { recursive: true, force: true });
  });

  await writeJson(path.join(tmp, "package.json"), {
    name: "hello",
    publisher: "publisher",
    version: "1.0.0",
    main: "./dist/extension.js",
    engines: { formula: "^1.0.0" },
  });
  await fs.mkdir(path.join(tmp, "dist"), { recursive: true });
  await fs.writeFile(path.join(tmp, "dist", "extension.js"), "module.exports = {};\n", "utf8");

  const { privateKey } = crypto.generateKeyPairSync("ed25519");
  const privateKeyPem = privateKey.export({ type: "pkcs8", format: "pem" });

  const calls = [];
  const originalFetch = globalThis.fetch;
  globalThis.fetch = async (url) => {
    calls.push(String(url));
    return new Response(JSON.stringify({ id: "publisher.hello", version: "1.0.0" }), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    });
  };

  t.after(() => {
    globalThis.fetch = originalFetch;
  });

  const published = await publishExtension({
    extensionDir: tmp,
    marketplaceUrl: "https://marketplace.example.com/api?x=y#frag",
    token: "publisher-token",
    privateKeyPemOrPath: privateKeyPem,
    formatVersion: 2,
  });

  assert.equal(published.id, "publisher.hello");
  assert.equal(published.version, "1.0.0");
  assert.deepEqual(calls, ["https://marketplace.example.com/api/publish-bin"]);
});
