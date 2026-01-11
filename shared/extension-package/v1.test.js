const test = require("node:test");
const assert = require("node:assert/strict");
const zlib = require("node:zlib");

const { readExtensionPackageV1 } = require("./v1");

test("v1 reader rejects gzip bombs via uncompressed size limit", () => {
  // A large, highly-compressible payload. The reader should reject based on
  // uncompressed size before it even reaches JSON parsing.
  const huge = Buffer.alloc(50 * 1024 * 1024 + 1, 0x61);
  const gz = zlib.gzipSync(huge, { level: 9 });

  assert.throws(() => readExtensionPackageV1(gz), /maximum uncompressed size/i);
});

test("v1 reader rejects packages where package.json does not match embedded manifest", () => {
  const bundle = {
    format: "formula-extension-package",
    formatVersion: 1,
    createdAt: "2020-01-01T00:00:00.000Z",
    manifest: { name: "a", publisher: "p", version: "1.0.0", main: "./dist/extension.js", engines: { formula: "^1.0.0" } },
    files: [
      {
        path: "package.json",
        dataBase64: Buffer.from(
          JSON.stringify({
            name: "b",
            publisher: "p",
            version: "1.0.0",
            main: "./dist/extension.js",
            engines: { formula: "^1.0.0" },
          }),
          "utf8",
        ).toString("base64"),
      },
    ],
  };

  const gz = zlib.gzipSync(Buffer.from(JSON.stringify(bundle), "utf8"), { level: 9 });
  assert.throws(() => readExtensionPackageV1(gz), /package\.json does not match/i);
});

test("v1 reader rejects case-insensitive duplicate paths", () => {
  const bundle = {
    format: "formula-extension-package",
    formatVersion: 1,
    createdAt: "2020-01-01T00:00:00.000Z",
    manifest: { name: "a", publisher: "p", version: "1.0.0", main: "./dist/extension.js", engines: { formula: "^1.0.0" } },
    files: [
      {
        path: "package.json",
        dataBase64: Buffer.from(
          JSON.stringify({
            name: "a",
            publisher: "p",
            version: "1.0.0",
            main: "./dist/extension.js",
            engines: { formula: "^1.0.0" },
          }),
          "utf8",
        ).toString("base64"),
      },
      { path: "README.md", dataBase64: Buffer.from("a").toString("base64") },
      { path: "readme.md", dataBase64: Buffer.from("b").toString("base64") },
    ],
  };

  const gz = zlib.gzipSync(Buffer.from(JSON.stringify(bundle), "utf8"), { level: 9 });
  assert.throws(() => readExtensionPackageV1(gz), /case-insensitive/i);
});
