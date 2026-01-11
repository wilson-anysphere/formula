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
