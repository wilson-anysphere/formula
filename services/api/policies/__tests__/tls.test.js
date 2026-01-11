const test = require("node:test");
const assert = require("node:assert/strict");

const {
  createPinnedCheckServerIdentity,
  createTlsServerOptions,
  sha256FingerprintHexFromCertRaw
} = require("../tls.js");

test("TLS server options enforce TLSv1.3 minimum", () => {
  const options = createTlsServerOptions({ foo: "bar" });
  assert.equal(options.minVersion, "TLSv1.3");
  assert.equal(options.foo, "bar");
});

test("certificate pinning accepts matching fingerprint and rejects mismatch", () => {
  const raw = Buffer.from("fake-cert-bytes");
  const pin = sha256FingerprintHexFromCertRaw(raw);

  const check = createPinnedCheckServerIdentity({ pins: [pin] });
  const ok = check("example.com", { raw, subjectaltname: "DNS:example.com" });
  assert.equal(ok, undefined);

  const badCheck = createPinnedCheckServerIdentity({ pins: ["deadbeef"] });
  const err = badCheck("example.com", { raw, subjectaltname: "DNS:example.com" });
  assert.ok(err instanceof Error);
  assert.equal(err.retriable, false);
});
