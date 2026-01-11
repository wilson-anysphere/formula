const crypto = require("node:crypto");
const tls = require("node:tls");

function createTlsServerOptions(baseOptions = {}) {
  return {
    ...baseOptions,
    minVersion: "TLSv1.3"
  };
}

function normalizeFingerprintHex(value) {
  return value.replaceAll(":", "").toLowerCase();
}

function sha256FingerprintHexFromCertRaw(raw) {
  return crypto.createHash("sha256").update(raw).digest("hex");
}

function createPinnedCheckServerIdentity({ pins }) {
  if (!Array.isArray(pins) || pins.length === 0) {
    throw new TypeError("pins must be a non-empty array");
  }
  const normalizedPins = new Set(
    pins.map((pin) => {
      if (typeof pin !== "string" || pin.length === 0) {
        throw new TypeError("pin must be a non-empty string");
      }
      return normalizeFingerprintHex(pin);
    })
  );

  return function checkServerIdentity(hostname, cert) {
    const defaultError = tls.checkServerIdentity(hostname, cert);
    if (defaultError) {
      defaultError.retriable = false;
      return defaultError;
    }

    const fingerprint = cert?.raw
      ? sha256FingerprintHexFromCertRaw(cert.raw)
      : typeof cert?.fingerprint256 === "string"
        ? normalizeFingerprintHex(cert.fingerprint256)
        : null;

    if (!fingerprint) {
      const err = new Error("Certificate pinning failed: certificate fingerprint not available");
      err.retriable = false;
      return err;
    }

    if (!normalizedPins.has(normalizeFingerprintHex(fingerprint))) {
      const err = new Error("Certificate pinning failed: server certificate fingerprint mismatch");
      err.retriable = false;
      return err;
    }

    return undefined;
  };
}

module.exports = { createTlsServerOptions, createPinnedCheckServerIdentity, sha256FingerprintHexFromCertRaw };
