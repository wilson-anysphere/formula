import crypto from "node:crypto";

export function assertBufferLength(buf, expected, name) {
  if (!Buffer.isBuffer(buf)) {
    throw new TypeError(`${name} must be a Buffer`);
  }
  if (buf.length !== expected) {
    throw new RangeError(`${name} must be ${expected} bytes (got ${buf.length})`);
  }
}

export function toBase64(buf) {
  if (!Buffer.isBuffer(buf)) {
    throw new TypeError("toBase64 expects a Buffer");
  }
  return buf.toString("base64");
}

export function fromBase64(value, name = "base64") {
  if (typeof value !== "string") {
    throw new TypeError(`${name} must be a base64 string`);
  }
  return Buffer.from(value, "base64");
}

function isPlainObject(value) {
  return value !== null && typeof value === "object" && value.constructor === Object;
}

function sortJson(value) {
  if (Array.isArray(value)) {
    return value.map(sortJson);
  }
  if (isPlainObject(value)) {
    const sorted = {};
    for (const key of Object.keys(value).sort()) {
      sorted[key] = sortJson(value[key]);
    }
    return sorted;
  }
  return value;
}

/**
 * Deterministic JSON encoding suitable for use as AAD / encryption context.
 * Do NOT use this for security-sensitive canonicalization of untrusted input; it
 * exists so encryption context bytes are stable across runtime instances.
 */
export function canonicalJson(value) {
  return JSON.stringify(sortJson(value));
}

export function aadFromContext(context) {
  if (context === undefined || context === null) return null;
  const json = canonicalJson(context);
  return Buffer.from(json, "utf8");
}

export function randomId(bytes = 16) {
  return crypto.randomBytes(bytes).toString("hex");
}

