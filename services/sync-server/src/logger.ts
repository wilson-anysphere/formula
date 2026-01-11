import pino, { type DestinationStream, type LoggerOptions } from "pino";

// IMPORTANT: Avoid logging raw request headers or entire config objects.
// Redaction is defense-in-depth, but it's easy to miss secrets if we log large,
// unstructured objects.
const REDACT_PATHS = [
  // Standard auth.
  "req.headers.authorization",

  // Internal/admin auth (Task 168/172).
  "req.headers.x-internal-admin-token",
  'req.headers["x-internal-admin-token"]',
  "req.headers.x-sync-server-admin-token",
  'req.headers["x-sync-server-admin-token"]',

  // Encryption-at-rest (Task 169) — redact keys if a config object ever gets logged.
  "persistence.encryption.keyBase64",
  "config.persistence.encryption.keyBase64",
  "persistence.encryption.keyRing",
  "config.persistence.encryption.keyRing",
  "persistence.leveldbEncryption.key",
  "config.persistence.leveldbEncryption.key",

  // Common token fields (legacy / defensive).
  "token",
  "authToken",
] as const;

export function createLogger(level: string, destination?: DestinationStream) {
  const options: LoggerOptions = {
    level,
    base: {
      service: "sync-server",
    },
    redact: {
      paths: [...REDACT_PATHS],
      remove: true,
    },
  };

  return destination ? pino(options, destination) : pino(options);
}
