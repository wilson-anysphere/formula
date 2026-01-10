export const DEFAULT_REDACTION_TEXT = "[REDACTED]";

export const DEFAULT_SENSITIVE_KEY_PATTERNS = [
  /pass(word)?/i,
  /secret/i,
  /token/i,
  /api[-_]?key/i,
  /authorization/i,
  /cookie/i,
  /set[-_]?cookie/i,
  /private[-_]?key/i,
  /client[-_]?secret/i,
  /refresh[-_]?token/i,
  /access[-_]?token/i
];

function isPlainObject(value) {
  if (!value || typeof value !== "object") return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function looksLikeJwt(value) {
  if (typeof value !== "string") return false;
  const trimmed = value.trim();
  if (trimmed.length < 40) return false;
  const parts = trimmed.split(".");
  if (parts.length !== 3) return false;
  return parts.every((part) => /^[A-Za-z0-9_-]+={0,2}$/.test(part));
}

function redactString(value, redactionText = DEFAULT_REDACTION_TEXT) {
  if (typeof value !== "string") return value;

  const trimmed = value.trim();
  if (/^Bearer\s+/i.test(trimmed)) return `Bearer ${redactionText}`;
  if (/^Splunk\s+/i.test(trimmed)) return `Splunk ${redactionText}`;
  if (looksLikeJwt(trimmed)) return redactionText;

  return value;
}

function shouldRedactKey(key, patterns) {
  if (typeof key !== "string") return false;
  return patterns.some((pattern) => pattern.test(key));
}

export function redactValue(value, options) {
  const {
    redactionText = DEFAULT_REDACTION_TEXT,
    sensitiveKeyPatterns = DEFAULT_SENSITIVE_KEY_PATTERNS
  } = options || {};

  if (value === null || value === undefined) return value;

  if (typeof value === "string") return redactString(value, redactionText);

  if (Array.isArray(value)) return value.map((item) => redactValue(item, options));

  if (value instanceof Date) return new Date(value.getTime());

  if (!isPlainObject(value)) return value;

  const output = {};
  for (const [key, nestedValue] of Object.entries(value)) {
    if (shouldRedactKey(key, sensitiveKeyPatterns)) {
      output[key] = redactionText;
      continue;
    }

    output[key] = redactValue(nestedValue, options);
  }

  return output;
}

export function redactAuditEvent(event, options) {
  return redactValue(event, options);
}
