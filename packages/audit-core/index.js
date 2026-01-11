const crypto = require("node:crypto");
const auditEventJsonSchema = require("./auditEvent.schema.json");

const AUDIT_EVENT_SCHEMA_VERSION = 1;

function isPlainObject(value) {
  if (!value || typeof value !== "object") return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function isNonEmptyString(value) {
  return typeof value === "string" && value.length > 0;
}

function isOptionalString(value) {
  return value === undefined || value === null || typeof value === "string";
}

function isOptionalNonEmptyString(value) {
  return value === undefined || value === null || (typeof value === "string" && value.length > 0);
}

function isIsoDateTime(value) {
  if (!isNonEmptyString(value)) return false;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return false;
  // Ensure round-trippable ISO 8601 timestamps (what we emit everywhere).
  return date.toISOString() === value;
}

function validateAuditEvent(event) {
  const errors = [];
  if (!isPlainObject(event)) {
    return { valid: false, errors: ["AuditEvent must be an object"] };
  }

  const allowedTopLevelKeys = new Set([
    "schemaVersion",
    "id",
    "timestamp",
    "eventType",
    "actor",
    "context",
    "resource",
    "success",
    "error",
    "details",
    "correlation"
  ]);
  for (const key of Object.keys(event)) {
    if (!allowedTopLevelKeys.has(key)) errors.push(`Unknown top-level key: ${key}`);
  }

  if (event.schemaVersion !== AUDIT_EVENT_SCHEMA_VERSION) {
    errors.push(`schemaVersion must be ${AUDIT_EVENT_SCHEMA_VERSION}`);
  }

  if (!isNonEmptyString(event.id)) errors.push("id must be a non-empty string");
  if (!isIsoDateTime(event.timestamp)) errors.push("timestamp must be an ISO 8601 string (e.g. 2025-01-01T00:00:00.000Z)");
  if (!isNonEmptyString(event.eventType)) errors.push("eventType must be a non-empty string");

  if (!isPlainObject(event.actor)) {
    errors.push("actor must be an object");
  } else {
    const allowedActorKeys = new Set(["type", "id"]);
    for (const key of Object.keys(event.actor)) {
      if (!allowedActorKeys.has(key)) errors.push(`Unknown actor key: ${key}`);
    }
    if (!isNonEmptyString(event.actor.type)) errors.push("actor.type must be a non-empty string");
    if (!isNonEmptyString(event.actor.id)) errors.push("actor.id must be a non-empty string");
  }

  if (event.context !== undefined) {
    if (!isPlainObject(event.context)) {
      errors.push("context must be an object when provided");
    } else {
      const ctx = event.context;
      const allowedContextKeys = new Set(["orgId", "userId", "userEmail", "ipAddress", "userAgent", "sessionId"]);
      for (const key of Object.keys(ctx)) {
        if (!allowedContextKeys.has(key)) errors.push(`Unknown context key: ${key}`);
      }
      if (!isOptionalNonEmptyString(ctx.orgId)) errors.push("context.orgId must be a string or null");
      if (!isOptionalNonEmptyString(ctx.userId)) errors.push("context.userId must be a string or null");
      if (!isOptionalNonEmptyString(ctx.userEmail)) errors.push("context.userEmail must be a string or null");
      if (!isOptionalNonEmptyString(ctx.ipAddress)) errors.push("context.ipAddress must be a string or null");
      if (!isOptionalNonEmptyString(ctx.userAgent)) errors.push("context.userAgent must be a string or null");
      if (!isOptionalNonEmptyString(ctx.sessionId)) errors.push("context.sessionId must be a string or null");
    }
  }

  if (event.resource !== undefined) {
    if (!isPlainObject(event.resource)) {
      errors.push("resource must be an object when provided");
    } else {
      const allowedResourceKeys = new Set(["type", "id", "name"]);
      for (const key of Object.keys(event.resource)) {
        if (!allowedResourceKeys.has(key)) errors.push(`Unknown resource key: ${key}`);
      }
      if (!isNonEmptyString(event.resource.type)) errors.push("resource.type must be a non-empty string");
      if (!isOptionalNonEmptyString(event.resource.id)) errors.push("resource.id must be a string or null");
      if (!isOptionalNonEmptyString(event.resource.name)) errors.push("resource.name must be a string or null");
    }
  }

  if (typeof event.success !== "boolean") errors.push("success must be a boolean");

  if (event.error !== undefined) {
    if (!isPlainObject(event.error)) {
      errors.push("error must be an object when provided");
    } else {
      const allowedErrorKeys = new Set(["code", "message"]);
      for (const key of Object.keys(event.error)) {
        if (!allowedErrorKeys.has(key)) errors.push(`Unknown error key: ${key}`);
      }
      if (!isOptionalNonEmptyString(event.error.code)) errors.push("error.code must be a string or null");
      if (!isOptionalNonEmptyString(event.error.message)) errors.push("error.message must be a string or null");
    }
  }

  if (event.details !== undefined) {
    if (!isPlainObject(event.details)) errors.push("details must be an object when provided");
  }

  if (event.correlation !== undefined) {
    if (!isPlainObject(event.correlation)) {
      errors.push("correlation must be an object when provided");
    } else {
      const allowedCorrelationKeys = new Set(["requestId", "traceId"]);
      for (const key of Object.keys(event.correlation)) {
        if (!allowedCorrelationKeys.has(key)) errors.push(`Unknown correlation key: ${key}`);
      }
      if (!isOptionalNonEmptyString(event.correlation.requestId)) errors.push("correlation.requestId must be a string or null");
      if (!isOptionalNonEmptyString(event.correlation.traceId)) errors.push("correlation.traceId must be a string or null");
    }
  }

  if (Object.keys(event).some((key) => key === "ts" || key === "metadata")) {
    errors.push("Legacy fields are not allowed on canonical AuditEvent (use timestamp/details)");
  }

  return { valid: errors.length === 0, errors };
}

function assertAuditEvent(event) {
  const result = validateAuditEvent(event);
  if (result.valid) return;
  const err = new Error(`Invalid AuditEvent: ${result.errors.join("; ")}`);
  err.errors = result.errors;
  throw err;
}

function createAuditEvent(input) {
  if (!input || typeof input !== "object") {
    throw new TypeError("createAuditEvent requires an input object");
  }

  const id = isNonEmptyString(input.id) ? input.id : crypto.randomUUID();
  const timestamp = isNonEmptyString(input.timestamp) ? input.timestamp : new Date().toISOString();
  const schemaVersion = input.schemaVersion ?? AUDIT_EVENT_SCHEMA_VERSION;

  const event = {
    schemaVersion,
    id,
    timestamp,
    eventType: input.eventType,
    actor: input.actor,
    context: input.context,
    resource: input.resource,
    success: Boolean(input.success),
    error: input.error,
    details: input.details ?? {},
    correlation: input.correlation
  };

  assertAuditEvent(event);
  return event;
}

// --- Redaction helpers (shared across stores + SIEM exports)

const DEFAULT_REDACTION_TEXT = "[REDACTED]";

const DEFAULT_SENSITIVE_KEY_PATTERNS = [
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

function redactValue(value, options) {
  const { redactionText = DEFAULT_REDACTION_TEXT, sensitiveKeyPatterns = DEFAULT_SENSITIVE_KEY_PATTERNS } = options || {};

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

function redactAuditEvent(event, options) {
  return redactValue(event, options);
}

// --- SIEM serialization (JSON/CEF/LEEF)

function escapeCefHeaderField(value) {
  return String(value).replace(/\\/g, "\\\\").replace(/\|/g, "\\|").replace(/\n/g, "\\n").replace(/\r/g, "\\r");
}

function escapeCefExtensionValue(value) {
  return String(value)
    .replace(/\\/g, "\\\\")
    .replace(/=/g, "\\=")
    .replace(/\n/g, "\\n")
    .replace(/\r/g, "\\r")
    .replace(/\t/g, "\\t");
}

function escapeLeefHeaderField(value) {
  return String(value).replace(/\\/g, "\\\\").replace(/\|/g, "\\|").replace(/\n/g, "\\n").replace(/\r/g, "\\r");
}

function escapeLeefValue(value, delimiter) {
  const delimiterText = delimiter === "\t" ? "\\t" : delimiter;

  return String(value)
    .replace(/\\/g, "\\\\")
    .replace(/=/g, "\\=")
    .replace(/\n/g, "\\n")
    .replace(/\r/g, "\\r")
    .replace(/\t/g, "\\t")
    .replace(new RegExp(delimiter.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g"), delimiterText);
}

function toTimestampMs(value) {
  if (!value) return undefined;
  if (typeof value === "number") return value;
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return undefined;
  return date.getTime();
}

function defaultSeverity(event) {
  if (event && event.success === false) return 8;
  if (event && typeof event.eventType === "string" && /failed|denied|blocked/i.test(event.eventType)) return 8;
  if (event && typeof event.eventType === "string" && /^admin\./.test(event.eventType)) return 6;
  return 5;
}

function pickUserId(event) {
  return event?.context?.userId ?? (event?.actor?.type === "user" ? event.actor.id : undefined);
}

function buildCefExtension(event) {
  const timestampMs = toTimestampMs(event.timestamp) ?? Date.now();
  const ctx = event.context || {};
  const res = event.resource || {};
  const error = event.error || {};
  const correlation = event.correlation || {};

  const pairs = [];
  pairs.push(["rt", timestampMs]);
  if (event.id) pairs.push(["externalId", event.id]);
  if (ctx.orgId) {
    pairs.push(["cs1Label", "orgId"]);
    pairs.push(["cs1", ctx.orgId]);
  }
  if (ctx.sessionId) {
    pairs.push(["cs2Label", "sessionId"]);
    pairs.push(["cs2", ctx.sessionId]);
  }

  const userId = pickUserId(event);
  if (userId) pairs.push(["suid", userId]);
  if (ctx.userEmail) pairs.push(["suser", ctx.userEmail]);
  if (ctx.ipAddress) pairs.push(["src", ctx.ipAddress]);
  if (ctx.userAgent) pairs.push(["requestClientApplication", ctx.userAgent]);
  if (res.type) {
    pairs.push(["cs3Label", "resourceType"]);
    pairs.push(["cs3", res.type]);
  }
  if (res.id) {
    pairs.push(["cs4Label", "resourceId"]);
    pairs.push(["cs4", res.id]);
  }
  if (res.name) {
    pairs.push(["cs5Label", "resourceName"]);
    pairs.push(["cs5", res.name]);
  }
  if (typeof event.success === "boolean") pairs.push(["outcome", event.success ? "success" : "failure"]);
  if (error.code) pairs.push(["reason", error.code]);
  if (error.message) pairs.push(["msg", error.message]);

  if (correlation.requestId) {
    pairs.push(["cs7Label", "requestId"]);
    pairs.push(["cs7", correlation.requestId]);
  }
  if (correlation.traceId) {
    pairs.push(["cs8Label", "traceId"]);
    pairs.push(["cs8", correlation.traceId]);
  }

  if (event.details) {
    pairs.push(["cs6Label", "details"]);
    pairs.push(["cs6", JSON.stringify(event.details)]);
  }

  return pairs.map(([key, value]) => `${key}=${escapeCefExtensionValue(value)}`).join(" ");
}

function toCef(event, options = {}) {
  const safeEvent = options.redact === false ? event : redactAuditEvent(event, options.redactionOptions);
  const vendor = options.vendor ?? "Formula";
  const product = options.product ?? "Spreadsheet";
  const deviceVersion = options.deviceVersion ?? "1.0";

  const signature = safeEvent.eventType ?? "audit";
  const name = safeEvent.eventType ?? signature;
  const severity = options.severity ?? defaultSeverity(safeEvent);

  const header = [
    "CEF:0",
    escapeCefHeaderField(vendor),
    escapeCefHeaderField(product),
    escapeCefHeaderField(deviceVersion),
    escapeCefHeaderField(signature),
    escapeCefHeaderField(name),
    String(severity)
  ].join("|");

  return `${header}|${buildCefExtension(safeEvent)}`;
}

function toLeef(event, options = {}) {
  const safeEvent = options.redact === false ? event : redactAuditEvent(event, options.redactionOptions);
  const vendor = options.vendor ?? "Formula";
  const product = options.product ?? "Spreadsheet";
  const productVersion = options.productVersion ?? "1.0";
  const delimiter = options.delimiter ?? "\t";

  const eventId = safeEvent.eventType ?? "audit";
  const header = `LEEF:2.0|${escapeLeefHeaderField(vendor)}|${escapeLeefHeaderField(product)}|${escapeLeefHeaderField(
    productVersion
  )}|${escapeLeefHeaderField(eventId)}|${delimiter}`;

  const timestampMs = toTimestampMs(safeEvent.timestamp);
  const timestampIso = timestampMs ? new Date(timestampMs).toISOString() : undefined;

  const ctx = safeEvent.context || {};
  const res = safeEvent.resource || {};
  const error = safeEvent.error || {};
  const correlation = safeEvent.correlation || {};

  const attributes = {
    ts: timestampIso,
    id: safeEvent.id,
    orgId: ctx.orgId,
    eventType: safeEvent.eventType,
    actorType: safeEvent.actor?.type,
    actorId: safeEvent.actor?.id,
    userId: pickUserId(safeEvent),
    userEmail: ctx.userEmail,
    ipAddress: ctx.ipAddress,
    userAgent: ctx.userAgent,
    sessionId: ctx.sessionId,
    resourceType: res.type,
    resourceId: res.id,
    resourceName: res.name,
    success: typeof safeEvent.success === "boolean" ? String(safeEvent.success) : undefined,
    errorCode: error.code,
    errorMessage: error.message,
    requestId: correlation.requestId,
    traceId: correlation.traceId,
    details: safeEvent.details ? JSON.stringify(safeEvent.details) : undefined
  };

  const segments = Object.entries(attributes)
    .filter(([, value]) => value !== undefined && value !== null)
    .map(([key, value]) => `${key}=${escapeLeefValue(value, delimiter)}`);

  return header + segments.join(delimiter);
}

function serializeBatch(events, options = {}) {
  const format = options.format ?? "json";
  const redactionOptions = options.redactionOptions;
  const safeEvents = options.redact === false ? events : events.map((event) => redactAuditEvent(event, redactionOptions));

  if (format === "cef") {
    const lines = safeEvents.map((event) => toCef(event, { ...options, redact: false }));
    return { contentType: "text/plain", body: Buffer.from(lines.join("\n") + "\n", "utf8") };
  }

  if (format === "leef") {
    const lines = safeEvents.map((event) => toLeef(event, { ...options, redact: false }));
    return { contentType: "text/plain", body: Buffer.from(lines.join("\n") + "\n", "utf8") };
  }

  return { contentType: "application/json", body: Buffer.from(JSON.stringify(safeEvents), "utf8") };
}

// --- Storage adapters

function auditEventToSqliteRow(event) {
  assertAuditEvent(event);

  const ts = toTimestampMs(event.timestamp);
  if (!ts) throw new Error("AuditEvent.timestamp is invalid");

  const ctx = event.context || {};
  const res = event.resource || {};
  const error = event.error || {};
  const correlation = event.correlation || {};

  return {
    id: event.id,
    ts,
    timestamp: event.timestamp,
    eventType: event.eventType,
    actorType: event.actor.type,
    actorId: event.actor.id,
    orgId: ctx.orgId ?? null,
    userId: ctx.userId ?? null,
    userEmail: ctx.userEmail ?? null,
    ipAddress: ctx.ipAddress ?? null,
    userAgent: ctx.userAgent ?? null,
    sessionId: ctx.sessionId ?? null,
    resourceType: res.type ?? null,
    resourceId: res.id ?? null,
    resourceName: res.name ?? null,
    success: event.success ? 1 : 0,
    errorCode: error.code ?? null,
    errorMessage: error.message ?? null,
    details: JSON.stringify(event.details ?? {}),
    requestId: correlation.requestId ?? null,
    traceId: correlation.traceId ?? null
  };
}

function buildPostgresAuditLogInsert(event) {
  assertAuditEvent(event);

  const ctx = event.context || {};
  const res = event.resource || {};
  const error = event.error || {};

  // NOTE: audit_log.resource_type is NOT NULL in the schema. If callers omit resource,
  // we still insert a placeholder so audits never fail the user request.
  const resourceType = res.type || "unknown";

  return {
    text: `
      INSERT INTO audit_log (
        id, org_id, user_id, user_email, event_type, resource_type, resource_id,
        ip_address, user_agent, session_id,
        success, error_code, error_message, details, created_at
      )
      VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14::jsonb,$15)
    `,
    values: [
      event.id,
      ctx.orgId ?? null,
      ctx.userId ?? null,
      ctx.userEmail ?? null,
      event.eventType,
      resourceType,
      res.id ?? null,
      ctx.ipAddress ?? null,
      ctx.userAgent ?? null,
      ctx.sessionId ?? null,
      event.success,
      error.code ?? null,
      error.message ?? null,
      JSON.stringify(event.details ?? {}),
      event.timestamp
    ]
  };
}

function retentionCutoffMs(now, retentionDays) {
  if (!Number.isFinite(now)) throw new TypeError("now must be a unix timestamp in ms");
  if (!Number.isFinite(retentionDays) || retentionDays < 0) throw new TypeError("retentionDays must be a non-negative number");
  return now - retentionDays * 24 * 60 * 60 * 1000;
}

module.exports = {
  AUDIT_EVENT_SCHEMA_VERSION,
  auditEventJsonSchema,
  createAuditEvent,
  validateAuditEvent,
  assertAuditEvent,
  DEFAULT_REDACTION_TEXT,
  DEFAULT_SENSITIVE_KEY_PATTERNS,
  redactValue,
  redactAuditEvent,
  escapeCefHeaderField,
  escapeCefExtensionValue,
  toCef,
  toLeef,
  serializeBatch,
  auditEventToSqliteRow,
  buildPostgresAuditLogInsert,
  retentionCutoffMs
};
