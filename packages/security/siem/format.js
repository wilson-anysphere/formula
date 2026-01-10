import { redactAuditEvent } from "./redaction.js";

export function escapeCefHeaderField(value) {
  return String(value).replace(/\\/g, "\\\\").replace(/\|/g, "\\|").replace(/\n/g, "\\n").replace(/\r/g, "\\r");
}

export function escapeCefExtensionValue(value) {
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

function buildCefExtension(event) {
  const timestampMs = toTimestampMs(event.timestamp) ?? Date.now();

  const pairs = [];
  pairs.push(["rt", timestampMs]);
  if (event.id) pairs.push(["externalId", event.id]);
  if (event.orgId) {
    pairs.push(["cs1Label", "orgId"]);
    pairs.push(["cs1", event.orgId]);
  }
  if (event.sessionId) {
    pairs.push(["cs2Label", "sessionId"]);
    pairs.push(["cs2", event.sessionId]);
  }
  if (event.userId) pairs.push(["suid", event.userId]);
  if (event.userEmail) pairs.push(["suser", event.userEmail]);
  if (event.ipAddress) pairs.push(["src", event.ipAddress]);
  if (event.userAgent) pairs.push(["requestClientApplication", event.userAgent]);
  if (event.resourceType) {
    pairs.push(["cs3Label", "resourceType"]);
    pairs.push(["cs3", event.resourceType]);
  }
  if (event.resourceId) {
    pairs.push(["cs4Label", "resourceId"]);
    pairs.push(["cs4", event.resourceId]);
  }
  if (event.resourceName) {
    pairs.push(["cs5Label", "resourceName"]);
    pairs.push(["cs5", event.resourceName]);
  }
  if (typeof event.success === "boolean") pairs.push(["outcome", event.success ? "success" : "failure"]);
  if (event.errorCode) pairs.push(["reason", event.errorCode]);
  if (event.errorMessage) pairs.push(["msg", event.errorMessage]);
  if (event.details) {
    pairs.push(["cs6Label", "details"]);
    pairs.push(["cs6", JSON.stringify(event.details)]);
  }

  return pairs.map(([key, value]) => `${key}=${escapeCefExtensionValue(value)}`).join(" ");
}

export function toCef(event, options = {}) {
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

export function toLeef(event, options = {}) {
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

  const attributes = {
    ts: timestampIso,
    id: safeEvent.id,
    orgId: safeEvent.orgId,
    eventType: safeEvent.eventType,
    userId: safeEvent.userId,
    userEmail: safeEvent.userEmail,
    ipAddress: safeEvent.ipAddress,
    userAgent: safeEvent.userAgent,
    sessionId: safeEvent.sessionId,
    resourceType: safeEvent.resourceType,
    resourceId: safeEvent.resourceId,
    resourceName: safeEvent.resourceName,
    success: typeof safeEvent.success === "boolean" ? String(safeEvent.success) : undefined,
    errorCode: safeEvent.errorCode,
    errorMessage: safeEvent.errorMessage,
    details: safeEvent.details ? JSON.stringify(safeEvent.details) : undefined
  };

  const segments = Object.entries(attributes)
    .filter(([, value]) => value !== undefined && value !== null)
    .map(([key, value]) => `${key}=${escapeLeefValue(value, delimiter)}`);

  return header + segments.join(delimiter);
}

export function serializeBatch(events, options = {}) {
  const format = options.format ?? "json";
  const redactionOptions = options.redactionOptions;

  const safeEvents = options.redact === false ? events : events.map((event) => redactAuditEvent(event, redactionOptions));

  if (format === "cef") {
    const lines = safeEvents.map((event) => toCef(event, { ...options, redact: false }));
    return {
      contentType: "text/plain",
      body: Buffer.from(lines.join("\n") + "\n", "utf8")
    };
  }

  if (format === "leef") {
    const lines = safeEvents.map((event) => toLeef(event, { ...options, redact: false }));
    return {
      contentType: "text/plain",
      body: Buffer.from(lines.join("\n") + "\n", "utf8")
    };
  }

  return {
    contentType: "application/json",
    body: Buffer.from(JSON.stringify(safeEvents), "utf8")
  };
}
