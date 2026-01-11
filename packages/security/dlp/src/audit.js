/**
 * Minimal audit logger for DLP/AI decisions.
 *
 * NOTE: This module must run in both Node (unit tests) and browser-like runtimes
 * (desktop/webview + Playwright). Keep it dependency-free and avoid importing
 * Node-only CommonJS helpers.
 */
const AUDIT_EVENT_SCHEMA_VERSION = 1;

let auditEventSeq = 0;

function createAuditEvent(input) {
  auditEventSeq += 1;
  const eventInput = input && typeof input === "object" ? input : {};
  const id =
    typeof eventInput.id === "string" && eventInput.id.length > 0
      ? eventInput.id
      : typeof globalThis.crypto?.randomUUID === "function"
        ? globalThis.crypto.randomUUID()
        : `audit_${Date.now()}_${auditEventSeq}`;
  const timestamp =
    typeof eventInput.timestamp === "string" && eventInput.timestamp.length > 0
      ? eventInput.timestamp
      : new Date().toISOString();

  // Keep the event shape compatible with `@formula/audit-core` without depending
  // on Node-only modules (the desktop UI runs in a browser runtime for e2e).
  return {
    schemaVersion: eventInput.schemaVersion ?? AUDIT_EVENT_SCHEMA_VERSION,
    id,
    timestamp,
    eventType: eventInput.eventType,
    actor: eventInput.actor,
    context: eventInput.context,
    resource: eventInput.resource,
    success: Boolean(eventInput.success),
    error: eventInput.error,
    details: eventInput.details ?? {},
    correlation: eventInput.correlation,
  };
}

export class InMemoryAuditLogger {
  constructor() {
    this.events = [];
  }

  /**
   * @param {any} event
   */
  log(event) {
    const input = event && typeof event === "object" ? event : {};

    const actor =
      input.actor && typeof input.actor === "object" && input.actor.type && input.actor.id
        ? input.actor
        : { type: "system", id: "dlp" };

    const decision = input.decision?.decision;
    const normalizedDecision = typeof decision === "string" ? decision.toLowerCase() : null;
    const success =
      typeof input.success === "boolean" ? input.success : normalizedDecision ? normalizedDecision !== "block" : true;

    const resource =
      typeof input.documentId === "string" && input.documentId.length > 0
        ? { type: "document", id: input.documentId }
        : undefined;

    const eventType =
      typeof input.eventType === "string" && input.eventType.length > 0
        ? input.eventType
        : typeof input.type === "string" && input.type.length > 0
          ? `dlp.${input.type}`
          : "dlp.event";

    const canonical = createAuditEvent({
      eventType,
      actor,
      success,
      resource,
      details: input
    });

    this.events.push(canonical);
    return canonical.id;
  }

  list() {
    return [...this.events];
  }
}
