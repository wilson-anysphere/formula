/**
 * Minimal audit logger for DLP/AI decisions.
 *
 * In production this would integrate with the enterprise audit log pipeline. Here we
 * keep the surface area small and deterministic for unit testing.
 */

import { createAuditEvent } from "../../../audit-core/index.js";

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
