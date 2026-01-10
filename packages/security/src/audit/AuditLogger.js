import crypto from "node:crypto";

export class AuditLogger {
  /**
   * @param {object} options
   * @param {{ append: (event: any) => void }} options.store
   */
  constructor({ store }) {
    if (!store || typeof store.append !== "function") {
      throw new TypeError("AuditLogger requires a store with append()");
    }
    this.store = store;
  }

  /**
   * @param {object} input
   * @param {string} input.eventType
   * @param {{type: string, id: string}} input.actor
   * @param {boolean} input.success
   * @param {object} [input.metadata]
   * @returns {string} event id
   */
  log(input) {
    const event = {
      id: crypto.randomUUID(),
      ts: Date.now(),
      eventType: input.eventType,
      actor: input.actor,
      success: Boolean(input.success),
      metadata: input.metadata ?? {}
    };

    this.store.append(event);
    return event.id;
  }
}
