import { createAuditEvent } from "../../../audit-core/index.js";

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
   * @param {object} [input.context]
   * @param {object} [input.resource]
   * @param {object} [input.error]
   * @param {object} [input.details]
   * @param {object} [input.metadata] - legacy alias for details
   * @param {object} [input.correlation]
   * @returns {string} event id
   */
  log(input) {
    const event = createAuditEvent({
      eventType: input.eventType,
      actor: input.actor,
      success: Boolean(input.success),
      context: input.context,
      resource: input.resource,
      error: input.error,
      details: input.details ?? input.metadata ?? {},
      correlation: input.correlation
    });

    this.store.append(event);
    return event.id;
  }
}
