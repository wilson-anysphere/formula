/**
 * Minimal audit logger for DLP/AI decisions.
 *
 * In production this would integrate with the enterprise audit log pipeline. Here we
 * keep the surface area small and deterministic for unit testing.
 */

export class InMemoryAuditLogger {
  constructor() {
    this.events = [];
  }

  /**
   * @param {any} event
   */
  log(event) {
    this.events.push({
      id: crypto.randomUUID(),
      timestamp: new Date().toISOString(),
      ...event,
    });
  }

  list() {
    return [...this.events];
  }
}

