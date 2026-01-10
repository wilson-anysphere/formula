import { EventEmitter } from "node:events";

export class AuditStore {
  constructor(options = {}) {
    this.maxEventsPerOrg = options.maxEventsPerOrg ?? 10_000;
    this.eventsByOrgId = new Map();
    this.emittersByOrgId = new Map();
  }

  getEmitter(orgId) {
    if (!this.emittersByOrgId.has(orgId)) this.emittersByOrgId.set(orgId, new EventEmitter());
    return this.emittersByOrgId.get(orgId);
  }

  append(orgId, event) {
    if (!this.eventsByOrgId.has(orgId)) this.eventsByOrgId.set(orgId, []);
    const events = this.eventsByOrgId.get(orgId);
    events.push(event);
    if (events.length > this.maxEventsPerOrg) events.splice(0, events.length - this.maxEventsPerOrg);
    this.getEmitter(orgId).emit("event", event);
  }

  list(orgId, limit = 100) {
    const events = this.eventsByOrgId.get(orgId) || [];
    return events.slice(Math.max(0, events.length - limit));
  }
}
