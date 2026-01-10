const crypto = require("node:crypto");

function randomId(bytes = 16) {
  return crypto.randomBytes(bytes).toString("hex");
}

class InMemoryAuditLogger {
  constructor() {
    this.events = [];
  }

  async log(event) {
    const full = {
      id: randomId(16),
      timestamp: new Date().toISOString(),
      ...event
    };
    this.events.push(full);
    return full;
  }
}

module.exports = { InMemoryAuditLogger };
