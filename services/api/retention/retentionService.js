function daysAgo(date, days) {
  return new Date(date.getTime() - days * 24 * 60 * 60 * 1000);
}

class InMemoryVersionStore {
  constructor() {
    this.versions = [];
  }

  async add(version) {
    this.versions.push(version);
  }

  async listByOrg(orgId) {
    return this.versions.filter((v) => v.orgId === orgId);
  }

  async deleteVersion(versionId) {
    this.versions = this.versions.filter((v) => v.id !== versionId);
  }
}

class InMemoryAuditLogStore {
  constructor() {
    this.events = [];
  }

  async add(event) {
    this.events.push(event);
  }

  async listByOrg(orgId) {
    return this.events.filter((e) => e.orgId === orgId);
  }

  async deleteEvent(eventId) {
    this.events = this.events.filter((e) => e.id !== eventId);
  }
}

class InMemoryAuditLogArchiveStore {
  constructor() {
    this.archivedEvents = [];
  }

  async add(event) {
    this.archivedEvents.push(event);
  }
}

class InMemoryDeletedDocumentStore {
  constructor() {
    this.deletedDocuments = [];
  }

  async add(record) {
    this.deletedDocuments.push(record);
  }

  async listByOrg(orgId) {
    return this.deletedDocuments.filter((d) => d.orgId === orgId);
  }

  async purge(docId) {
    this.deletedDocuments = this.deletedDocuments.filter((d) => d.docId !== docId);
  }
}

class InMemoryLegalHoldStore {
  constructor() {
    this.holds = new Map();
  }

  _key(orgId, docId) {
    return `${orgId}:${docId}`;
  }

  async addHold({ orgId, docId }) {
    this.holds.set(this._key(orgId, docId), true);
  }

  async hasHold({ orgId, docId }) {
    return this.holds.get(this._key(orgId, docId)) === true;
  }
}

class RetentionService {
  constructor({
    orgPolicyService,
    versionStore,
    auditLogStore,
    auditLogArchiveStore,
    deletedDocumentStore,
    legalHoldStore
  }) {
    this.orgPolicyService = orgPolicyService;
    this.versionStore = versionStore;
    this.auditLogStore = auditLogStore;
    this.auditLogArchiveStore = auditLogArchiveStore;
    this.deletedDocumentStore = deletedDocumentStore;
    this.legalHoldStore = legalHoldStore;
  }

  async apply(orgId, { now = new Date() } = {}) {
    const policy = await this.orgPolicyService.get(orgId);
    const retention = policy.retention;
    const report = {
      versionsDeleted: 0,
      auditLogsArchived: 0,
      deletedDocumentsPurged: 0
    };

    const versionCutoff = daysAgo(now, retention.versionRetentionDays);
    const auditLogCutoff = daysAgo(now, retention.auditLogRetentionDays);
    const deletedDocCutoff = daysAgo(now, retention.deletedDocumentRetentionDays);

    const legalHoldOverrides = retention.legalHoldOverridesRetention === true;

    const versions = await this.versionStore.listByOrg(orgId);
    for (const version of versions) {
      const createdAt = new Date(version.createdAt);
      if (createdAt > versionCutoff) continue;

      const held =
        legalHoldOverrides && (await this.legalHoldStore.hasHold({ orgId, docId: version.docId }));
      if (held) continue;

      await this.versionStore.deleteVersion(version.id);
      report.versionsDeleted++;
    }

    const events = await this.auditLogStore.listByOrg(orgId);
    for (const event of events) {
      const timestamp = new Date(event.timestamp);
      if (timestamp > auditLogCutoff) continue;
      await this.auditLogArchiveStore.add(event);
      await this.auditLogStore.deleteEvent(event.id);
      report.auditLogsArchived++;
    }

    const deletedDocs = await this.deletedDocumentStore.listByOrg(orgId);
    for (const record of deletedDocs) {
      const deletedAt = new Date(record.deletedAt);
      if (deletedAt > deletedDocCutoff) continue;

      const held =
        legalHoldOverrides && (await this.legalHoldStore.hasHold({ orgId, docId: record.docId }));
      if (held) continue;

      await this.deletedDocumentStore.purge(record.docId);
      report.deletedDocumentsPurged++;
    }

    return report;
  }
}

module.exports = {
  InMemoryAuditLogArchiveStore,
  InMemoryAuditLogStore,
  InMemoryDeletedDocumentStore,
  InMemoryLegalHoldStore,
  InMemoryVersionStore,
  RetentionService
};
