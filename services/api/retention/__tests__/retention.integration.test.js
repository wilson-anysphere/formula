const test = require("node:test");
const assert = require("node:assert/strict");

const { InMemoryAuditLogger } = require("../../audit/auditLogger.js");
const { InMemoryOrgPolicyStore, OrgPolicyService } = require("../../org/orgPolicyService.js");
const {
  InMemoryAuditLogArchiveStore,
  InMemoryAuditLogStore,
  InMemoryDeletedDocumentStore,
  InMemoryLegalHoldStore,
  InMemoryVersionStore,
  RetentionService
} = require("../retentionService.js");

function daysBefore(now, days) {
  return new Date(now.getTime() - days * 24 * 60 * 60 * 1000).toISOString();
}

test("retention job deletes/archives expected records (legal hold respected)", async () => {
  const now = new Date("2026-01-10T00:00:00.000Z");

  const auditLogger = new InMemoryAuditLogger();
  const orgPolicyStore = new InMemoryOrgPolicyStore();
  const orgPolicyService = new OrgPolicyService({ store: orgPolicyStore, auditLogger });

  const orgId = "org-1";

  await orgPolicyService.update(
    orgId,
    {
      retention: {
        versionRetentionDays: 10,
        auditLogRetentionDays: 365,
        deletedDocumentRetentionDays: 30,
        legalHoldOverridesRetention: true
      }
    },
    { actor: { userId: "admin-1" } }
  );

  const versionStore = new InMemoryVersionStore();
  await versionStore.add({
    id: "v1",
    orgId,
    docId: "doc-1",
    createdAt: daysBefore(now, 20)
  });
  await versionStore.add({
    id: "v2",
    orgId,
    docId: "doc-1",
    createdAt: daysBefore(now, 5)
  });
  await versionStore.add({
    id: "v3",
    orgId,
    docId: "doc-2",
    createdAt: daysBefore(now, 20)
  });

  const auditLogStore = new InMemoryAuditLogStore();
  await auditLogStore.add({
    id: "a1",
    orgId,
    timestamp: daysBefore(now, 400),
    eventType: "document.opened"
  });
  await auditLogStore.add({
    id: "a2",
    orgId,
    timestamp: daysBefore(now, 1),
    eventType: "document.opened"
  });

  const auditLogArchiveStore = new InMemoryAuditLogArchiveStore();

  const deletedDocumentStore = new InMemoryDeletedDocumentStore();
  await deletedDocumentStore.add({
    orgId,
    docId: "doc-1",
    deletedAt: daysBefore(now, 40)
  });
  await deletedDocumentStore.add({
    orgId,
    docId: "doc-2",
    deletedAt: daysBefore(now, 40)
  });

  const legalHoldStore = new InMemoryLegalHoldStore();
  await legalHoldStore.addHold({ orgId, docId: "doc-2" });

  const retentionService = new RetentionService({
    orgPolicyService,
    versionStore,
    auditLogStore,
    auditLogArchiveStore,
    deletedDocumentStore,
    legalHoldStore
  });

  const report = await retentionService.apply(orgId, { now });

  assert.deepEqual(report, {
    versionsDeleted: 1,
    auditLogsArchived: 1,
    deletedDocumentsPurged: 1
  });

  assert.deepEqual(
    (await versionStore.listByOrg(orgId)).map((v) => v.id).sort(),
    ["v2", "v3"],
    "expected old version under legal hold to remain"
  );

  assert.deepEqual(
    (await auditLogStore.listByOrg(orgId)).map((e) => e.id).sort(),
    ["a2"]
  );
  assert.deepEqual(
    auditLogArchiveStore.archivedEvents.map((e) => e.id).sort(),
    ["a1"]
  );

  assert.deepEqual(
    (await deletedDocumentStore.listByOrg(orgId)).map((d) => d.docId).sort(),
    ["doc-2"],
    "expected deleted doc under legal hold to remain"
  );

  assert.ok(
    auditLogger.events.some((e) => e.eventType === "org.policy.retention.updated"),
    "expected retention policy change to be audited"
  );
});
