const test = require("node:test");
const assert = require("node:assert/strict");

const { InMemoryOrgPolicyStore, OrgPolicyService } = require("../../org/orgPolicyService.js");
const { RegionalObjectStore } = require("../regionalObjectStore.js");
const { DocumentStorageRouter } = require("../documentStorageRouter.js");

test("cloud: documents are encrypted at rest and routed by residency policy", async () => {
  const orgId = "org-1";
  const store = new RegionalObjectStore();
  const { LocalKmsProvider } = await import(
    "../../../../packages/security/crypto/kms/localKmsProvider.js"
  );
  const kms = new LocalKmsProvider();

  const orgPolicyService = new OrgPolicyService({ store: new InMemoryOrgPolicyStore() });
  await orgPolicyService.update(orgId, { dataResidency: { region: "eu" } });

  const router = new DocumentStorageRouter({
    orgPolicyService,
    regionalObjectStore: store,
    kmsProvider: kms
  });

  const plaintext = Buffer.from("classified", "utf8");
  const { region, key } = await router.putDocument({ orgId, docId: "doc-1", plaintext });

  assert.equal(region, "eu");

  const raw = await store.getObject(region, key);
  assert.ok(raw, "expected object stored");
  assert.ok(!raw.toString("utf8").includes("classified"), "expected ciphertext-only in storage");

  const roundTripped = await router.getDocument({ orgId, docId: "doc-1" });
  assert.equal(roundTripped.toString("utf8"), "classified");
});
