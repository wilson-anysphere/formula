import assert from "node:assert/strict";
import test from "node:test";

import { createCollabSession } from "@formula/collab-session";

// Explicit `.ts` imports so the repo's node:test runner can skip this suite when
// TypeScript execution isn't available.
import { createCollabVersioningForPanel } from "../src/panels/version-history/createCollabVersioningForPanel.ts";
import { resolveBranchStoreForPanel } from "../src/panels/branch-manager/resolveBranchStoreForPanel.ts";

test("collab panels use injected VersionStore/BranchStore factories when provided", async () => {
  const session = createCollabSession();

  const versionStore = {
    async saveVersion() {},
    async getVersion() {
      return null;
    },
    async listVersions() {
      return [];
    },
    async updateVersion() {},
    async deleteVersion() {},
  };

  const branchStore = {
    async ensureDocument() {},
    async listBranches() {
      return [];
    },
    async getBranch() {
      return null;
    },
    async createBranch() {
      return null;
    },
    async renameBranch() {},
    async deleteBranch() {},
    async updateBranchHead() {},
    async createCommit() {
      return null;
    },
    async getCommit() {
      return null;
    },
    async getDocumentStateAtCommit() {
      return { schemaVersion: 1, sheets: { order: [], metaById: {} }, cells: {}, metadata: {}, namedRanges: {}, comments: {} };
    },
  };

  let versionStoreCalls = 0;
  let branchStoreCalls = 0;

  const panelRendererOptions = {
    createVersionStore: async () => {
      versionStoreCalls += 1;
      return versionStore;
    },
    createBranchStore: async () => {
      branchStoreCalls += 1;
      return branchStore;
    },
  };

  const versioning = await createCollabVersioningForPanel({ session, createVersionStore: panelRendererOptions.createVersionStore });
  assert.equal(versionStoreCalls, 1, "expected createVersionStore to be called");
  assert.equal(versioning.manager.store, versionStore, "expected injected VersionStore to be used");
  versioning.destroy();

  const resolvedBranchStore = await resolveBranchStoreForPanel({
    session,
    createBranchStore: panelRendererOptions.createBranchStore,
    compressionFallbackWarning: "fallback",
  });
  assert.equal(branchStoreCalls, 1, "expected createBranchStore to be called");
  assert.equal(resolvedBranchStore.store, branchStore, "expected injected BranchStore to be used");
  assert.equal(resolvedBranchStore.storeWarning, null);

  session.destroy();
  session.doc.destroy();
});
