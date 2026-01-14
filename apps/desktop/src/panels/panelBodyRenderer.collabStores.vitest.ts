// @vitest-environment jsdom

import { act } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import * as Y from "yjs";

import { createCollabSession } from "@formula/collab-session";

const mocks = vi.hoisted(() => {
  const createCollabVersioning = vi.fn((_opts: any) => ({
    listVersions: vi.fn(async () => []),
    createCheckpoint: vi.fn(async () => {}),
    restoreVersion: vi.fn(async () => {}),
    destroy: vi.fn(),
  }));

  const BranchService = vi.fn(function (_opts: any) {
    return {
      init: vi.fn(async () => {}),
      listBranches: vi.fn(async () => []),
      commit: vi.fn(async () => {}),
      createBranch: vi.fn(async () => ({})),
      renameBranch: vi.fn(async () => {}),
      deleteBranch: vi.fn(async () => {}),
      checkoutBranch: vi.fn(async () => ({})),
      previewMerge: vi.fn(async () => ({ conflicts: [], merged: {} })),
      merge: vi.fn(async () => ({ state: {} })),
    };
  });

  const YjsBranchStore = vi.fn(function (_opts: any) {
    return { __sentinel: "yjs-branch-store" };
  });

  const yjsDocToDocumentState = vi.fn(() => ({}));
  const applyDocumentStateToYjsDoc = vi.fn();

  return {
    createCollabVersioning,
    BranchService,
    YjsBranchStore,
    yjsDocToDocumentState,
    applyDocumentStateToYjsDoc,
  };
});

vi.mock("../../../../packages/collab/versioning/src/index.js", () => ({
  createCollabVersioning: mocks.createCollabVersioning,
}));

vi.mock("../../../../packages/versioning/branches/src/browser.js", () => ({
  BranchService: mocks.BranchService,
  YjsBranchStore: mocks.YjsBranchStore,
  yjsDocToDocumentState: mocks.yjsDocToDocumentState,
  applyDocumentStateToYjsDoc: mocks.applyDocumentStateToYjsDoc,
}));

vi.mock("./branch-manager/BranchManagerPanel.js", () => ({
  BranchManagerPanel: () => null,
}));

vi.mock("./branch-manager/MergeBranchPanel.js", () => ({
  MergeBranchPanel: () => null,
}));

// React 18 relies on this flag to suppress act() warnings in test runners.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;

// Import the panel modules at file-evaluation time so Vite transform work does not
// count toward per-test timeouts.
//
// Keep these as dynamic imports (not static `import` statements) so the `vi.mock(...)`
// declarations above apply before the modules are evaluated.
const [{ createPanelBodyRenderer }, { PanelIds }] = await Promise.all([
  import("./panelBodyRenderer.js"),
  import("./panelRegistry.js"),
  // Preload lazy panels so Vite transform time doesn't count against the per-test wait.
  import("./version-history/CollabVersionHistoryPanel.js"),
  import("./branch-manager/CollabBranchManagerPanel.js"),
]);

function flushPromises() {
  return new Promise<void>((resolve) => setTimeout(resolve, 0));
}

async function waitForInAct(condition: () => boolean, timeoutMs = 2_000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (condition()) return;
    // Flush both microtasks and any React effects triggered by them.
    // eslint-disable-next-line no-await-in-loop
    await act(async () => {
      await flushPromises();
    });
  }
  throw new Error("Timed out waiting for condition");
}

afterEach(() => {
  document.body.innerHTML = "";
  mocks.createCollabVersioning.mockClear();
  mocks.BranchService.mockClear();
  mocks.YjsBranchStore.mockClear();
  mocks.yjsDocToDocumentState.mockClear();
  mocks.applyDocumentStateToYjsDoc.mockClear();
});

describe("panelBodyRenderer collab store injection", () => {
  it("uses injected VersionStore when rendering Version History", async () => {
    const session = createCollabSession({ doc: new Y.Doc({ guid: "doc-versioning" }) });
    const sentinelStore = { __sentinel: "version-store" } as any;
    const createCollabVersioningStore = vi.fn(() => sentinelStore);

    const renderer = createPanelBodyRenderer({
      getDocumentController: () => ({}),
      getCollabSession: () => session,
      createCollabVersioningStore,
    });

    const body = document.createElement("div");
    document.body.appendChild(body);

    await act(async () => {
      renderer.renderPanelBody(PanelIds.VERSION_HISTORY, body);
    });

    // The Version History panel lazily imports the collab versioning subsystem; in Node-based
    // test environments the first chunk load can take a few seconds.
    await waitForInAct(() => mocks.createCollabVersioning.mock.calls.length > 0, 10_000);

    expect(createCollabVersioningStore).toHaveBeenCalledWith(session);
    const opts = mocks.createCollabVersioning.mock.calls[0]?.[0] as any;
    expect(opts.session).toBe(session);
    expect(opts.store).toBe(sentinelStore);
  });

  it("uses injected branch store when rendering Branch Manager", async () => {
    const session = createCollabSession({ doc: new Y.Doc({ guid: "doc-branching" }) });
    const sentinelBranchStore = { __sentinel: "branch-store" } as any;
    const createCollabBranchStore = vi.fn(() => sentinelBranchStore);

    const renderer = createPanelBodyRenderer({
      getDocumentController: () => ({}),
      getCollabSession: () => session,
      createCollabBranchStore,
    });

    const body = document.createElement("div");
    document.body.appendChild(body);

    await act(async () => {
      renderer.renderPanelBody(PanelIds.BRANCH_MANAGER, body);
    });

    await waitForInAct(() => mocks.BranchService.mock.calls.length > 0);

    expect(createCollabBranchStore).toHaveBeenCalledWith(session);
    // When an injected store is provided, the panel should not fall back to the Yjs store.
    expect(mocks.YjsBranchStore).not.toHaveBeenCalled();

    const branchServiceOpts = mocks.BranchService.mock.calls[0]?.[0] as any;
    expect(branchServiceOpts.docId).toBe("doc-branching");
    expect(branchServiceOpts.store).toBe(sentinelBranchStore);
  });
});
