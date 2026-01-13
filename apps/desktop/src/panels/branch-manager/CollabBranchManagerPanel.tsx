import React, { useEffect, useMemo, useState } from "react";

import type { CollabSession } from "@formula/collab-session";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver.js";
import { t } from "../../i18n/index.js";

import { BranchManagerPanel, type Actor as BranchActor } from "./BranchManagerPanel.js";
import { MergeBranchPanel } from "./MergeBranchPanel.js";
import { useReservedRootGuardError } from "../collabReservedRootGuard.js";

// Import the browser-safe branch store/service modules directly.
//
// `packages/versioning/branches/src/index.js` also re-exports a Node-only `SQLiteBranchStore`,
// which pulls `node:*` built-ins into the web bundle and breaks the desktop web shell / e2e.
import { BranchService } from "../../../../../packages/versioning/branches/src/BranchService.js";
import { YjsBranchStore } from "../../../../../packages/versioning/branches/src/store/YjsBranchStore.js";
import { applyDocumentStateToYjsDoc, yjsDocToDocumentState } from "../../../../../packages/versioning/branches/src/yjs/index.js";
import { BRANCHING_APPLY_ORIGIN } from "../../collab/conflict-monitors.js";

export function CollabBranchManagerPanel({
  session,
  sheetNameResolver,
}: {
  session: CollabSession;
  sheetNameResolver?: SheetNameResolver | null;
}) {
  const reservedRootGuardError = useReservedRootGuardError((session as any)?.provider ?? null);
  const mutationsDisabled = Boolean(reservedRootGuardError);

  const banner = reservedRootGuardError ? (
    <div className="collab-panel__message collab-panel__message--error" data-testid="reserved-root-guard-error">
      {reservedRootGuardError}
    </div>
  ) : null;

  const actor = useMemo<BranchActor>(() => {
    const userId = session.presence?.localPresence?.id ?? "desktop";
    const roleMaybe = (session as any)?.permissions?.role;
    const role: BranchActor["role"] =
      roleMaybe === "owner" ||
      roleMaybe === "admin" ||
      roleMaybe === "editor" ||
      roleMaybe === "commenter" ||
      roleMaybe === "viewer"
        ? roleMaybe
        : "owner";
    return { userId, role };
  }, [session]);
  const docId = session.doc.guid;

  const { store, storeWarning } = useMemo(() => {
    // Conservative defaults so large branching commits don't exceed common sync-server
    // websocket message limits (close code 1009). Smaller chunks mean more Yjs updates,
    // but keeps the feature usable even when `SYNC_SERVER_MAX_MESSAGE_BYTES` is tuned low.
    const chunkSize = 8 * 1024;
    const maxChunksPerTransaction = 2;

    const proc = (globalThis as any).process;
    const isNodeRuntime = Boolean(proc?.versions?.node);
    const CompressionStreamCtor = (globalThis as any).CompressionStream as any;
    const DecompressionStreamCtor = (globalThis as any).DecompressionStream as any;

    try {
      if (!isNodeRuntime) {
        if (typeof CompressionStreamCtor === "undefined" || typeof DecompressionStreamCtor === "undefined") {
          throw new Error("CompressionStream is unavailable");
        }
        // Some runtimes expose the constructor but don't support gzip.
        // (If either throws, fall back to JSON payloads.)
        void new CompressionStreamCtor("gzip");
        void new DecompressionStreamCtor("gzip");
      }
      return {
        store: new YjsBranchStore({
          ydoc: session.doc,
          payloadEncoding: "gzip-chunks",
          chunkSize,
          maxChunksPerTransaction,
        }),
        storeWarning: null as string | null,
      };
    } catch {
      return {
        store: new YjsBranchStore({ ydoc: session.doc, payloadEncoding: "json" }),
        storeWarning: t("branchManager.compressionFallbackWarning"),
      };
    }
  }, [session.doc]);
  const branchService = useMemo(() => new BranchService({ docId, store }), [docId, store]);

  const [error, setError] = useState<string | null>(null);
  const [ready, setReady] = useState(false);
  const [mergeSource, setMergeSource] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        setError(null);
        setReady(false);
        const initialState = yjsDocToDocumentState(session.doc);
        await branchService.init(actor as any, initialState as any);
        if (cancelled) return;
        setReady(true);
      } catch (e) {
        if (cancelled) return;
        setError((e as Error).message);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [actor, branchService, session]);

  const workflow = useMemo(() => {
    const commitCurrentState = async (message: string) => {
      const nextState = yjsDocToDocumentState(session.doc);
      await branchService.commit(actor as any, { nextState, message });
    };

    return {
      getCurrentBranchName: () => branchService.getCurrentBranchName(),
      listBranches: () => branchService.listBranches(),
      createBranch: async (a: BranchActor, input: { name: string; description?: string }) => {
        await commitCurrentState("auto: create branch");
        return branchService.createBranch(a as any, input as any);
      },
      renameBranch: (a: BranchActor, input: { oldName: string; newName: string }) =>
        branchService.renameBranch(a as any, input as any),
      deleteBranch: (a: BranchActor, input: { name: string }) => branchService.deleteBranch(a as any, input as any),
      checkoutBranch: async (a: BranchActor, input: { name: string }) => {
        await commitCurrentState("auto: checkout");
        const state = await branchService.checkoutBranch(a as any, input as any);
        // Branch checkout is a bulk "time travel" operation and must not be captured by
        // collaborative undo tracking. CollabSession also treats this origin as ignored
        // for conflict monitors so it doesn't surface spurious conflicts.
        applyDocumentStateToYjsDoc(session.doc, state as any, { origin: BRANCHING_APPLY_ORIGIN });
        return state;
      },
      previewMerge: async (a: BranchActor, input: { sourceBranch: string }) => {
        await commitCurrentState("auto: preview merge");
        return branchService.previewMerge(a as any, input as any);
      },
      merge: async (a: BranchActor, input: { sourceBranch: string; resolutions: any[]; message?: string }) => {
        await commitCurrentState("auto: merge");
        const result = await branchService.merge(a as any, input as any);
        // See checkoutBranch origin note above.
        applyDocumentStateToYjsDoc(session.doc, (result as any).state, { origin: BRANCHING_APPLY_ORIGIN });
        return result;
      },
    } as any;
  }, [actor, branchService, session]);

  if (error) {
    return (
      <div className="collab-branch-manager">
        {storeWarning ? <div className="collab-panel__message collab-panel__message--warning">{storeWarning}</div> : null}
        {banner}
        <div className="collab-panel__message collab-panel__message--error">{error}</div>
      </div>
    );
  }

  if (!ready) {
    return (
      <div className="collab-branch-manager">
        {storeWarning ? <div className="collab-panel__message collab-panel__message--warning">{storeWarning}</div> : null}
        {banner}
        <div className="collab-panel__message">Loading branches…</div>
      </div>
    );
  }

  if (mergeSource) {
    return (
      <div className="collab-branch-manager">
        {storeWarning ? <div className="collab-panel__message collab-panel__message--warning">{storeWarning}</div> : null}
        {banner}
        <MergeBranchPanel
          actor={actor}
          branchService={workflow}
          sourceBranch={mergeSource}
          sheetNameResolver={sheetNameResolver ?? null}
          mutationsDisabled={mutationsDisabled}
          onClose={() => setMergeSource(null)}
        />
      </div>
    );
  }

  return (
    <div className="collab-branch-manager">
      {storeWarning ? <div className="collab-panel__message collab-panel__message--warning">{storeWarning}</div> : null}
      {banner}
      <BranchManagerPanel
        actor={actor}
        branchService={workflow}
        mutationsDisabled={mutationsDisabled}
        onStartMerge={(sourceBranch) => {
          if (mutationsDisabled) return;
          setMergeSource(sourceBranch);
        }}
      />
    </div>
  );
}
