import React, { useEffect, useMemo, useState } from "react";

import type { CollabSession } from "@formula/collab-session";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver.js";
import { t } from "../../i18n/index.js";

import { BranchManagerPanel, type Actor as BranchActor } from "./BranchManagerPanel.js";
import { MergeBranchPanel } from "./MergeBranchPanel.js";
import { clearReservedRootGuardError, useReservedRootGuardError } from "../collabReservedRootGuard.js";

// Import branching helpers from the browser-safe entrypoint so bundlers don't
// accidentally pull Node-only stores (e.g. SQLite) into the WebView bundle.
import {
  BranchService,
  YjsBranchStore,
  applyDocumentStateToYjsDoc,
  yjsDocToDocumentState,
} from "../../../../../packages/versioning/branches/src/browser.js";
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

  const localPresenceId = session.presence?.localPresence?.id;
  const sessionPermissions = (session as any)?.permissions as { role?: unknown; userId?: unknown } | null | undefined;
  const permissionsRole = sessionPermissions?.role;
  const permissionsUserId = sessionPermissions?.userId;

  const actor = useMemo<BranchActor>(() => {
    const userId =
      (typeof permissionsUserId === "string" && permissionsUserId.length > 0 ? permissionsUserId : null) ??
      localPresenceId ??
      "desktop";
    const roleMaybe = permissionsRole;
    const role: BranchActor["role"] =
      roleMaybe === "owner" ||
      roleMaybe === "admin" ||
      roleMaybe === "editor" ||
      roleMaybe === "commenter" ||
      roleMaybe === "viewer"
        ? roleMaybe
        : "owner";
    return { userId, role };
  }, [localPresenceId, permissionsRole, permissionsUserId]);
  const docId = session.doc.guid;

  const { store, storeWarning } = useMemo(() => {
    if (mutationsDisabled) {
      return { store: null as any, storeWarning: null as string | null };
    }
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
  }, [session.doc, mutationsDisabled]);

  const branchService = useMemo(() => {
    if (!store) return null;
    return new BranchService({ docId, store });
  }, [docId, store]);

  const [error, setError] = useState<string | null>(null);
  const [ready, setReady] = useState(false);
  const [mergeSource, setMergeSource] = useState<string | null>(null);

  const banner = reservedRootGuardError ? (
    <div className="collab-panel__message collab-panel__message--error" data-testid="reserved-root-guard-error">
      <div>{reservedRootGuardError}</div>
      <button
        type="button"
        onClick={() => {
          clearReservedRootGuardError((session as any)?.provider ?? null);
          setError(null);
          setReady(false);
          setMergeSource(null);
        }}
      >
        {t("collab.retry")}
      </button>
    </div>
  ) : null;

  useEffect(() => {
    // If the sync server has disconnected due to reserved root mutations, branch
    // merge cannot proceed. Close any in-progress merge UI so we don't strand the
    // user on a "Loading…" screen (MergeBranchPanel intentionally skips fetching
    // previews when mutations are disabled).
    if (!mutationsDisabled) return;
    setMergeSource(null);
  }, [mutationsDisabled]);

  useEffect(() => {
    if (mutationsDisabled) {
      setError(null);
      setReady(true);
      return;
    }
    if (!branchService) return;
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
  }, [actor, branchService, session, mutationsDisabled]);

  const workflow = useMemo(() => {
    if (!branchService) {
      const disabled = async () => {
        throw new Error(reservedRootGuardError ?? "Branching is unavailable");
      };
      return {
        getCurrentBranchName: async () => "main",
        listBranches: async () => [],
        createBranch: disabled,
        renameBranch: disabled,
        deleteBranch: disabled,
        checkoutBranch: disabled,
        previewMerge: disabled,
        merge: disabled,
      } as any;
    }

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
  }, [actor, branchService, session, reservedRootGuardError]);

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
