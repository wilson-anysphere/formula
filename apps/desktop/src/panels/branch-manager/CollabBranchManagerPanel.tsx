import React, { useEffect, useMemo, useState } from "react";

import type { CollabSession } from "@formula/collab-session";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver.js";
import { t } from "../../i18n/index.js";

import { BranchManagerPanel, type Actor as BranchActor } from "./BranchManagerPanel.js";
import { MergeBranchPanel } from "./MergeBranchPanel.js";
import { clearReservedRootGuardError, useReservedRootGuardError } from "../collabReservedRootGuard.react.js";
import { useCollabSessionSyncState } from "../useCollabSessionSyncState.js";
import type { BranchStore, CreateBranchStore } from "./branchStoreTypes.js";
import { resolveBranchStoreForPanel } from "./resolveBranchStoreForPanel.js";
import { commitIfDocumentStateChanged } from "./commitIfChanged.js";

// Import branching helpers from the browser-safe entrypoint so bundlers don't
// accidentally pull Node-only stores (e.g. SQLite) into the WebView bundle.
import {
  BranchService,
  applyDocumentStateToYjsDoc,
  yjsDocToDocumentState,
} from "../../../../../packages/versioning/branches/src/browser.js";
import { BRANCHING_APPLY_ORIGIN } from "../../collab/conflict-monitors.js";

export function CollabBranchManagerPanel({
  session,
  sheetNameResolver,
  createBranchStore,
  branchStore,
}: {
  session: CollabSession;
  sheetNameResolver?: SheetNameResolver | null;
  /**
   * Optional BranchStore provider. Use this to inject an out-of-doc store so the
   * panel does not write to reserved Yjs roots (`branching:*`).
   */
  createBranchStore?: CreateBranchStore;
  /**
   * Optional pre-constructed BranchStore instance. Prefer {@link createBranchStore}
   * to keep store construction lazy.
   */
  branchStore?: BranchStore;
}) {
  const syncState = useCollabSessionSyncState(session);
  const hasInjectedStore = Boolean(createBranchStore || branchStore);
  const reservedRootGuardError = useReservedRootGuardError((session as any)?.provider ?? null);
  // Reserved root guard disconnects are sticky (we remember them per provider) so
  // panels opened later can show the banner. When using an out-of-doc store, we
  // only need to disable actions if the provider is currently disconnected.
  const mutationsDisabled = Boolean(reservedRootGuardError) && (!hasInjectedStore || !syncState.connected);

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

  const [store, setStore] = useState<BranchStore | null>(null);
  const [storeWarning, setStoreWarning] = useState<string | null>(null);

  const branchService = useMemo(() => {
    if (!store) return null;
    if (mutationsDisabled) return null;
    return new BranchService({ docId, store });
  }, [docId, mutationsDisabled, store]);

  const [error, setError] = useState<string | null>(null);
  const [ready, setReady] = useState(false);
  const [mergeSource, setMergeSource] = useState<string | null>(null);

  const banner = mutationsDisabled && reservedRootGuardError ? (
    <div className="collab-panel__message collab-panel__message--error" data-testid="reserved-root-guard-error">
      <div>{reservedRootGuardError}</div>
      <button
        type="button"
        data-testid="reserved-root-guard-retry"
        onClick={() => {
          clearReservedRootGuardError((session as any)?.provider ?? null);
          try {
            (session as any)?.provider?.connect?.();
          } catch {
            // ignore
          }
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
    if (mutationsDisabled) {
      setStore(null);
      setStoreWarning(null);
      return;
    }
    let cancelled = false;
    setError(null);
    setReady(false);
    setStore(null);
    setStoreWarning(null);

    void (async () => {
      try {
        const resolved = await resolveBranchStoreForPanel({
          session,
          store: branchStore,
          createBranchStore,
          compressionFallbackWarning: t("branchManager.compressionFallbackWarning"),
        });
        if (cancelled) return;
        setStore(resolved.store);
        setStoreWarning(resolved.storeWarning);
      } catch (e) {
        if (cancelled) return;
        setError((e as Error).message);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [branchStore, createBranchStore, mutationsDisabled, session]);

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
      await commitIfDocumentStateChanged({
        actor: actor as any,
        branchService: branchService as any,
        doc: session.doc,
        message,
        docToState: yjsDocToDocumentState as any,
      });
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
      {error ? <div className="collab-panel__message collab-panel__message--error">{error}</div> : null}
      {!ready && !error ? (
        <div role="status" className="collab-panel__message">
          {t("branchManager.loading")}
        </div>
      ) : null}
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
