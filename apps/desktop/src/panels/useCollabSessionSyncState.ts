import { useEffect, useState } from "react";

import type { CollabSession } from "@formula/collab-session";

export type CollabSessionSyncState = { connected: boolean; synced: boolean };

/**
 * Best-effort hook for observing {@link CollabSession}'s provider connection state.
 *
 * Used by collab panels to decide whether a reserved-root-guard disconnect is
 * currently impacting the UI (vs a historical close event that has since
 * reconnected).
 */
export function useCollabSessionSyncState(session: CollabSession | null): CollabSessionSyncState {
  const [state, setState] = useState<CollabSessionSyncState>(() => {
    const getSyncState = (session as any)?.getSyncState;
    if (typeof getSyncState !== "function") return { connected: false, synced: false };
    try {
      return getSyncState.call(session) as CollabSessionSyncState;
    } catch {
      return { connected: false, synced: false };
    }
  });

  useEffect(() => {
    const getSyncState = (session as any)?.getSyncState;
    const onStatusChange = (session as any)?.onStatusChange;
    if (!session || typeof getSyncState !== "function" || typeof onStatusChange !== "function") {
      setState({ connected: false, synced: false });
      return;
    }

    let disposed = false;
    const safeSet = (next: CollabSessionSyncState) => {
      if (disposed) return;
      setState(next);
    };

    try {
      safeSet(getSyncState.call(session) as CollabSessionSyncState);
    } catch {
      // ignore
    }

    let unsubscribe: (() => void) | null = null;
    try {
      unsubscribe = onStatusChange.call(session, (next: CollabSessionSyncState) => {
        safeSet(next);
      }) as (() => void) | null;
    } catch {
      unsubscribe = null;
    }

    return () => {
      disposed = true;
      try {
        unsubscribe?.();
      } catch {
        // ignore
      }
    };
  }, [session]);

  return state;
}

