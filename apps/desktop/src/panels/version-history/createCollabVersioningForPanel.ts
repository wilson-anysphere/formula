import type { CollabSession } from "@formula/collab-session";

import type { CollabVersioning, VersionStore } from "../../../../../packages/collab/versioning/src/index.ts";

export type MaybePromise<T> = T | Promise<T>;

/**
 * Factory hook for providing an alternate VersionStore implementation (e.g.
 * ApiVersionStore, SQLite) that does not write to reserved Yjs roots like
 * `versions*`.
 */
export type CreateVersionStore = (session: CollabSession) => MaybePromise<VersionStore>;

/**
 * Create a `CollabVersioning` instance for the desktop Version History panel.
 *
 * This intentionally loads `@formula/collab-versioning` lazily to avoid pulling
 * it into the desktop shell startup bundle.
 */
export async function createCollabVersioningForPanel({
  session,
  store,
  createVersionStore,
}: {
  session: CollabSession;
  store?: VersionStore;
  createVersionStore?: CreateVersionStore;
}): Promise<CollabVersioning> {
  const mod = await import("../../../../../packages/collab/versioning/src/index.js");
  const resolvedStore = store ?? (createVersionStore ? await createVersionStore(session) : undefined);

  const localPresence = session.presence?.localPresence ?? null;
  return mod.createCollabVersioning({
    session,
    store: resolvedStore,
    user: localPresence ? { userId: localPresence.id, userName: localPresence.name } : undefined,
  }) as CollabVersioning;
}

