import type { CollabSession } from "@formula/collab-session";

import type { CollabVersioning, VersionStore } from "../../../../../packages/collab/versioning/src/index.ts";

export type MaybePromise<T> = T | Promise<T>;

let collabVersioningModulePromise: Promise<any> | null = null;

function isNodeRuntime(): boolean {
  const proc = (globalThis as any).process;
  // Require `process.release.name === "node"` to avoid false positives from lightweight
  // `process` polyfills some bundlers inject into browser environments.
  return Boolean(proc?.versions?.node) && proc?.release?.name === "node";
}

function loadCollabVersioningModule(): Promise<any> {
  if (!collabVersioningModulePromise) {
    collabVersioningModulePromise = import("../../../../../packages/collab/versioning/src/index.js");
  }
  return collabVersioningModulePromise;
}

// In Node-based environments (vitest / unit tests), Vite's first dynamic import of
// the versioning subsystem can be slow enough that panels don't mount within the
// default 1s test wait. Kick off the chunk load eagerly so the panel UI can
// become interactive quickly once mounted, while keeping the module lazy in
// browser/WebView builds.
if (isNodeRuntime()) {
  void loadCollabVersioningModule().catch(() => {
    // Best-effort prefetch; errors are surfaced when the panel actually tries to
    // construct CollabVersioning.
  });
}

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
  const mod = await loadCollabVersioningModule();
  const resolvedStore = store ?? (createVersionStore ? await createVersionStore(session) : undefined);

  const localPresence = session.presence?.localPresence ?? null;
  return mod.createCollabVersioning({
    session,
    store: resolvedStore,
    user: localPresence ? { userId: localPresence.id, userName: localPresence.name } : undefined,
  }) as CollabVersioning;
}
