import type * as Y from "yjs";
import type {
  CollabPersistence,
  CollabPersistenceBinding,
  CollabPersistenceFlushOptions,
} from "@formula/collab-persistence";

/**
 * Lazy wrapper around `@formula/collab-persistence/indexeddb`.
 *
 * This keeps the IndexedDB implementation out of the desktop entry chunk so
 * local (non-collab) startup doesn't pay its parse/execute cost.
 *
 * The underlying module is loaded on first use (`load`/`bind`/`flush`/etc).
 */
export class LazyIndexedDbCollabPersistence implements CollabPersistence {
  private instance: CollabPersistence | null = null;
  private instancePromise: Promise<CollabPersistence> | null = null;

  private async getInstance(): Promise<CollabPersistence> {
    if (this.instance) return this.instance;
    if (!this.instancePromise) {
      this.instancePromise = import("@formula/collab-persistence/indexeddb")
        .then((mod: any) => {
          const inst = new mod.IndexedDbCollabPersistence();
          this.instance = inst as CollabPersistence;
          return inst as CollabPersistence;
        })
        .catch((err) => {
          // Allow retries if the first dynamic import fails.
          this.instancePromise = null;
          throw err;
        });
    }
    return this.instancePromise;
  }

  async load(docId: string, doc: Y.Doc): Promise<void> {
    const inst = await this.getInstance();
    await inst.load(docId, doc);
  }

  bind(docId: string, doc: Y.Doc): CollabPersistenceBinding {
    let destroyed = false;
    let binding: CollabPersistenceBinding | null = null;
    let bindingPromise: Promise<CollabPersistenceBinding> | null = null;

    const ensureBinding = async (): Promise<CollabPersistenceBinding> => {
      if (binding) return binding;
      if (!bindingPromise) {
        bindingPromise = (async () => {
          const inst = await this.getInstance();
          const next = inst.bind(docId, doc);
          binding = next;
          if (destroyed) {
            // If the consumer destroyed the binding before the lazy import settled, ensure
            // we don't leak listeners/resources.
            await next.destroy().catch(() => {});
          }
          return next;
        })();
      }
      return bindingPromise;
    };

    return {
      destroy: async () => {
        destroyed = true;
        if (binding) {
          await binding.destroy();
          return;
        }
        if (bindingPromise) {
          const resolved = await bindingPromise.catch(() => null);
          if (resolved) await resolved.destroy().catch(() => {});
        }
      },
    };
  }

  async flush(docId: string, opts?: CollabPersistenceFlushOptions): Promise<void> {
    const inst = await this.getInstance();
    if (typeof inst.flush === "function") {
      await inst.flush(docId, opts);
    }
  }

  async compact(docId: string): Promise<void> {
    const inst = await this.getInstance();
    if (typeof inst.compact === "function") {
      await inst.compact(docId);
    }
  }

  async clear(docId: string): Promise<void> {
    const inst = await this.getInstance();
    if (typeof inst.clear === "function") {
      await inst.clear(docId);
    }
  }
}

