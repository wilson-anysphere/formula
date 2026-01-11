import type { DocumentController } from "../document/documentController.js";

// DocumentController batching is global (single active batch), so overlapping Power Query
// apply operations can corrupt undo grouping and prevent cancellation from reverting
// partial writes. Serialize apply work per-document across all orchestrator instances.

const APPLY_QUEUE_BY_DOC: WeakMap<DocumentController, Promise<unknown>> = new WeakMap();

export function enqueueApplyForDocument<T>(doc: DocumentController, work: () => Promise<T>): Promise<T> {
  const prior = APPLY_QUEUE_BY_DOC.get(doc) ?? Promise.resolve();
  const resultPromise = prior.then(work);
  // Keep the internal queue in a resolved state so subsequent work always runs.
  APPLY_QUEUE_BY_DOC.set(
    doc,
    resultPromise.then(
      () => undefined,
      () => undefined,
    ),
  );
  return resultPromise;
}
