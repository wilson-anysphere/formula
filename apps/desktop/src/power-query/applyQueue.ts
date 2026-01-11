import type { DocumentController } from "../document/documentController.js";

// DocumentController batching is global (single active batch), so overlapping Power Query
// apply operations can corrupt undo grouping and prevent cancellation from reverting
// partial writes. Serialize apply work per-document across all orchestrator instances.

const APPLY_QUEUE_BY_DOC: WeakMap<DocumentController, Promise<void>> = new WeakMap();

export function enqueueApplyForDocument(doc: DocumentController, work: () => Promise<void>): Promise<void> {
  const prior = APPLY_QUEUE_BY_DOC.get(doc) ?? Promise.resolve();
  const next = prior.then(work).catch(() => {});
  APPLY_QUEUE_BY_DOC.set(doc, next);
  return next;
}

