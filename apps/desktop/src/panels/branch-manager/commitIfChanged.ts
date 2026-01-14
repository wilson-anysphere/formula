import { normalizeDocumentState } from "../../../../../packages/versioning/branches/src/state.js";

type BranchServiceLike = {
  getCurrentState(): Promise<unknown>;
  commit(actor: unknown, input: { nextState: unknown; message: string }): Promise<unknown>;
};

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return (
    value !== null &&
    typeof value === "object" &&
    (Object.getPrototypeOf(value) === Object.prototype || Object.getPrototypeOf(value) === null)
  );
}

/**
 * Deep equality for JSON-ish values (plain objects/arrays/primitives).
 *
 * This intentionally does not attempt to support Maps/Sets/classes; DocumentState is
 * expected to be structured-clone / JSON serializable.
 */
export function deepEqualJson(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (Number.isNaN(a) && Number.isNaN(b)) return true;

  if (Array.isArray(a) && Array.isArray(b)) {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i += 1) {
      if (!deepEqualJson(a[i], b[i])) return false;
    }
    return true;
  }

  if (isPlainObject(a) && isPlainObject(b)) {
    const keysA = Object.keys(a);
    const keysB = Object.keys(b);
    if (keysA.length !== keysB.length) return false;
    for (const key of keysA) {
      if (!Object.prototype.hasOwnProperty.call(b, key)) return false;
      if (!deepEqualJson(a[key], b[key])) return false;
    }
    return true;
  }

  return false;
}

export function documentStatesEqual(a: unknown, b: unknown): boolean {
  // Normalize into the canonical BranchService state shape so semantically-identical
  // states (e.g. missing optional workbook maps) compare equal.
  const normA = normalizeDocumentState(a);
  const normB = normalizeDocumentState(b);
  return deepEqualJson(normA, normB);
}

/**
 * Commit the current document state iff it differs from the current branch head.
 *
 * Returns `true` when a commit was created, `false` when skipped as a no-op.
 */
export async function commitIfDocumentStateChanged<TDoc>({
  actor,
  branchService,
  doc,
  message,
  docToState,
}: {
  actor: unknown;
  branchService: BranchServiceLike;
  doc: TDoc;
  message: string;
  docToState: (doc: TDoc) => unknown;
}): Promise<boolean> {
  const nextState = docToState(doc);
  const currentState = await branchService.getCurrentState();
  if (documentStatesEqual(currentState, nextState)) return false;
  await branchService.commit(actor, { nextState, message });
  return true;
}

