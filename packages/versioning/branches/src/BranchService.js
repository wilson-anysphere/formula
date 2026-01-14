import { diffDocumentStates } from "./patch.js";
import { applyConflictResolutions, mergeDocumentStates } from "./merge.js";
import { normalizeDocumentState } from "./state.js";

/**
 * @typedef {import("./types.js").Actor} Actor
 * @typedef {import("./types.js").DocumentState} DocumentState
 * @typedef {import("./types.js").Branch} Branch
 * @typedef {import("./types.js").Commit} Commit
 * @typedef {import("./types.js").MergeResult} MergeResult
 * @typedef {import("./merge.js").ConflictResolution} ConflictResolution
 */

/**
 * Branch mutations (create/rename/delete/checkout/merge) are restricted to
 * document owners/admins. Editors can still create commits (see
 * `assertCanCommit`).
 *
 * @param {Actor} actor
 * @param {string} operation
 */
function assertCanManageBranches(actor, operation) {
  if (actor.role !== "owner" && actor.role !== "admin") {
    throw new Error(`${operation} requires owner/admin permissions (role=${actor.role})`);
  }
}

/**
 * Commits mutate document history but are allowed for any role that can edit the
 * document contents.
 *
 * @param {Actor} actor
 */
function assertCanCommit(actor) {
  if (actor.role !== "owner" && actor.role !== "admin" && actor.role !== "editor") {
    throw new Error(`Commit requires edit permission (role=${actor.role})`);
  }
}

/**
 * @param {any} value
 * @returns {value is Record<string, any>}
 */
function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

/**
 * Backwards-compatible detection for the BranchService v0 state shape:
 * `{ sheets: Record<sheetId, CellMap> }`.
 *
 * Old clients only know about cell edits; they should not be able to accidentally
 * wipe workbook metadata (sheet names/order, metadata map, namedRanges,
 * comments) by omitting those fields from commits.
 *
 * @param {any} value
 */
function isLegacyCellsOnlyState(value) {
  return (
    isRecord(value) &&
    value.schemaVersion !== 1 &&
    !("cells" in value) &&
    isRecord(value.sheets)
  );
}

/**
 * Backwards-compatible detection for malformed/partial schemaVersion=1 payloads
 * that only include `cells` but not a valid `sheets` ordering/metadata.
 *
 * Treat these like legacy cell-only commits so callers can't accidentally wipe
 * sheet ordering/names by omitting them.
 *
 * @param {any} value
 */
function isSchemaV1CellsOnlyState(value) {
  if (!isRecord(value) || value.schemaVersion !== 1) return false;
  const sheets = value.sheets;
  return !isRecord(sheets) || !Array.isArray(sheets.order) || !isRecord(sheets.metaById);
}

/**
 * Backwards-compatible detection for older schemaVersion=1 clients that existed
 * before BranchService started tracking some workbook-level maps (metadata,
 * namedRanges, comments).
 *
 * Those clients will send a schemaVersion=1 state missing those fields (or with
 * `null`/`undefined`/invalid values). Treat that as an overlay on the current
 * branch head so older callers cannot accidentally delete keys they don't know
 * about.
 *
 * @param {any} value
 * @param {"metadata" | "namedRanges" | "comments"} field
 */
function shouldPreserveSchemaV1WorkbookMap(value, field) {
  if (!isRecord(value) || value.schemaVersion !== 1) return false;
  if (!(field in value)) return true;
  return !isRecord(value[field]);
}

/**
 * Backwards-compatible detection for schemaVersion=1 callers that existed
 * before BranchService tracked per-sheet `view` state (e.g. frozen panes).
 *
 * Those callers will send sheet metadata maps that omit `view` (and any
 * `frozenRows`/`frozenCols` keys). Treat the missing view as "no change" so
 * older clients can't accidentally wipe existing sheet view state.
 *
 * @param {any} value
 * @param {string} sheetId
 */
function shouldPreserveSchemaV1SheetView(value, sheetId) {
  if (!isRecord(value) || value.schemaVersion !== 1) return false;
  const sheets = value.sheets;
  if (!isRecord(sheets) || !isRecord(sheets.metaById)) return false;
  const meta = sheets.metaById[sheetId];
  if (!isRecord(meta)) return false;

  // If the sheet meta explicitly includes view state (either nested under `view`
  // or as legacy top-level fields), then it's an intentional update.
  if ("view" in meta) {
    const view = meta.view;
    // Treat non-object view values (null/undefined) as omitted.
    if (isRecord(view)) return false;
    // If a key exists but is invalid, assume the caller doesn't support view and preserve.
    return true;
  }
  if (
    "frozenRows" in meta ||
    "frozenCols" in meta ||
    "backgroundImageId" in meta ||
    "background_image_id" in meta ||
    "backgroundImage" in meta ||
    "background_image" in meta ||
    "colWidths" in meta ||
    "rowHeights" in meta ||
    "mergedRanges" in meta ||
    "mergedCells" in meta ||
    "merged_cells" in meta ||
    "merged_ranges" in meta ||
    "mergedRegions" in meta ||
    "merged_regions" in meta ||
    "drawings" in meta
  ) {
    return false;
  }
  return true;
}

/**
 * Backwards-compatible detection for schemaVersion=1 callers that existed before
 * BranchService tracked sheet visibility/tabColor metadata.
 *
 * Older callers will omit these fields from `sheets.metaById[sheetId]`. Treat
 * omission (or invalid values) as "no change" so they can't accidentally wipe
 * sheet metadata they don't understand.
 *
 * @param {any} value
 * @param {string} sheetId
 */
function shouldPreserveSchemaV1SheetVisibility(value, sheetId) {
  if (!isRecord(value) || value.schemaVersion !== 1) return false;
  const sheets = value.sheets;
  if (!isRecord(sheets) || !isRecord(sheets.metaById)) return false;
  const meta = sheets.metaById[sheetId];
  if (!isRecord(meta)) return false;

  if ("visibility" in meta) {
    const vis = meta.visibility;
    if (vis === "visible" || vis === "hidden" || vis === "veryHidden") return false;
    // Key exists but is invalid -> assume caller doesn't support; preserve.
    return true;
  }
  return true;
}

/**
 * @param {any} value
 * @param {string} sheetId
 */
function shouldPreserveSchemaV1SheetTabColor(value, sheetId) {
  if (!isRecord(value) || value.schemaVersion !== 1) return false;
  const sheets = value.sheets;
  if (!isRecord(sheets) || !isRecord(sheets.metaById)) return false;
  const meta = sheets.metaById[sheetId];
  if (!isRecord(meta)) return false;

  if ("tabColor" in meta) {
    const c = meta.tabColor;
    // `null` is an explicit "clear tab color" operation.
    if (c === null) return false;
    if (typeof c === "string") {
      // Treat invalid strings as omitted so older callers can't wipe a valid tabColor
      // with a non-canonical representation.
      if (/^[0-9A-Fa-f]{8}$/.test(c)) return false;
      return true;
    }
    // Key exists but is invalid -> assume caller doesn't support; preserve.
    return true;
  }
  return true;
}

/**
 * BranchService provides high-level branch/merge operations for a single
 * document.
 *
 * It is intentionally UI-agnostic; the desktop panel wires user interactions to
 * these methods.
 */
export class BranchService {
  /** @type {string} */
  #docId;
  /** @type {any} */
  #store;
  /** @type {string} */
  #currentBranchName = "main";

  /**
   * @param {{ docId: string, store: any }} input
   */
  constructor({ docId, store }) {
    this.#docId = docId;
    this.#store = store;
  }

  async #getCurrentBranchName() {
    const store = this.#store;
    if (store && typeof store.getCurrentBranchName === "function") {
      const name = await store.getCurrentBranchName(this.#docId);
      if (typeof name === "string" && name.length > 0) {
        // Keep the local cache in sync even when the store is authoritative.
        this.#currentBranchName = name;
        return name;
      }
    }
    return this.#currentBranchName;
  }

  /**
   * Read the current branch name.
   *
   * For collaborative stores (e.g. YjsBranchStore), this reflects the global
   * checked-out branch shared via the store; for local-only stores it reflects
   * this BranchService instance's local pointer.
   *
   * @returns {Promise<string>}
   */
  async getCurrentBranchName() {
    return this.#getCurrentBranchName();
  }

  /**
   * @param {string} name
   */
  async #setCurrentBranchName(name) {
    const store = this.#store;
    if (store && typeof store.setCurrentBranchName === "function") {
      await store.setCurrentBranchName(this.#docId, name);
      this.#currentBranchName = name;
      return;
    }
    this.#currentBranchName = name;
  }

  /**
   * @param {Actor} actor
   * @param {DocumentState} initialState
   */
  async init(actor, initialState) {
    // `ensureDocument` will create the root commit + main branch if the document
    // doesn't exist yet. Creating a new document is an admin-level action, but
    // calling `init` on an existing document is safe for any role (it becomes a
    // no-op in the store).
    const store = this.#store;
    const exists =
      store && typeof store.hasDocument === "function"
        ? await store.hasDocument(this.#docId)
        : Boolean(await store.getBranch(this.#docId, "main"));
    if (!exists) {
      assertCanManageBranches(actor, "init");
    }
    await this.#store.ensureDocument(this.#docId, actor, initialState);
  }

  async listBranches() {
    return this.#store.listBranches(this.#docId);
  }

  /**
   * @returns {Promise<Branch>}
   */
  async getCurrentBranch() {
    const name = await this.#getCurrentBranchName();
    const branch = await this.#store.getBranch(this.#docId, name);
    if (!branch) throw new Error(`Current branch not found: ${name}`);
    return branch;
  }

  /**
   * Convenience helper: load the current branch head state.
   *
   * This is intentionally *not* permission-gated (it's a read of already
   * reachable history) and is useful for adapters that can't represent the full
   * workbook metadata surface area (e.g. DocumentController).
   *
   * @returns {Promise<DocumentState>}
   */
  async getCurrentState() {
    const branch = await this.getCurrentBranch();
    return this.#store.getDocumentStateAtCommit(branch.headCommitId);
  }

  /**
   * @param {Actor} actor
   * @param {{ name: string, description?: string }} input
   */
  async createBranch(actor, { name, description }) {
    assertCanManageBranches(actor, "createBranch");
    const current = await this.getCurrentBranch();
    return this.#store.createBranch({
      docId: this.#docId,
      name,
      createdBy: actor.userId,
      createdAt: Date.now(),
      description: description ?? null,
      headCommitId: current.headCommitId
    });
  }

  /**
   * @param {Actor} actor
   * @param {{ oldName: string, newName: string }} input
   */
  async renameBranch(actor, { oldName, newName }) {
    assertCanManageBranches(actor, "renameBranch");
    await this.#store.renameBranch(this.#docId, oldName, newName);
    const current = await this.#getCurrentBranchName();
    if (oldName === current) await this.#setCurrentBranchName(newName);
  }

  /**
   * @param {Actor} actor
   * @param {{ name: string }} input
   */
  async deleteBranch(actor, { name }) {
    assertCanManageBranches(actor, "deleteBranch");
    if (name === "main") throw new Error("Cannot delete main branch");
    const current = await this.#getCurrentBranchName();
    if (name === current) {
      throw new Error("Cannot delete the currently checked-out branch");
    }
    await this.#store.deleteBranch(this.#docId, name);
  }

  /**
   * Checks out the specified branch and returns its current state.
   *
   * @param {Actor} actor
   * @param {{ name: string }} input
   * @returns {Promise<DocumentState>}
   */
  async checkoutBranch(actor, { name }) {
    assertCanManageBranches(actor, "checkoutBranch");
    const branch = await this.#store.getBranch(this.#docId, name);
    if (!branch) throw new Error(`Branch not found: ${name}`);
    await this.#setCurrentBranchName(name);
    return this.#store.getDocumentStateAtCommit(branch.headCommitId);
  }

  /**
   * Creates a new commit on the current branch and advances the branch head.
   *
   * @param {Actor} actor
   * @param {{ nextState: DocumentState, message?: string }} input
   */
  async commit(actor, { nextState, message }) {
    assertCanCommit(actor);
    const branch = await this.getCurrentBranch();
    const currentState = normalizeDocumentState(
      await this.#store.getDocumentStateAtCommit(branch.headCommitId)
    );

    let effectiveNextState = nextState;
    if (isLegacyCellsOnlyState(nextState) || isSchemaV1CellsOnlyState(nextState)) {
      const legacy = normalizeDocumentState(nextState);

      // Legacy commits only provide per-sheet cell maps. Treat them as an overlay
      // on the current branch head so older callers cannot accidentally delete
      // workbook metadata or unrelated sheets.
      const merged = structuredClone(currentState);
      for (const [sheetId, cellMap] of Object.entries(legacy.cells ?? {})) {
        merged.cells[sheetId] = structuredClone(cellMap);
        if (!merged.sheets.metaById[sheetId]) {
          merged.sheets.metaById[sheetId] = { id: sheetId, name: sheetId };
        }
        if (!merged.sheets.order.includes(sheetId)) merged.sheets.order.push(sheetId);
      }
      effectiveNextState = merged;
    } else if (isRecord(nextState) && nextState.schemaVersion === 1) {
      // Legacy/partial schemaVersion=1 callers: preserve any workbook-level maps
      // they omit to avoid unintentionally deleting unknown keys.
      const merged = normalizeDocumentState(nextState);

      let didOverlay = false;

      if (shouldPreserveSchemaV1WorkbookMap(nextState, "metadata")) {
        merged.metadata = structuredClone(currentState.metadata ?? {});
        didOverlay = true;
      }
      if (shouldPreserveSchemaV1WorkbookMap(nextState, "namedRanges")) {
        merged.namedRanges = structuredClone(currentState.namedRanges ?? {});
        didOverlay = true;
      }
      if (shouldPreserveSchemaV1WorkbookMap(nextState, "comments")) {
        merged.comments = structuredClone(currentState.comments ?? {});
        didOverlay = true;
      }

      // Preserve sheet-level view state (e.g. frozen panes) when the caller omits it.
      for (const sheetId of Object.keys(merged.sheets.metaById ?? {})) {
        if (!shouldPreserveSchemaV1SheetView(nextState, sheetId)) continue;
        const currentView = currentState.sheets.metaById?.[sheetId]?.view;
        if (currentView !== undefined) {
          merged.sheets.metaById[sheetId].view = structuredClone(currentView);
          didOverlay = true;
        }
      }

      // Preserve sheet metadata (visibility/tabColor) when the caller omits it.
      for (const sheetId of Object.keys(merged.sheets.metaById ?? {})) {
        if (!shouldPreserveSchemaV1SheetVisibility(nextState, sheetId)) continue;
        const currentVisibility = currentState.sheets.metaById?.[sheetId]?.visibility;
        if (currentVisibility !== undefined) {
          merged.sheets.metaById[sheetId].visibility = currentVisibility;
          didOverlay = true;
        }
      }
      for (const sheetId of Object.keys(merged.sheets.metaById ?? {})) {
        if (!shouldPreserveSchemaV1SheetTabColor(nextState, sheetId)) continue;
        const currentTabColor = currentState.sheets.metaById?.[sheetId]?.tabColor;
        if (currentTabColor !== undefined) {
          merged.sheets.metaById[sheetId].tabColor = currentTabColor;
          didOverlay = true;
        }
      }

      // Preserve axis size overrides when the caller omits them from the view payload.
      // (e.g. older schemaVersion=1 clients that know about frozen panes but not row/col sizing).
      for (const sheetId of Object.keys(merged.sheets.metaById ?? {})) {
        const sheets = nextState.sheets;
        if (!isRecord(sheets) || !isRecord(sheets.metaById)) continue;
        const rawMeta = sheets.metaById[sheetId];
        if (!isRecord(rawMeta)) continue;
        const rawView = rawMeta.view;
        if (!isRecord(rawView)) continue;

        const currentView = currentState.sheets.metaById?.[sheetId]?.view;
        if (!isRecord(currentView)) continue;

        const mergedView = merged.sheets.metaById[sheetId].view;
        if (!isRecord(mergedView)) continue;

        if (!("colWidths" in rawView) && currentView.colWidths !== undefined) {
          mergedView.colWidths = structuredClone(currentView.colWidths);
          didOverlay = true;
        }

        if (
          !("backgroundImageId" in rawView) &&
          !("background_image_id" in rawView) &&
          !("backgroundImage" in rawView) &&
          !("background_image" in rawView) &&
          currentView.backgroundImageId !== undefined
        ) {
          mergedView.backgroundImageId = structuredClone(currentView.backgroundImageId);
          didOverlay = true;
        }

        if (!("rowHeights" in rawView) && currentView.rowHeights !== undefined) {
          mergedView.rowHeights = structuredClone(currentView.rowHeights);
          didOverlay = true;
        }

        if (
          !("mergedRanges" in rawView) &&
          !("mergedCells" in rawView) &&
          !("merged_cells" in rawView) &&
          !("merged_ranges" in rawView) &&
          !("mergedRegions" in rawView) &&
          !("merged_regions" in rawView) &&
          currentView.mergedRanges !== undefined
        ) {
          mergedView.mergedRanges = structuredClone(currentView.mergedRanges);
          didOverlay = true;
        }

        if (!("drawings" in rawView) && currentView.drawings !== undefined) {
          mergedView.drawings = structuredClone(currentView.drawings);
          didOverlay = true;
        }

        // Preserve range-run formatting when the caller omits it from the view payload.
        // (e.g. older schemaVersion=1 clients that know about frozen panes but not compressed
        // range formatting).
        if (!("formatRunsByCol" in rawView) && currentView.formatRunsByCol !== undefined) {
          mergedView.formatRunsByCol = structuredClone(currentView.formatRunsByCol);
          didOverlay = true;
        }
      }

      if (didOverlay) effectiveNextState = merged;
    }

    const patch = diffDocumentStates(currentState, effectiveNextState);
    const commit = await this.#store.createCommit({
      docId: this.#docId,
      parentCommitId: branch.headCommitId,
      mergeParentCommitId: null,
      createdBy: actor.userId,
      createdAt: Date.now(),
      message: message ?? null,
      patch,
      nextState: effectiveNextState,
    });
    await this.#store.updateBranchHead(this.#docId, branch.name, commit.id);
    return commit;
  }

  /**
   * @param {Actor} actor
   * @param {{ sourceBranch: string }} input
   * @returns {Promise<MergeResult & { baseCommitId: string, oursHeadCommitId: string, theirsHeadCommitId: string }>}
   */
  async previewMerge(actor, { sourceBranch }) {
    assertCanManageBranches(actor, "previewMerge");
    const oursBranch = await this.getCurrentBranch();
    const theirsBranch = await this.#store.getBranch(this.#docId, sourceBranch);
    if (!theirsBranch) throw new Error(`Branch not found: ${sourceBranch}`);

    const baseCommitId = await this.#findCommonAncestorCommitId(
      oursBranch.headCommitId,
      theirsBranch.headCommitId
    );

    const [baseState, oursState, theirsState] = await Promise.all([
      this.#store.getDocumentStateAtCommit(baseCommitId),
      this.#store.getDocumentStateAtCommit(oursBranch.headCommitId),
      this.#store.getDocumentStateAtCommit(theirsBranch.headCommitId)
    ]);

    const merge = mergeDocumentStates({ base: baseState, ours: oursState, theirs: theirsState });
    return {
      ...merge,
      baseCommitId,
      oursHeadCommitId: oursBranch.headCommitId,
      theirsHeadCommitId: theirsBranch.headCommitId
    };
  }

  /**
   * Applies a merge (with conflict resolutions) and creates a new merge commit on
   * the current branch.
   *
   * @param {Actor} actor
   * @param {{ sourceBranch: string, resolutions: ConflictResolution[], message?: string }} input
   * @returns {Promise<{ commit: Commit, state: DocumentState }>}
   */
  async merge(actor, { sourceBranch, resolutions, message }) {
    assertCanManageBranches(actor, "merge");
    const preview = await this.previewMerge(actor, { sourceBranch });
    const finalState = applyConflictResolutions(preview, resolutions);

    // Ensure all conflicts were addressed by validating that each conflict index
    // is present in `resolutions`.
    const resolved = new Set(resolutions.map((r) => r.conflictIndex));
    for (let i = 0; i < preview.conflicts.length; i += 1) {
      if (!resolved.has(i)) throw new Error("All merge conflicts must be resolved before merging");
    }

    const oursState = await this.#store.getDocumentStateAtCommit(preview.oursHeadCommitId);
    const patch = diffDocumentStates(oursState, finalState);

    const commit = await this.#store.createCommit({
      docId: this.#docId,
      parentCommitId: preview.oursHeadCommitId,
      mergeParentCommitId: preview.theirsHeadCommitId,
      createdBy: actor.userId,
      createdAt: Date.now(),
      message: message ?? `Merge branch '${sourceBranch}'`,
      patch,
      nextState: finalState,
    });

    const currentBranch = await this.getCurrentBranch();
    await this.#store.updateBranchHead(this.#docId, currentBranch.name, commit.id);

    return { commit, state: finalState };
  }

  /**
   * @param {string} oursHead
   * @param {string} theirsHead
   * @returns {Promise<string>}
   */
  async #findCommonAncestorCommitId(oursHead, theirsHead) {
    /** @type {Map<string, number>} */
    const oursDepth = new Map();
    const queue = [{ id: oursHead, depth: 0 }];
    while (queue.length > 0) {
      const { id, depth } = queue.shift();
      if (oursDepth.has(id)) continue;
      oursDepth.set(id, depth);
      const commit = await this.#store.getCommit(id);
      if (!commit) throw new Error(`Commit not found: ${id}`);
      if (commit.parentCommitId) queue.push({ id: commit.parentCommitId, depth: depth + 1 });
      if (commit.mergeParentCommitId)
        queue.push({ id: commit.mergeParentCommitId, depth: depth + 1 });
    }

    let best = null;
    let bestScore = Infinity;
    const queue2 = [{ id: theirsHead, depth: 0 }];
    /** @type {Set<string>} */
    const seen = new Set();
    while (queue2.length > 0) {
      const { id, depth } = queue2.shift();
      if (seen.has(id)) continue;
      seen.add(id);
      const ours = oursDepth.get(id);
      if (ours !== undefined) {
        const score = ours + depth;
        if (score < bestScore) {
          best = id;
          bestScore = score;
        }
      }
      const commit = await this.#store.getCommit(id);
      if (!commit) throw new Error(`Commit not found: ${id}`);
      if (commit.parentCommitId) queue2.push({ id: commit.parentCommitId, depth: depth + 1 });
      if (commit.mergeParentCommitId)
        queue2.push({ id: commit.mergeParentCommitId, depth: depth + 1 });
    }

    if (!best) throw new Error("No common ancestor found (corrupt history)");
    return best;
  }
}
