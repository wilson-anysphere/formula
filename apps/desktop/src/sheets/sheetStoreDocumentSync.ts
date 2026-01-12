import type { WorkbookSheetStore } from "./workbookSheetStore";

type DocumentControllerLike = {
  getSheetIds(): string[];
  on(event: string, listener: (payload: any) => void): () => void;
};

type SyncHandle = {
  /**
   * Force an immediate sync (no microtask debounce).
   */
  syncNow(): void;
  dispose(): void;
};

/**
 * Keep a `WorkbookSheetStore` in sync with the set of sheet ids present in the DocumentController.
 *
 * Important: this sync is one-way (DocumentController -> store). Store mutations should not
 * produce DocumentController changes unless explicitly routed elsewhere.
 */
export function startSheetStoreDocumentSync(
  doc: DocumentControllerLike,
  store: WorkbookSheetStore,
  getActiveSheetId: () => string,
  onActivateSheet: (sheetId: string) => void,
): SyncHandle {
  let scheduled = false;
  let disposed = false;
  let lastChangeSource: string | null = null;

  const schedule = () => {
    if (disposed) return;
    if (scheduled) return;
    scheduled = true;
    queueMicrotask(() => {
      scheduled = false;
      if (disposed) return;
      syncNow();
    });
  };

  const syncNow = () => {
    if (disposed) return;
    const activeSheetId = getActiveSheetId();

    const rawDocSheetIds = doc.getSheetIds();
    const docSheetIds = rawDocSheetIds.length > 0 ? rawDocSheetIds : activeSheetId ? [activeSheetId] : [];
    const docSet = new Set(docSheetIds);

    const existing = store.listAll();
    const existingIdSet = new Set(existing.map((s) => s.id));

    // Add any sheets that the doc created lazily (e.g. `setCellValue("Sheet2", ...)`).
    //
    // We insert after the active sheet when possible so newly-created sheets appear nearby.
    // When multiple sheets are missing, keep their relative order stable by inserting each
    // subsequent sheet after the one we just added.
    let insertAfterId =
      (activeSheetId && existingIdSet.has(activeSheetId) ? activeSheetId : undefined) ?? existing.at(-1)?.id ?? "";

    for (const sheetId of docSheetIds) {
      if (existingIdSet.has(sheetId)) continue;
      try {
        store.addAfter(insertAfterId, { id: sheetId, name: sheetId });
      } catch {
        // If the sheet id isn't a valid sheet name (or violates uniqueness constraints),
        // fall back to the store's default name generation.
        try {
          store.addAfter(insertAfterId, { id: sheetId });
        } catch {
          // If the store still can't accept the sheet (e.g. invalid id), skip it.
          continue;
        }
      }
      existingIdSet.add(sheetId);
      insertAfterId = sheetId;
    }

    // Remove metadata for sheets that no longer exist in the doc (e.g. `applyState` removed them).
    for (const sheet of existing) {
      if (docSet.has(sheet.id)) continue;
      try {
        store.remove(sheet.id);
      } catch {
        // Best-effort: avoid crashing the UI if the store refuses removal (e.g. last-sheet guard).
      }
    }

    // When restoring document state (VersionManager restore, workbook open, etc), prefer the
    // sheet ordering from the DocumentController snapshot so sheet tabs reflect the restored
    // workbook navigation order.
    if (lastChangeSource === "applyState") {
      const desiredOrder = docSheetIds.slice();
      const current = store.listAll().map((s) => s.id);
      const currentSet = new Set(current);
      const desired = desiredOrder.filter((id) => currentSet.has(id));
      // Append any store-only ids (should be rare; e.g. removal guarded).
      for (const id of current) {
        if (!desired.includes(id)) desired.push(id);
      }

      for (let targetIndex = 0; targetIndex < desired.length; targetIndex += 1) {
        const sheetId = desired[targetIndex]!;
        const currentIndex = store.listAll().findIndex((s) => s.id === sheetId);
        if (currentIndex === -1) continue;
        if (currentIndex === targetIndex) continue;
        try {
          store.move(sheetId, targetIndex);
        } catch {
          // Best-effort: if the store rejects a move (shouldn't happen), continue syncing the rest.
        }
      }
    }

    // If the active sheet is no longer valid, fall back to the first visible sheet.
    if (activeSheetId && !docSet.has(activeSheetId)) {
      const firstVisible = store.listVisible()[0] ?? store.listAll()[0] ?? null;
      if (firstVisible && firstVisible.id !== activeSheetId) {
        onActivateSheet(firstVisible.id);
      }
    }
  };

  const unsubscribeChange = doc.on("change", (payload) => {
    lastChangeSource = typeof (payload as any)?.source === "string" ? (payload as any).source : null;
    schedule();
  });
  const unsubscribeUpdate = doc.on("update", schedule);

  // Run a first-pass sync on startup.
  schedule();

  return {
    syncNow,
    dispose() {
      disposed = true;
      unsubscribeChange();
      unsubscribeUpdate();
    },
  };
}
