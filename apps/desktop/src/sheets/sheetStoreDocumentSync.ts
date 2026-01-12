import type { WorkbookSheetStore } from "./workbookSheetStore";

type DocumentControllerLike = {
  getSheetIds(): string[];
  /**
   * Optional sheet metadata accessor (DocumentController now supports this, but we keep the
   * sync layer tolerant so it can work with simple mocks/tests).
   */
  getSheetMeta?: (
    sheetId: string,
  ) => { name: string; visibility: "visible" | "hidden" | "veryHidden"; tabColor?: any } | null;
  on(event: string, listener: (payload: any) => void): () => void;
};

type SyncHandle = {
  /**
   * Force an immediate sync (no microtask debounce).
   */
  syncNow(): void;
  dispose(): void;
};

type SyncOptions = {
  /**
   * Optional wrapper invoked around store mutations (add/remove/move/rename/hide/tabColor).
   *
   * The desktop shell uses this to temporarily suppress "store -> doc" syncing while applying
   * authoritative DocumentController updates (e.g. undo/redo) into the UI store.
   */
  withStoreMutations?: <T>(fn: () => T) => T;
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
  options: SyncOptions = {},
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

    const withStoreMutations =
      typeof options.withStoreMutations === "function" ? options.withStoreMutations : <T>(fn: () => T) => fn();

    // Add any sheets that the doc created lazily (e.g. `setCellValue("Sheet2", ...)`).
    //
    // We insert after the active sheet when possible so newly-created sheets appear nearby.
    // When multiple sheets are missing, keep their relative order stable by inserting each
    // subsequent sheet after the one we just added.
    let insertAfterId =
      (activeSheetId && existingIdSet.has(activeSheetId) ? activeSheetId : undefined) ?? existing.at(-1)?.id ?? "";

    withStoreMutations(() => {
      for (const sheetId of docSheetIds) {
        if (existingIdSet.has(sheetId)) continue;
        const meta =
          typeof doc.getSheetMeta === "function"
            ? doc.getSheetMeta(sheetId) ?? { name: sheetId, visibility: "visible" as const }
            : { name: sheetId, visibility: "visible" as const };
        try {
          store.addAfter(insertAfterId, { id: sheetId, name: meta.name || sheetId });
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
    });

    // Remove metadata for sheets that no longer exist in the doc (e.g. `applyState` removed them).
    withStoreMutations(() => {
      for (const sheet of existing) {
        if (docSet.has(sheet.id)) continue;
        try {
          store.remove(sheet.id);
        } catch {
          // Best-effort: avoid crashing the UI if the store refuses removal (e.g. last-sheet guard).
        }
      }
    });

    // Keep the store order aligned with DocumentController order when the doc is the
    // authoritative source of truth for ordering:
    // - applyState restores (workbook open / version restore)
    // - undo/redo of sheet reorders
    //
    // Avoid forcing the doc ordering for unrelated change sources (e.g. endBatch),
    // since in normal editing flows the UI store is the canonical sheet tab ordering.
    const shouldSyncOrder =
      lastChangeSource === "applyState" || lastChangeSource === "undo" || lastChangeSource === "redo";
    if (shouldSyncOrder) {
      const desiredOrder = docSheetIds.slice();
      const current = store.listAll().map((s) => s.id);
      const currentSet = new Set(current);
      const desired = desiredOrder.filter((id) => currentSet.has(id));
      // Append any store-only ids (should be rare; e.g. removal guarded).
      for (const id of current) {
        if (!desired.includes(id)) desired.push(id);
      }

      withStoreMutations(() => {
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
      });
    }

    // Sync sheet metadata (name/visibility/tabColor) into the UI store so undo/redo of sheet
    // metadata operations updates the tab strip and sheet switcher.
    if (typeof doc.getSheetMeta === "function") {
      withStoreMutations(() => {
        for (const sheetId of docSheetIds) {
          const meta = doc.getSheetMeta?.(sheetId);
          if (!meta) continue;
          const storeMeta = store.getById(sheetId);
          if (!storeMeta) continue;

          // Name.
          if (storeMeta.name !== meta.name) {
            try {
              store.rename(sheetId, meta.name);
            } catch {
              // Best-effort: if the store rejects the name (should be rare since doc validated),
              // leave the existing UI name.
            }
          }

          // Visibility (treat veryHidden as hidden in the UI store for now).
          const desiredVisibility = meta.visibility === "visible" ? "visible" : "hidden";
          if (desiredVisibility === "visible" && storeMeta.visibility !== "visible") {
            try {
              store.unhide(sheetId);
            } catch {
              // ignore
            }
          } else if (desiredVisibility !== "visible" && storeMeta.visibility === "visible") {
            try {
              store.hide(sheetId);
            } catch {
              // ignore
            }
          }

          // Tab color.
          const a = storeMeta.tabColor ?? null;
          const b = meta.tabColor ?? null;
          const sameColor = JSON.stringify(a) === JSON.stringify(b);
          if (!sameColor) {
            try {
              store.setTabColor(sheetId, meta.tabColor ?? undefined);
            } catch {
              // ignore
            }
          }
        }
      });
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
