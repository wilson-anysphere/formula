import React from "react";
import { createRoot } from "react-dom/client";

import { markKeybindingBarrier } from "../keybindingBarrier.js";
import { normalizeExcelColorToCss } from "../shared/colors.js";

import { rewriteDocumentFormulasForSheetDelete } from "./sheetFormulaRewrite";
import type { SheetMeta, WorkbookSheetStore } from "./workbookSheetStore";

export type OrganizeSheetsDialogHost = {
  store: WorkbookSheetStore;
  /**
   * Optional: return the current authoritative sheet store.
   *
   * In collaboration mode `main.ts` can rebuild the sheet store instance when remote
   * metadata changes. Providing this hook lets the dialog re-bind to the latest store
   * instead of operating on a stale instance.
   */
  getStore?: () => WorkbookSheetStore;
  /**
   * Read the current active sheet id.
   *
   * Used to:
   * - render an "Active" state
   * - decide whether delete/hide should activate a fallback sheet
   */
  getActiveSheetId: () => string;
  /**
   * Activate/select a sheet (usually `app.activateSheet`).
   */
  activateSheet: (sheetId: string) => void;
  /**
   * Rename a sheet by id, including formula rewrites.
   *
   * In the desktop shell, this should be `renameSheetById(...)` from `main.ts`.
   */
  renameSheetById: (sheetId: string, newName: string) => Promise<unknown> | unknown;
  /**
   * The current document controller (needed for delete formula rewrites).
   */
  getDocument: () => any;
  /**
   * Used to disable operations while the spreadsheet is actively editing.
   */
  isEditing: () => boolean;
  /**
   * When true, disable sheet-structure mutations (rename/hide/delete/reorder/unhide).
   *
   * Used in read-only collaboration sessions so the dialog behaves consistently with
   * the sheet tab strip (which disables these controls instead of error-toasting).
   */
  readOnly?: boolean;
  /**
   * Called after the dialog closes (e.g. restore grid focus).
   */
  focusGrid: () => void;
  /**
   * Optional error surface (main.ts wires this to `showToast(..., "error")`).
   */
  onError?: (message: string) => void;
};

type OrganizeSheetsDialogProps = {
  host: OrganizeSheetsDialogHost;
  onClose: () => void;
};

function visibilityBadgeLabel(visibility: SheetMeta["visibility"]): string | null {
  if (visibility === "hidden") return "Hidden";
  if (visibility === "veryHidden") return "Very Hidden";
  return null;
}

function OrganizeSheetsDialog({ host, onClose }: OrganizeSheetsDialogProps) {
  const { store } = host;
  const [sheets, setSheets] = React.useState<SheetMeta[]>(() => store.listAll());
  const [activeSheetId, setActiveSheetId] = React.useState(() => host.getActiveSheetId());
  const [renameSheetId, setRenameSheetId] = React.useState<string | null>(null);
  const [renameDraft, setRenameDraft] = React.useState("");
  const renameDraftRef = React.useRef("");
  const [deleteConfirmSheetId, setDeleteConfirmSheetId] = React.useState<string | null>(null);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    setSheets(store.listAll());
    return store.subscribe(() => {
      setSheets(store.listAll());
    });
  }, [store]);

  // Keep the highlighted "Active" marker in sync with the host's current active sheet.
  // This matters if the active sheet changes while the dialog is open (e.g. remote delete triggers
  // a fallback activation in collab mode, or programmatic sheet navigation).
  React.useEffect(() => {
    try {
      setActiveSheetId(host.getActiveSheetId());
    } catch {
      // ignore
    }
  }, [host, sheets]);

  React.useEffect(() => {
    if (typeof window === "undefined") return;

    const onSheetActivated = (evt: Event) => {
      const detail = (evt as CustomEvent).detail as any;
      const next = typeof detail?.sheetId === "string" ? String(detail.sheetId).trim() : "";
      if (next) {
        setActiveSheetId(next);
        return;
      }
      try {
        setActiveSheetId(host.getActiveSheetId());
      } catch {
        // ignore
      }
    };

    window.addEventListener("formula:sheet-activated", onSheetActivated);
    return () => window.removeEventListener("formula:sheet-activated", onSheetActivated);
  }, [host]);

  // If the underlying store is replaced (collab) or a sheet is removed remotely while the
  // dialog is open, clear any inline UI state that references non-existent sheets so the
  // dialog doesn't get stuck in a "disabled" state.
  React.useEffect(() => {
    if (renameSheetId && !sheets.some((s) => s.id === renameSheetId)) {
      setRenameSheetId(null);
      setRenameDraft("");
      renameDraftRef.current = "";
    }
    if (deleteConfirmSheetId && !sheets.some((s) => s.id === deleteConfirmSheetId)) {
      setDeleteConfirmSheetId(null);
    }
  }, [deleteConfirmSheetId, renameSheetId, sheets]);

  const visibleCount = React.useMemo(() => sheets.filter((s) => s.visibility === "visible").length, [sheets]);

  const reportError = React.useCallback(
    (err: unknown) => {
      const message = err instanceof Error ? err.message : String(err);
      setError(message);
      try {
        host.onError?.(message);
      } catch {
        // ignore
      }
    },
    [host],
  );

  const isSpreadsheetEditing = host.isEditing();
  const isReadOnly = host.readOnly === true;

  const beginRename = React.useCallback(
    (sheet: SheetMeta) => {
      if (isSpreadsheetEditing || isReadOnly) return;
      setError(null);
      setDeleteConfirmSheetId(null);
      setRenameSheetId(sheet.id);
      renameDraftRef.current = sheet.name;
      setRenameDraft(sheet.name);
    },
    [isReadOnly, isSpreadsheetEditing],
  );

  const commitRename = React.useCallback(async () => {
    const sheetId = renameSheetId;
    if (!sheetId) return;
    if (isSpreadsheetEditing || isReadOnly) return;

    setBusy(true);
    setError(null);
    try {
      // Read from a ref to ensure we commit the latest input value even when React
      // batches state updates across multiple events (input + Enter in the same tick).
      await host.renameSheetById(sheetId, renameDraftRef.current);
      setRenameSheetId(null);
    } catch (err) {
      reportError(err);
    } finally {
      setBusy(false);
    }
  }, [host, isReadOnly, isSpreadsheetEditing, renameSheetId, reportError]);

  const cancelRename = React.useCallback(() => {
    setRenameSheetId(null);
    setError(null);
  }, []);

  const cancelDeleteConfirm = React.useCallback(() => {
    setDeleteConfirmSheetId(null);
    setError(null);
  }, []);

  const activate = React.useCallback(
    (sheetId: string) => {
      if (isSpreadsheetEditing) return;
      try {
        const meta = store.getById(sheetId);
        // Keep the spreadsheet in a consistent state: hidden/veryHidden sheets should not become
        // the active sheet. If the user chooses to activate one, unhide it first.
        if (meta && meta.visibility !== "visible") {
          if (isReadOnly) return;
          store.unhide(sheetId);
        }
        host.activateSheet(sheetId);
        setActiveSheetId(sheetId);
      } catch (err) {
        reportError(err);
      }
    },
    [host, isReadOnly, isSpreadsheetEditing, reportError, store],
  );

  const hide = React.useCallback(
    (sheet: SheetMeta) => {
      if (isSpreadsheetEditing || isReadOnly) return;
      setError(null);
      const currentActive = host.getActiveSheetId();
      const wasActive = sheet.id === currentActive;

      let nextActiveId: string | null = null;
      if (wasActive) {
        const visibleSheets = store.listAll().filter((s) => s.visibility === "visible");
        const idx = visibleSheets.findIndex((s) => s.id === sheet.id);
        nextActiveId = idx === -1 ? null : (visibleSheets[idx + 1]?.id ?? visibleSheets[idx - 1]?.id ?? null);
      }

      try {
        store.hide(sheet.id);
      } catch (err) {
        reportError(err);
        return;
      }

      if (wasActive) {
        const fallback = store.listVisible().at(0)?.id ?? store.listAll().at(0)?.id ?? null;
        const next = nextActiveId ?? fallback;
        if (next && next !== sheet.id) {
          activate(next);
        }
      }
    },
    [activate, host, isReadOnly, isSpreadsheetEditing, reportError, store],
  );

  const unhide = React.useCallback(
    (sheet: SheetMeta) => {
      if (isSpreadsheetEditing || isReadOnly) return;
      setError(null);
      try {
        store.unhide(sheet.id);
      } catch (err) {
        reportError(err);
      }
    },
    [isReadOnly, isSpreadsheetEditing, reportError, store],
  );

  const remove = React.useCallback(
    async (sheet: SheetMeta) => {
      if (isSpreadsheetEditing || isReadOnly) return;
      setError(null);
      setDeleteConfirmSheetId(null);

      const deletedName = sheet.name;
      // Use the pre-delete sheet ordering (by name) so 3D refs like `Sheet1:Sheet3!A1`
      // can shift boundaries correctly when an endpoint is deleted (Excel-like).
      const allSheets = store.listAll();
      const sheetOrder = allSheets.map((s) => s.name);
      // Mirror Excel: prevent deleting the last visible sheet (even if hidden sheets remain).
      const visibleCount = allSheets.filter((s) => s.visibility === "visible").length;
      const sheetMeta = allSheets.find((s) => s.id === sheet.id) ?? sheet;
      if (sheetMeta.visibility === "visible" && visibleCount <= 1) {
        reportError(new Error("Cannot delete the last visible sheet"));
        return;
      }

      const currentActive = host.getActiveSheetId();
      const wasActive = sheet.id === currentActive;

      try {
        store.remove(sheet.id);
      } catch (err) {
        reportError(err);
        return;
      }

      if (wasActive) {
        const next = store.listVisible().at(0)?.id ?? store.listAll().at(0)?.id ?? null;
        if (next && next !== sheet.id) {
          activate(next);
        }
      }

      try {
        const doc = host.getDocument();
        rewriteDocumentFormulasForSheetDelete(doc, deletedName, sheetOrder);
      } catch (err) {
        reportError(err);
      }
    },
    [activate, host, isReadOnly, isSpreadsheetEditing, reportError, store],
  );

  const move = React.useCallback(
    (sheetId: string, toIndex: number) => {
      if (isSpreadsheetEditing || isReadOnly) return;
      setError(null);
      try {
        store.move(sheetId, toIndex);
      } catch (err) {
        reportError(err);
      }
    },
    [isReadOnly, isSpreadsheetEditing, reportError, store],
  );

  return (
    <div
      className="organize-sheets-dialog__body"
      data-testid="organize-sheets-dialog-body"
      onKeyDown={(e) => {
        if (e.key !== "Escape") return;
        // When an inline modal state is active (rename/delete confirm), Escape should dismiss that
        // state first (Excel-like) rather than closing the entire dialog.
        if (renameSheetId) {
          e.preventDefault();
          e.stopPropagation();
          cancelRename();
          return;
        }
        if (deleteConfirmSheetId) {
          e.preventDefault();
          e.stopPropagation();
          cancelDeleteConfirm();
          return;
        }
        e.preventDefault();
        e.stopPropagation();
        onClose();
      }}
    >
      <header className="organize-sheets-dialog__header">
        <h2 className="organize-sheets-dialog__title">Organize Sheets</h2>
        <button
          type="button"
          className="organize-sheets-dialog__close"
          onClick={onClose}
          data-testid="organize-sheets-close"
        >
          Close
        </button>
      </header>

      <div className="organize-sheets-dialog__list" role="list" data-testid="organize-sheets-list">
        {sheets.map((sheet, index) => {
          const isActive = sheet.id === activeSheetId;
          const badge = visibilityBadgeLabel(sheet.visibility);
          const tabColorCss = sheet.tabColor ? normalizeExcelColorToCss(sheet.tabColor) : undefined;
          const rowIsRenaming = renameSheetId === sheet.id;
          const rowIsConfirmingDelete = deleteConfirmSheetId === sheet.id;
          const hasAnyInlineModal = renameSheetId != null || deleteConfirmSheetId != null;
          const otherInlineModalActive =
            (renameSheetId != null && !rowIsRenaming) || (deleteConfirmSheetId != null && !rowIsConfirmingDelete);

          // Mirror Excel: you cannot delete the last *visible* sheet (even if hidden sheets remain).
          const canDelete = sheets.length > 1 && (sheet.visibility !== "visible" || visibleCount > 1);
          const canHide = sheet.visibility === "visible" && visibleCount > 1;
          const canUnhide = sheet.visibility !== "visible";
          const moveUpDisabled = index <= 0;
          const moveDownDisabled = index >= sheets.length - 1;

          const actionDisabled = busy || isSpreadsheetEditing || otherInlineModalActive;
          const confirmingDelete = rowIsConfirmingDelete;
          const activateLabel = sheet.visibility === "visible" ? "Activate" : "Unhide & Activate";
          const mutationDisabled = actionDisabled || isReadOnly || rowIsRenaming || rowIsConfirmingDelete;
          const canActivate = sheet.visibility === "visible" || !isReadOnly;
          const activateDisabled = actionDisabled || rowIsRenaming || rowIsConfirmingDelete || !canActivate || (hasAnyInlineModal && !rowIsRenaming && !rowIsConfirmingDelete);

          return (
            <div
              key={sheet.id}
              className="organize-sheets-dialog__row"
              role="listitem"
              data-testid={`organize-sheet-row-${sheet.id}`}
            >
              <div className="organize-sheets-dialog__name">
                {renameSheetId === sheet.id ? (
                  <input
                    type="text"
                    value={renameDraft}
                    onInput={(e) => {
                      const next = (e.target as HTMLInputElement).value;
                      renameDraftRef.current = next;
                      setRenameDraft(next);
                    }}
                    className="organize-sheets-dialog__rename-input"
                    data-testid={`organize-sheet-rename-input-${sheet.id}`}
                    autoFocus
                    onKeyDown={(e) => {
                      if (e.key === "Enter") {
                        e.preventDefault();
                        void commitRename();
                      }
                    }}
                    disabled={busy || isSpreadsheetEditing || isReadOnly}
                  />
                ) : (
                  <>
                    {tabColorCss ? (
                      <span
                        className="organize-sheets-dialog__tab-color"
                        style={{ backgroundColor: tabColorCss }}
                        aria-label="Tab color"
                        data-testid={`organize-sheet-tab-color-${sheet.id}`}
                      />
                    ) : null}
                    <span className="organize-sheets-dialog__name-text" data-testid={`organize-sheet-name-${sheet.id}`}>
                      {sheet.name}
                    </span>
                    {badge ? (
                      <span
                        className="organize-sheets-dialog__badge"
                        data-testid={`organize-sheet-visibility-${sheet.id}`}
                        aria-label={badge}
                      >
                        {badge}
                      </span>
                    ) : null}
                    {isActive ? (
                      <span className="organize-sheets-dialog__active" data-testid={`organize-sheet-active-${sheet.id}`}>
                        Active
                      </span>
                    ) : null}
                  </>
                )}
              </div>

              <div className="organize-sheets-dialog__actions">
                <button
                  type="button"
                  onClick={() => activate(sheet.id)}
                  disabled={activateDisabled}
                  data-testid={`organize-sheet-activate-${sheet.id}`}
                >
                  {activateLabel}
                </button>

                {renameSheetId === sheet.id ? (
                  <>
                    <button
                      type="button"
                      onClick={() => void commitRename()}
                      disabled={busy || isSpreadsheetEditing || isReadOnly}
                      data-testid={`organize-sheet-rename-save-${sheet.id}`}
                    >
                      Save
                    </button>
                    <button
                      type="button"
                      onClick={cancelRename}
                      disabled={busy}
                      data-testid={`organize-sheet-rename-cancel-${sheet.id}`}
                    >
                      Cancel
                    </button>
                  </>
                ) : (
                  <button
                    type="button"
                    onClick={() => beginRename(sheet)}
                    disabled={mutationDisabled}
                    data-testid={`organize-sheet-rename-${sheet.id}`}
                  >
                    Rename
                  </button>
                )}

                {sheet.visibility === "visible" ? (
                  <button
                    type="button"
                    onClick={() => hide(sheet)}
                    disabled={mutationDisabled || !canHide}
                    data-testid={`organize-sheet-hide-${sheet.id}`}
                  >
                    Hide
                  </button>
                ) : (
                  <button
                    type="button"
                    onClick={() => unhide(sheet)}
                    disabled={mutationDisabled || !canUnhide}
                    data-testid={`organize-sheet-unhide-${sheet.id}`}
                  >
                    Unhide
                  </button>
                )}

                {confirmingDelete ? (
                  <>
                    <button
                      type="button"
                      onClick={() => void remove(sheet)}
                      disabled={busy || isSpreadsheetEditing || isReadOnly || !canDelete}
                      data-testid={`organize-sheet-delete-confirm-${sheet.id}`}
                    >
                      Confirm Delete
                    </button>
                    <button
                      type="button"
                      onClick={cancelDeleteConfirm}
                      disabled={actionDisabled}
                      data-testid={`organize-sheet-delete-cancel-${sheet.id}`}
                    >
                      Cancel
                    </button>
                  </>
                ) : (
                  <button
                    type="button"
                    onClick={() => {
                      if (mutationDisabled || !canDelete) return;
                      setDeleteConfirmSheetId(sheet.id);
                      setError(null);
                    }}
                    disabled={mutationDisabled || !canDelete}
                    data-testid={`organize-sheet-delete-${sheet.id}`}
                  >
                    Delete
                  </button>
                )}

                <button
                  type="button"
                  onClick={() => move(sheet.id, index - 1)}
                  disabled={mutationDisabled || moveUpDisabled}
                  data-testid={`organize-sheet-move-up-${sheet.id}`}
                >
                  Up
                </button>
                <button
                  type="button"
                  onClick={() => move(sheet.id, index + 1)}
                  disabled={mutationDisabled || moveDownDisabled}
                  data-testid={`organize-sheet-move-down-${sheet.id}`}
                >
                  Down
                </button>
              </div>
            </div>
          );
        })}
      </div>

      {error ? (
        <div className="organize-sheets-dialog__error" role="alert" data-testid="organize-sheets-error">
          {error}
        </div>
      ) : null}
    </div>
  );
}

function trapDialogTabFocus(dialog: HTMLDialogElement): void {
  dialog.addEventListener("keydown", (event) => {
    if (event.key !== "Tab") return;
    const focusables = Array.from(
      dialog.querySelectorAll<HTMLElement>(
        'button:not(:disabled), [href], input:not(:disabled), select:not(:disabled), textarea:not(:disabled), [tabindex]:not([tabindex="-1"])',
      ),
    ).filter((el) => el.getAttribute("aria-hidden") !== "true");
    if (focusables.length === 0) return;
    const first = focusables[0]!;
    const last = focusables[focusables.length - 1]!;
    const active = document.activeElement as HTMLElement | null;
    if (!active) return;

    if (event.shiftKey) {
      if (active === first) {
        event.preventDefault();
        last.focus();
      }
      return;
    }

    if (active === last) {
      event.preventDefault();
      first.focus();
    }
  });
}

function showDialogModal(dialog: HTMLDialogElement): void {
  const showModal = (dialog as any).showModal as (() => void) | undefined;
  if (typeof showModal === "function") {
    try {
      showModal.call(dialog);
      return;
    } catch {
      // Fall through to non-modal open attribute.
    }
  }

  // jsdom doesn't implement showModal(). Best-effort fallback so unit tests can
  // still exercise the dialog contents.
  dialog.setAttribute("open", "");
}

function closeDialog(dialog: HTMLDialogElement): void {
  const close = (dialog as any).close as ((returnValue?: string) => void) | undefined;
  if (typeof close === "function") {
    close.call(dialog);
    return;
  }

  // Fallback for jsdom environments without `HTMLDialogElement.close`.
  dialog.removeAttribute("open");
  dialog.dispatchEvent(new Event("close"));
}

export function openOrganizeSheetsDialog(host: OrganizeSheetsDialogHost): void {
  if (host.isEditing()) return;

  // Avoid throwing when another modal dialog is already open.
  const openModal = document.querySelector("dialog[open]");
  if (openModal) {
    if (openModal.classList.contains("organize-sheets-dialog")) return;
    return;
  }

  const dialog = document.createElement("dialog");
  dialog.className = "dialog organize-sheets-dialog";
  dialog.dataset.testid = "organize-sheets-dialog";
  dialog.setAttribute("aria-label", "Organize Sheets");
  markKeybindingBarrier(dialog);

  const container = document.createElement("div");
  dialog.appendChild(container);
  document.body.appendChild(dialog);

  const root = createRoot(container);

  const close = () => closeDialog(dialog);

  function Wrapper() {
    const [store, setStore] = React.useState<WorkbookSheetStore>(() => host.store);

    // Ensure the dialog updates its store binding if the host swaps it (e.g. collab metadata update).
    React.useEffect(() => {
      setStore(host.store);
    }, [host.store]);

    React.useEffect(() => {
      if (typeof window === "undefined") return;
      if (typeof host.getStore !== "function") return;

      const updateStore = () => {
        try {
          const next = host.getStore?.();
          if (next && next !== store) {
            setStore(next);
          }
        } catch {
          // ignore
        }
      };

      // `main.ts` emits this on any sheet metadata change (including store replacement).
      window.addEventListener("formula:sheet-metadata-changed", updateStore);
      return () => window.removeEventListener("formula:sheet-metadata-changed", updateStore);
    }, [host.getStore, store]);

    const hostWithStore = React.useMemo(() => ({ ...host, store }), [store]);
    return React.createElement(OrganizeSheetsDialog, { host: hostWithStore, onClose: close });
  }

  root.render(React.createElement(Wrapper));

  dialog.addEventListener(
    "close",
    () => {
      try {
        root.unmount();
      } catch {
        // ignore
      }
      dialog.remove();
      try {
        host.focusGrid();
      } catch {
        // ignore
      }
    },
    { once: true },
  );

  dialog.addEventListener("cancel", (e) => {
    // Prevent the dialog from auto-closing on Escape so the React UI can decide what to do:
    // - when renaming / confirming delete, Escape should dismiss that inline state
    // - otherwise Escape should close the dialog (handled by the React keydown handler)
    e.preventDefault();
  });

  trapDialogTabFocus(dialog);
  showDialogModal(dialog);
}
