import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

import { ContextMenu, type ContextMenuItem } from "../menus/contextMenu.js";
import { READ_ONLY_SHEET_MUTATION_MESSAGE } from "../collab/permissionGuards";
import { normalizeExcelColorToCss } from "../shared/colors.js";
import { resolveCssVar } from "../theme/cssVars.js";
import { showQuickPick } from "../extensions/ui.js";
import * as nativeDialogs from "../tauri/nativeDialogs";
import type { SheetMeta, SheetVisibility, TabColor, WorkbookSheetStore } from "./workbookSheetStore";
import { computeWorkbookSheetMoveIndex } from "./sheetReorder";
import { pickAdjacentVisibleSheetId } from "./sheetNavigation";

type SheetTabPaletteEntry = {
  label: string;
  token: string;
  /**
   * OOXML / Excel-style ARGB (AARRGGBB) string stored in workbook metadata.
   *
   * This keeps the palette stable without hardcoding CSS hex colors in source (enforced by
   * `apps/desktop/test/noHardcodedColors.test.js`).
   */
  excelArgb: string;
};

const SHEET_TAB_COLOR_PALETTE: SheetTabPaletteEntry[] = [
  { label: "Red", token: "--sheet-tab-red", excelArgb: "FFFF0000" },
  { label: "Orange", token: "--sheet-tab-orange", excelArgb: "FFFF9900" },
  { label: "Yellow", token: "--sheet-tab-yellow", excelArgb: "FFFFFF00" },
  { label: "Green", token: "--sheet-tab-green", excelArgb: "FF00B050" },
  { label: "Teal", token: "--sheet-tab-teal", excelArgb: "FF00B0F0" },
  { label: "Blue", token: "--sheet-tab-blue", excelArgb: "FF0070C0" },
  { label: "Purple", token: "--sheet-tab-purple", excelArgb: "FF7030A0" },
  { label: "Gray", token: "--sheet-tab-gray", excelArgb: "FF7F7F7F" },
];

function excelArgbToCssHex(argb: string): string | null {
  const raw = String(argb ?? "").trim();
  if (!/^[0-9a-fA-F]{8}$/.test(raw)) return null;
  // Ignore alpha; CSS hex expects #RRGGBB.
  return `#${raw.slice(2).toLowerCase()}`;
}

function normalizeCssHexToExcelArgb(cssColor: string): string | null {
  const trimmed = String(cssColor ?? "").trim();
  if (!trimmed) return null;

  const hasHash = trimmed.startsWith("#");
  const raw = hasHash ? trimmed.slice(1) : trimmed;

  // Fast path: hex-like strings (including ARGB).
  if (/^[0-9a-fA-F]+$/.test(raw)) {
    // #RRGGBB / RRGGBB
    if (raw.length === 6) {
      return `FF${raw.toUpperCase()}`;
    }

    // #RGB / RGB (expand to RRGGBB)
    if (raw.length === 3) {
      const [r, g, b] = raw.toUpperCase().split("");
      if (!r || !g || !b) return null;
      return `FF${r}${r}${g}${g}${b}${b}`;
    }

    // Already ARGB (#AARRGGBB / AARRGGBB)
    if (raw.length === 8) {
      return raw.toUpperCase();
    }
  }

  // Fallback: resolve non-hex CSS colors (named colors, rgb()/rgba(), etc) through
  // the browser's CSS parser so we can still persist them as Excel ARGB. This is
  // especially important in test environments (jsdom) where CSS tokens are not
  // available and we fall back to named colors like "red".
  const resolved = (() => {
    if (typeof document === "undefined") return null;
    const root = document.body ?? document.documentElement;
    if (!root) return null;

    const probe = document.createElement("span");
    probe.style.color = trimmed;
    probe.style.display = "none";
    root.appendChild(probe);
    try {
      const computed = typeof getComputedStyle === "function" ? getComputedStyle(probe).color : "";
      const normalized = String(computed ?? "").trim();
      return normalized ? normalized : null;
    } catch {
      return null;
    } finally {
      probe.remove();
    }
  })();
  if (!resolved) return null;

  const parsed = (() => {
    const match = resolved.match(
      /^rgba?\(\s*([0-9]+)\s*,\s*([0-9]+)\s*,\s*([0-9]+)(?:\s*,\s*([0-9]*\.?[0-9]+))?\s*\)$/i,
    );
    if (!match) return null;
    const r = Number(match[1]);
    const g = Number(match[2]);
    const b = Number(match[3]);
    const a = match[4] != null ? Number(match[4]) : 1;
    if (![r, g, b, a].every((n) => Number.isFinite(n))) return null;
    return { r, g, b, a };
  })();
  if (!parsed) return null;

  const clampByte = (value: number): number => Math.max(0, Math.min(255, Math.round(value)));
  const toHex2 = (value: number): string => clampByte(value).toString(16).padStart(2, "0").toUpperCase();

  const alpha = Math.max(0, Math.min(1, parsed.a));
  return `${toHex2(alpha * 255)}${toHex2(parsed.r)}${toHex2(parsed.g)}${toHex2(parsed.b)}`;
}

type Props = {
  store: WorkbookSheetStore;
  activeSheetId: string;
  /**
   * When true, disable sheet-structure mutations (add/rename/move/delete/hide/tabColor).
   *
   * This is primarily used in collab sessions with viewer/commenter roles where mutations
   * would otherwise diverge from the authoritative shared state.
   */
  readOnly?: boolean;
  /**
   * When true, disable sheet-structure mutations while the user is actively editing a cell/formula.
   *
   * Sheet navigation remains enabled so formulas can still reference ranges across sheets (Excel-style).
   */
  disableMutations?: boolean;
  onActivateSheet: (sheetId: string) => void;
  onAddSheet: () => Promise<void> | void;
  /**
   * Sheet rename handler.
   *
   * The handler is expected to validate + persist the rename (backend + formula rewrite)
   * and update the passed sheet store on success.
   *
   * It should throw on failure so the tab strip can keep the input focused and surface
   * the error inline.
   */
  onRenameSheet: (sheetId: string, newName: string) => Promise<void> | void;
  /**
   * Optional hook to persist sheet deletion before updating the local sheet metadata store.
   *
   * Used by the desktop (Tauri) shell to route deletes through the backend so the workbook
   * stays consistent across reloads.
   */
  onPersistSheetDelete?: (sheetId: string) => Promise<void> | void;
  /**
   * Optional hook to persist sheet visibility changes before updating the local sheet metadata store.
   *
   * Used by the desktop (Tauri) shell to route sheet hide/unhide through the backend so the workbook
   * stays consistent across reloads.
   */
  onPersistSheetVisibility?: (sheetId: string, visibility: SheetVisibility) => Promise<void> | void;
  /**
   * Optional hook to persist sheet tab color changes before updating the local sheet metadata store.
   *
   * Used by the desktop (Tauri) shell to route tab color changes through the backend so the workbook
   * stays consistent across reloads.
   */
  onPersistSheetTabColor?: (sheetId: string, tabColor: TabColor | undefined) => Promise<void> | void;
  /**
   * Called after a sheet tab reorder is committed (drag-and-drop).
   *
   * Used by `main.ts` to restore focus back to the grid so users can continue
   * editing after reordering.
   */
  onSheetsReordered?: () => void;
  /**
   * Called after a sheet is successfully deleted from the metadata store.
   *
   * Used by `main.ts` to rewrite DocumentController formulas referencing the deleted sheet name.
   */
  onSheetDeleted?: (event: { sheetId: string; name: string; sheetOrder: string[] }) => void;
  /**
   * Optional hook invoked when a sheet tab reorder is committed (drag-and-drop).
   *
   * The desktop shell can use this to persist the new sheet order to the backend.
   * If this throws/rejects, the tab strip rolls back the reorder (best-effort).
   */
  onSheetMoved?: (event: { sheetId: string; toIndex: number }) => Promise<void> | void;
  /**
   * Optional toast/error surface (used by the desktop shell).
   */
  onError?: (message: string) => void;
};

export function SheetTabStrip({
  store,
  activeSheetId,
  readOnly = false,
  disableMutations = false,
  onActivateSheet,
  onAddSheet,
  onPersistSheetDelete,
  onPersistSheetVisibility,
  onPersistSheetTabColor,
  onSheetsReordered,
  onSheetDeleted,
  onSheetMoved,
  onRenameSheet,
  onError,
}: Props) {
  const mutationsDisabled = readOnly || disableMutations;
  const mutationsDisabledMessage = readOnly
    ? READ_ONLY_SHEET_MUTATION_MESSAGE
    : disableMutations
      ? "Finish editing to modify sheets."
      : null;
  const reportMutationsDisabled = (): void => {
    if (!mutationsDisabledMessage) return;
    onError?.(mutationsDisabledMessage);
  };

  const [sheets, setSheets] = useState<SheetMeta[]>(() => store.listAll());
  const activeSheetIdRef = useRef(activeSheetId);

  useEffect(() => {
    activeSheetIdRef.current = activeSheetId;
  }, [activeSheetId]);

  useEffect(() => {
    setSheets(store.listAll());
    return store.subscribe(() => {
      setSheets(store.listAll());
    });
  }, [store]);

  const visibleSheets = useMemo(() => {
    const visible = sheets.filter((s) => s.visibility === "visible");
    if (visible.length > 0) return visible;
    if (sheets.length === 0) return [];

    // Defensive: if workbook metadata is invalid (all sheets hidden/veryHidden), fall back
    // to showing exactly one sheet so the tab strip remains usable (and the user can unhide
    // sheets via the context menu). Prefer the active sheet when possible.
    const active = sheets.find((s) => s.id === activeSheetId) ?? null;
    const nonVeryHiddenActive = active && active.visibility !== "veryHidden" ? active : null;
    const firstNonVeryHidden = sheets.find((s) => s.visibility !== "veryHidden") ?? null;
    const fallback = nonVeryHiddenActive ?? firstNonVeryHidden ?? active ?? sheets[0]!;
    return fallback ? [fallback] : [];
  }, [activeSheetId, sheets]);
  const [draggingSheetId, setDraggingSheetId] = useState<string | null>(null);
  const [dropIndicator, setDropIndicator] = useState<{ targetSheetId: string; position: "before" | "after" } | null>(null);

  const containerRef = useRef<HTMLDivElement | null>(null);
  const autoScrollRef = useRef<{ raf: number | null; direction: -1 | 0 | 1 }>({ raf: null, direction: 0 });
  const activeTabRef = useRef<HTMLButtonElement | null>(null);

  const ensureActiveTabVisible = useCallback(() => {
    const container = containerRef.current;
    const tab = activeTabRef.current;
    if (!container || !tab) return;

    // Avoid `scrollIntoView` here: the tab strip opts into smooth scrolling via CSS (`scroll-behavior: smooth`),
    // and `scrollIntoView` can therefore schedule a smooth scroll that emits `scroll` events asynchronously.
    // That interacts poorly with context menus (they close on outside scroll events).
    const containerRect = container.getBoundingClientRect();
    const tabRect = tab.getBoundingClientRect();

    // Give ourselves a tiny buffer to avoid jitter/rounding issues.
    const margin = 2;
    const leftOverflow = containerRect.left + margin - tabRect.left;
    const rightOverflow = tabRect.right - (containerRect.right - margin);

    if (leftOverflow > 0) {
      container.scrollLeft -= leftOverflow;
      return;
    }
    if (rightOverflow > 0) {
      container.scrollLeft += rightOverflow;
    }
  }, []);

  const [editingSheetId, setEditingSheetId] = useState<string | null>(null);
  const editingSheetIdRef = useRef<string | null>(null);
  const [draftName, setDraftName] = useState("");
  const [renameError, setRenameError] = useState<string | null>(null);
  const [renameInFlight, setRenameInFlight] = useState(false);
  const renameInputRef = useRef<HTMLInputElement>(null!);
  const renameCommitRef = useRef<Promise<boolean> | null>(null);
  const moveCommitSeqRef = useRef(0);
  const [canScroll, setCanScroll] = useState<{ left: boolean; right: boolean }>({ left: false, right: false });
  const tabColorPickerRef = useRef<HTMLInputElement | null>(null);
  const tabColorPickerDefaultValueRef = useRef<string | null>(null);

  const lastContextMenuFocusRef = useRef<HTMLElement | null>(null);
  const pendingEnsureActiveTabVisibleRef = useRef(false);
  const tabContextMenu = useMemo(
    () =>
      new ContextMenu({
        testId: "sheet-tab-context-menu",
        onClose: () => {
          // Restore focus so keyboard users don't "fall off" the tab strip after dismissing the menu.
          const target = lastContextMenuFocusRef.current;
          if (target?.isConnected) {
            target.focus({ preventScroll: true });
          }

          // If we skipped keeping the active tab visible while the context menu was open (to avoid
          // self-triggered scroll events immediately closing the menu), perform it now.
          if (pendingEnsureActiveTabVisibleRef.current) {
            pendingEnsureActiveTabVisibleRef.current = false;
            requestAnimationFrame(() => {
              if (editingSheetIdRef.current) return;
              ensureActiveTabVisible();
            });
          }
        },
      }),
    [],
  );

  useEffect(() => {
    return () => {
      tabContextMenu.destroy();
    };
  }, [tabContextMenu]);

  useEffect(() => {
    if (typeof document === "undefined") return;
    const input = document.createElement("input");
    input.type = "color";
    input.tabIndex = -1;
    // Match `main.ts`'s hidden color pickers (see styles/shell.css).
    input.className = "hidden-color-input shell-hidden-input";
    document.body.appendChild(input);
    tabColorPickerRef.current = input;
    // Capture the default value assigned by the browser for <input type="color"> so we can
    // reset the picker even when a sheet has no existing tab color. (Avoid hard-coding a
    // hex literal here; the token policy test enforces that UI colors live in CSS.)
    tabColorPickerDefaultValueRef.current = input.value;
    return () => {
      tabColorPickerRef.current = null;
      tabColorPickerDefaultValueRef.current = null;
      input.remove();
    };
  }, []);

  useEffect(() => {
    editingSheetIdRef.current = editingSheetId;
  }, [editingSheetId]);

  const stopAutoScroll = () => {
    const raf = autoScrollRef.current.raf;
    if (raf != null) cancelAnimationFrame(raf);
    autoScrollRef.current.raf = null;
    autoScrollRef.current.direction = 0;
  };

  const maybeAutoScroll = (clientX: number) => {
    const el = containerRef.current;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const threshold = 32;

    let direction: -1 | 0 | 1 = 0;
    if (clientX < rect.left + threshold) direction = -1;
    else if (clientX > rect.right - threshold) direction = 1;

    autoScrollRef.current.direction = direction;
    if (direction === 0) {
      stopAutoScroll();
      return;
    }

    if (autoScrollRef.current.raf != null) return;

    const tick = () => {
      const container = containerRef.current;
      if (!container) {
        stopAutoScroll();
        return;
      }
      const dir = autoScrollRef.current.direction;
      if (dir === 0) {
        stopAutoScroll();
        return;
      }
      container.scrollLeft += dir * 8;
      autoScrollRef.current.raf = requestAnimationFrame(tick);
    };
    autoScrollRef.current.raf = requestAnimationFrame(tick);
  };

  useEffect(() => {
    return () => {
      stopAutoScroll();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const clearDragIndicators = () => {
    setDropIndicator(null);
    setDraggingSheetId(null);
  };

  const commitRename = (sheetId: string): Promise<boolean> => {
    if (renameCommitRef.current) return renameCommitRef.current;

    const promise = (async () => {
      setRenameInFlight(true);
      try {
        if (mutationsDisabled) {
          throw new Error(mutationsDisabledMessage ?? READ_ONLY_SHEET_MUTATION_MESSAGE);
        }
        // Read the current input value at commit time instead of trusting the latest
        // `draftName` state. React batches updates inside event handlers, so it's possible
        // for a "commit" action (Enter/blur) to run before the state update from the last
        // keystroke has been flushed, which would otherwise drop characters from the rename.
        const currentValue = renameInputRef.current?.value ?? draftName;
        await onRenameSheet(sheetId, currentValue);
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        setRenameError(message);
        onError?.(message);
        requestAnimationFrame(() => renameInputRef.current?.focus());
        return false;
      } finally {
        setRenameInFlight(false);
        renameCommitRef.current = null;
      }

      setRenameError(null);
      setEditingSheetId(null);
      return true;
    })();

    renameCommitRef.current = promise;
    return promise;
  };

  const moveSheet = async (
    sheetId: string,
    dropTarget: Parameters<typeof computeWorkbookSheetMoveIndex>[0]["dropTarget"],
  ) => {
    if (mutationsDisabled) {
      reportMutationsDisabled();
      return;
    }
    const all = store.listAll();
    const fromIndex = all.findIndex((s) => s.id === sheetId);
    if (fromIndex < 0) return;
    const toIndex = computeWorkbookSheetMoveIndex({ sheets: all, fromSheetId: sheetId, dropTarget });
    if (toIndex == null) return;
    if (toIndex === fromIndex) return;

    try {
      store.move(sheetId, toIndex);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(message);
      return;
    }
    onSheetsReordered?.();

    if (onSheetMoved) {
      const seq = (moveCommitSeqRef.current += 1);
      try {
        await onSheetMoved({ sheetId, toIndex });
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        onError?.(message);

        // Best-effort rollback: only if no newer move has been committed since.
        if (moveCommitSeqRef.current === seq) {
          try {
            store.move(sheetId, fromIndex);
          } catch (rollbackErr) {
            onError?.(rollbackErr instanceof Error ? rollbackErr.message : String(rollbackErr));
          }
        }
      }
    }
  };

  const activateSheetWithRenameGuard = async (sheetId: string) => {
    if (editingSheetId && editingSheetId !== sheetId) {
      const ok = await commitRename(editingSheetId);
      if (!ok) return;
    }
    onActivateSheet(sheetId);
  };

  const beginRenameWithGuard = async (sheet: SheetMeta) => {
    if (mutationsDisabled) {
      reportMutationsDisabled();
      return;
    }
    if (renameInFlight) return;
    if (editingSheetId && editingSheetId !== sheet.id) {
      const ok = await commitRename(editingSheetId);
      if (!ok) return;
    }

    // Excel-style rename: start inline editing directly on the tab.
    setEditingSheetId(sheet.id);
    setDraftName(sheet.name);
    setRenameError(null);
    requestAnimationFrame(() => renameInputRef.current?.focus());
  };

  const openSheetPicker = useCallback(async () => {
    // Match the "Add sheet" behavior: if the user is mid-rename and the rename is invalid,
    // keep them in the rename flow instead of navigating away.
    if (editingSheetId) {
      const ok = await commitRename(editingSheetId);
      if (!ok) return;
    }

    const selected = await showQuickPick(
      visibleSheets.map((sheet) => ({
        label: sheet.name,
        value: sheet.id,
      })),
      { placeHolder: "Sheets" },
    );

    if (!selected) return;
    await activateSheetWithRenameGuard(selected);
  }, [activateSheetWithRenameGuard, commitRename, editingSheetId, visibleSheets]);

  const openSheetTabContextMenu = (sheetId: string, anchor: { x: number; y: number }) => {
    const sheet = store.getById(sheetId);
    if (!sheet) return;

    const allSheets = store.listAll();
    // Only allow unhiding "hidden" sheets. Excel does not offer UI affordances for
    // "veryHidden" sheets (those are typically VBA-only), so keep them out of the menu.
    const hiddenSheets = allSheets.filter((s) => s.visibility === "hidden");

    // Prevent deleting/hiding the last visible sheet.
    //
    // Note: deletion is still allowed in non-Tauri (web / e2e harness) environments; the optional
    // `onPersistSheetDelete` hook can no-op when no workbook backend is available.
    const canDelete = visibleSheets.length > 1;
    const canHide = sheet.visibility === "visible" && visibleSheets.length > 1;

    const items: ContextMenuItem[] = [
      {
        type: "item",
        label: "Rename",
        enabled: !mutationsDisabled,
        onSelect: () => {
          void beginRenameWithGuard(sheet).catch(() => {
            // Best-effort: avoid unhandled rejections from async rename bootstrapping.
          });
        },
      },
      {
        type: "item",
        label: "Hide",
        enabled: canHide && !mutationsDisabled,
        onSelect: async () => {
          const wasActive = sheet.id === activeSheetIdRef.current;
          let nextActiveId: string | null = null;
          if (wasActive) {
            nextActiveId = pickAdjacentVisibleSheetId(visibleSheets, sheet.id);
          }

          try {
            await onPersistSheetVisibility?.(sheet.id, "hidden");
          } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            onError?.(message);
            return;
          }

          try {
            store.hide(sheet.id);
          } catch (err) {
            const message = err instanceof Error ? err.message : String(err);
            onError?.(message);
            return;
          }

          if (wasActive) {
            const fallback = store.listVisible().at(0)?.id ?? null;
            const next = nextActiveId ?? fallback;
            if (next && next !== sheet.id) {
              onActivateSheet(next);
            }
            return;
          }
          onActivateSheet(activeSheetIdRef.current);
        },
      },
      {
        type: "submenu",
        label: "Unhide…",
        enabled: hiddenSheets.length > 0 && !mutationsDisabled,
        items: hiddenSheets.map((hidden) => ({
          type: "item" as const,
          label: hidden.name,
          onSelect: async () => {
            try {
              await onPersistSheetVisibility?.(hidden.id, "visible");
            } catch (err) {
              const message = err instanceof Error ? err.message : String(err);
              onError?.(message);
              return;
            }

            try {
              store.unhide(hidden.id);
            } catch (err) {
              const message = err instanceof Error ? err.message : String(err);
              onError?.(message);
              return;
            }
            onActivateSheet(activeSheetIdRef.current);
          },
        })),
      },
      {
        type: "submenu",
        label: "Tab Color",
        enabled: !mutationsDisabled,
        items: [
          {
            type: "item",
            label: "No Color",
            onSelect: async () => {
              try {
                await onPersistSheetTabColor?.(sheet.id, undefined);
              } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                onError?.(message);
                return;
              }

              try {
                store.setTabColor(sheet.id, undefined);
              } catch (err) {
                const message = err instanceof Error ? err.message : String(err);
                onError?.(message);
              }
            },
          },
          { type: "separator" },
          ...SHEET_TAB_COLOR_PALETTE.map((entry) => {
            const fallbackCss = excelArgbToCssHex(entry.excelArgb) ?? "";
            const rgb = resolveCssVar(entry.token, { fallback: fallbackCss });
            return {
              type: "item" as const,
              label: entry.label,
              // Use the resolved token value so the swatch color works in SVG `fill`
              // across WebView engines (some do not support `var(--token)` in SVG
              // presentation attributes).
              leading: { type: "swatch" as const, color: rgb },
              onSelect: async () => {
                // Prefer persisting the currently resolved token color so the workbook matches
                // what the user saw in the swatch, but fall back to the palette's canonical
                // ARGB if we cannot parse the CSS color (e.g. missing DOM APIs in tests).
                const excelArgb = normalizeCssHexToExcelArgb(rgb) ?? entry.excelArgb;
                const tabColor: TabColor = { rgb: excelArgb };
                try {
                  await onPersistSheetTabColor?.(sheet.id, tabColor);
                } catch (err) {
                  const message = err instanceof Error ? err.message : String(err);
                  onError?.(message);
                  return;
                }

                try {
                  store.setTabColor(sheet.id, tabColor);
                } catch (err) {
                  const message = err instanceof Error ? err.message : String(err);
                  onError?.(message);
                }
              },
            };
          }),
          { type: "separator" },
          {
            type: "item",
            label: "More Colors…",
            onSelect: () => {
              const input = tabColorPickerRef.current;
              if (!input) return;

              // Best-effort: use the current tab color as the initial value when it's a #RRGGBB hex string.
              // Otherwise, reset the picker to a token-backed default (so we don't leak the last selection from
              // a different sheet / previous open).
              const currentCss = normalizeExcelColorToCss(sheet.tabColor);
              const defaultValue = tabColorPickerDefaultValueRef.current ?? input.value;
              const initialValue = (() => {
                const normalized = String(currentCss ?? "").trim();
                if (/^#[0-9a-fA-F]{6}$/.test(normalized)) return normalized.toLowerCase();
                const grayFallbackArgb =
                  SHEET_TAB_COLOR_PALETTE.find((entry) => entry.token === "--sheet-tab-gray")?.excelArgb ?? "";
                const grayFallbackCss = excelArgbToCssHex(grayFallbackArgb) ?? defaultValue;
                const tokenFallback = resolveCssVar("--sheet-tab-gray", { fallback: grayFallbackCss }).trim();
                if (/^#[0-9a-fA-F]{6}$/.test(tokenFallback)) return tokenFallback.toLowerCase();
                return defaultValue;
              })();
              input.value = initialValue;

              // Preserve the current focus target so keyboard users return where they started.
              const restore = document.activeElement instanceof HTMLElement ? document.activeElement : null;

              // Avoid `addEventListener({ once: true })`: `<input type="color">` does not emit
              // a `change` event on cancel. Assigning `onchange` ensures we do not accumulate
              // listeners across cancels.
              input.onchange = () => {
                input.onchange = null;
                void (async () => {
                  const excelArgb = normalizeCssHexToExcelArgb(input.value);
                  if (!excelArgb) {
                    onError?.(`Failed to set tab color: selected color is not a hex value (${input.value}).`);
                    return;
                  }

                  const tabColor: TabColor = { rgb: excelArgb };
                  try {
                    await onPersistSheetTabColor?.(sheet.id, tabColor);
                  } catch (err) {
                    const message = err instanceof Error ? err.message : String(err);
                    onError?.(message);
                    return;
                  }

                  try {
                    store.setTabColor(sheet.id, tabColor);
                  } catch (err) {
                    const message = err instanceof Error ? err.message : String(err);
                    onError?.(message);
                  }
                })()
                  .finally(() => {
                    try {
                      if (restore?.isConnected) restore.focus({ preventScroll: true });
                    } catch {
                      // ignore
                    }
                  })
                  .catch(() => {
                    // Best-effort: avoid unhandled rejections from the `.finally` bookkeeping chain.
                  });
              };

              input.click();
            },
          },
        ],
      },
    ];

    items.push({ type: "separator" });
    items.push({
      type: "item",
      label: "Delete",
      enabled: canDelete && !mutationsDisabled,
      onSelect: () => {
        void deleteSheet(sheet).catch(() => {
          // Best-effort: avoid unhandled rejections from async delete flows.
        });
      },
    });

    tabContextMenu.open({ x: anchor.x, y: anchor.y, items });
  };

  const openSheetTabStripContextMenu = (anchor: { x: number; y: number }) => {
    // Only allow unhiding "hidden" sheets. Excel does not expose "veryHidden" sheets in the
    // standard UI (VBA-only), so keep them out of this picker.
    const hiddenSheets = store.listAll().filter((sheet) => sheet.visibility === "hidden");
    const items: ContextMenuItem[] = [
      {
        type: "item",
        label: "Unhide…",
        enabled: hiddenSheets.length > 0 && !mutationsDisabled,
        onSelect: async () => {
          try {
            const selected = await showQuickPick(
              hiddenSheets.map((sheet) => ({ label: sheet.name, value: sheet.id })),
              { placeHolder: "Unhide Sheet" },
            );
            if (!selected) return;
            try {
              await onPersistSheetVisibility?.(selected, "visible");
            } catch (err) {
              const message = err instanceof Error ? err.message : String(err);
              onError?.(message);
              return;
            }

            store.unhide(selected);
            // Keep the currently active sheet active (Excel-like), but re-trigger activation so
            // consumers (e.g. main.ts focus restoration) can reconcile if the active sheet was
            // previously hidden due to inconsistent metadata.
            onActivateSheet(activeSheetIdRef.current);
          } catch (err) {
            onError?.(err instanceof Error ? err.message : String(err));
          }
        },
      },
    ];

    tabContextMenu.open({ x: anchor.x, y: anchor.y, items });
  };

  const deleteSheet = async (sheet: SheetMeta): Promise<void> => {
    if (mutationsDisabled) {
      reportMutationsDisabled();
      return;
    }
    if (editingSheetId && editingSheetId !== sheet.id) {
      const ok = await commitRename(editingSheetId);
      if (!ok) return;
    }

    if (editingSheetId === sheet.id) return;

    let ok = false;
    try {
      ok = await nativeDialogs.confirm(`Delete sheet "${sheet.name}"?`);
    } catch {
      ok = false;
    }
    if (!ok) return;

    const deletedName = sheet.name;
    const allSheets = store.listAll();
    const sheetOrder = allSheets.map((s) => s.name);
    const nextActiveId =
      sheet.id === activeSheetIdRef.current ? pickAdjacentVisibleSheetId(allSheets, sheet.id) : null;

    try {
      await onPersistSheetDelete?.(sheet.id);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(message);
      return;
    }

    try {
      store.remove(sheet.id);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(message);
      return;
    }

    if (sheet.id === activeSheetIdRef.current) {
      const fallback = store.listVisible().at(0)?.id ?? store.listAll().at(0)?.id ?? null;
      const next = nextActiveId ?? fallback;
      if (next && next !== sheet.id) {
        onActivateSheet(next);
      }
    } else {
      // If we deleted a non-active sheet, re-focus the current sheet surface so the
      // user doesn't lose keyboard focus (especially after keyboard-invoked deletes).
      onActivateSheet(activeSheetIdRef.current);
    }

    try {
      onSheetDeleted?.({ sheetId: sheet.id, name: deletedName, sheetOrder });
    } catch (err) {
      onError?.(err instanceof Error ? err.message : String(err));
    }
  };

  const updateScrollButtons = useCallback(() => {
    const el = containerRef.current;
    if (!el) {
      setCanScroll({ left: false, right: false });
      return;
    }
    const maxScrollLeft = el.scrollWidth - el.clientWidth;
    setCanScroll({
      left: el.scrollLeft > 0,
      right: el.scrollLeft < maxScrollLeft - 1,
    });
  }, []);

  const scrollTabsBy = (delta: number) => {
    const el = containerRef.current;
    if (!el) return;
    const reducedMotion =
      (typeof document !== "undefined" &&
        document.documentElement?.getAttribute("data-reduced-motion") === "true") ||
      (typeof window !== "undefined" &&
        typeof window.matchMedia === "function" &&
        window.matchMedia("(prefers-reduced-motion: reduce)").matches);
    el.scrollBy({ left: delta, behavior: reducedMotion ? "auto" : "smooth" });
  };

  const isSheetDrag = (dt: DataTransfer): boolean => {
    // We set both the custom type and `text/plain` on dragStart. Some environments
    // can be finicky about exposing custom MIME types during drag operations.
    return dt.types.includes("text/sheet-id") || dt.types.includes("text/plain");
  };

  useEffect(() => {
    updateScrollButtons();
  }, [updateScrollButtons, visibleSheets.length]);

  useEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    updateScrollButtons();
    const onScroll = () => updateScrollButtons();
    el.addEventListener("scroll", onScroll, { passive: true });
    window.addEventListener("resize", onScroll);
    return () => {
      el.removeEventListener("scroll", onScroll);
      window.removeEventListener("resize", onScroll);
    };
  }, [updateScrollButtons]);

  useEffect(() => {
    // Keep the active tab visible when switching sheets via keyboard or programmatically.
    if (editingSheetId) return;
    if (tabContextMenu.isOpen()) {
      // Opening a context menu focuses its first menu item. If we scroll the active tab into view
      // while the menu is open (e.g. a delayed `scrollIntoView` after adding a sheet), the
      // resulting scroll event would immediately dismiss the menu. Defer it until close.
      pendingEnsureActiveTabVisibleRef.current = true;
      return;
    }
    pendingEnsureActiveTabVisibleRef.current = false;
    ensureActiveTabVisible();
  }, [activeSheetId, editingSheetId, tabContextMenu, visibleSheets.length]);

  return (
    <>
      <div className="sheet-nav">
        <button
          type="button"
          className="sheet-nav-btn"
          aria-label="Scroll sheet tabs left"
          tabIndex={-1}
          onClick={() => scrollTabsBy(-120)}
          disabled={!canScroll.left}
        >
          ‹
        </button>
        <button
          type="button"
          className="sheet-nav-btn"
          aria-label="Scroll sheet tabs right"
          tabIndex={-1}
          onClick={() => scrollTabsBy(120)}
          disabled={!canScroll.right}
        >
          ›
        </button>
      </div>

      <div
        className="sheet-tabs"
        ref={containerRef}
        role="tablist"
        aria-label="Sheets"
        aria-orientation="horizontal"
        data-dragging={draggingSheetId ? "true" : undefined}
        onContextMenu={(e) => {
          const target = e.target as Element | null;
          if (!target) return;

          // Let the native input context menu work while inline renaming.
          if (target.closest("input.sheet-tab__input")) return;

          // Tab context menus are handled by the tab button itself.
          if (target.closest('button[role="tab"][data-sheet-id]')) return;

          e.preventDefault();
          e.stopPropagation();

          // Restore focus to the surface that was active before the menu opened (typically the grid).
          const active = document.activeElement;
          lastContextMenuFocusRef.current = active instanceof HTMLElement ? active : null;

          openSheetTabStripContextMenu({ x: e.clientX, y: e.clientY });
        }}
        onKeyDown={(e) => {
          if (e.defaultPrevented) return;

          const target = e.target as HTMLElement | null;
          if (target && (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)) return;

          const tabs = containerRef.current
            ? Array.from(containerRef.current.querySelectorAll<HTMLButtonElement>('button[role="tab"]'))
            : [];

          if (target instanceof HTMLButtonElement && target.getAttribute("role") === "tab" && tabs.length > 0) {
            const idx = tabs.indexOf(target);
            if (idx !== -1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
               const focusTab = (nextIdx: number) => {
                 const next = tabs[nextIdx];
                 if (!next) return;
                 next.focus();
                 if (typeof (next as any).scrollIntoView === "function") {
                   next.scrollIntoView({ block: "nearest", inline: "nearest" });
                 }
               };

              const isContextMenuKey = (e.shiftKey && e.key === "F10") || e.key === "ContextMenu" || e.code === "ContextMenu";
              if (isContextMenuKey) {
                e.preventDefault();
                e.stopPropagation();
                const sheetId = target.dataset.sheetId;
                if (!sheetId) return;
                const activeEditingId = editingSheetId;
                void (async () => {
                  if (activeEditingId && activeEditingId !== sheetId) {
                    const ok = await commitRename(activeEditingId);
                    if (!ok) return;
                  }
                  lastContextMenuFocusRef.current = target;
                  const rect = target.getBoundingClientRect();
                  openSheetTabContextMenu(sheetId, { x: rect.left + rect.width / 2, y: rect.bottom });
                })().catch(() => {
                  // Best-effort: context menu opening should never surface as an unhandled rejection.
                });
                return;
              }

              if (e.key === "ArrowRight") {
                e.preventDefault();
                e.stopPropagation();
                focusTab(Math.min(tabs.length - 1, idx + 1));
                return;
              }
              if (e.key === "ArrowLeft") {
                e.preventDefault();
                e.stopPropagation();
                focusTab(Math.max(0, idx - 1));
                return;
              }
              if (e.key === "Home") {
                e.preventDefault();
                e.stopPropagation();
                focusTab(0);
                return;
              }
              if (e.key === "End") {
                e.preventDefault();
                e.stopPropagation();
                focusTab(Math.max(0, tabs.length - 1));
                return;
              }

              // Enter/Space are handled by the <button> itself to activate the focused tab.
            }
          }

           // Ctrl/Cmd+PgUp/PgDn sheet navigation is handled globally by the desktop shell's
           // keybinding/command registry so focus restoration remains consistent across
           // different UI surfaces.
         }}
        onDragOver={(e) => {
          if (!isSheetDrag(e.dataTransfer)) return;
          e.preventDefault();
          e.dataTransfer.dropEffect = "move";
          maybeAutoScroll(e.clientX);

          // Only apply an "end" indicator when dragging over the container itself (not over a tab),
          // otherwise the per-tab handler will compute before/after.
          if (e.target === e.currentTarget) {
            const last = visibleSheets.at(-1)?.id ?? null;
            if (last) setDropIndicator({ targetSheetId: last, position: "after" });
          }
        }}
        onDrop={(e) => {
          if (!isSheetDrag(e.dataTransfer)) return;
          e.preventDefault();
          stopAutoScroll();
          const fromId = e.dataTransfer.getData("text/sheet-id") || e.dataTransfer.getData("text/plain");
          if (!fromId) {
            clearDragIndicators();
            return;
          }
          // Dropping on the container inserts at the end of the visible list.
          void moveSheet(fromId, { kind: "end" }).catch(() => {
            // Best-effort: avoid unhandled rejections from async move handlers.
          });
          clearDragIndicators();
        }}
        onDragLeave={(e) => {
          stopAutoScroll();
          // Avoid flickering the drop indicator when moving between child tabs.
          if (e.target === e.currentTarget) setDropIndicator(null);
        }}
      >
        {visibleSheets.map((sheet) => (
          <SheetTab
            key={sheet.id}
            sheet={sheet}
            active={sheet.id === activeSheetId}
            editing={editingSheetId === sheet.id}
            readOnly={mutationsDisabled}
            dragging={draggingSheetId === sheet.id}
            dropPosition={dropIndicator?.targetSheetId === sheet.id ? dropIndicator.position : null}
            draftName={draftName}
            renameError={editingSheetId === sheet.id ? renameError : null}
            renameInputRef={renameInputRef}
            tabRef={sheet.id === activeSheetId ? activeTabRef : undefined}
            onActivate={() => {
              void activateSheetWithRenameGuard(sheet.id).catch(() => {
                // Best-effort: avoid unhandled rejections from async sheet activation.
              });
            }}
            onBeginRename={() => {
              void beginRenameWithGuard(sheet).catch(() => {
                // Best-effort: avoid unhandled rejections from async rename bootstrapping.
              });
            }}
            onContextMenu={(e) => {
              e.preventDefault();
              e.stopPropagation();
              const activeEditingId = editingSheetId;
              const anchor = { x: e.clientX, y: e.clientY };
              const target = e.currentTarget;
              void (async () => {
                if (activeEditingId && activeEditingId !== sheet.id) {
                  const ok = await commitRename(activeEditingId);
                  if (!ok) return;
                }
                lastContextMenuFocusRef.current = target;
                openSheetTabContextMenu(sheet.id, anchor);
              })().catch(() => {
                // Best-effort: context menu opening should never surface as an unhandled rejection.
              });
            }}
            renameInFlight={renameInFlight && editingSheetId === sheet.id}
            onCommitRename={() => void commitRename(sheet.id)}
            onCancelRename={() => {
              if (renameInFlight) return;
              setEditingSheetId(null);
              setRenameError(null);
            }}
            onDraftNameChange={setDraftName}
            onDragStart={() => {
              stopAutoScroll();
              setDraggingSheetId(sheet.id);
              setDropIndicator(null);
            }}
            onDragEnd={() => {
              stopAutoScroll();
              clearDragIndicators();
            }}
            onDragOverTab={(e) => {
              if (!isSheetDrag(e.dataTransfer)) return;
              if (draggingSheetId && draggingSheetId === sheet.id) {
                setDropIndicator(null);
                return;
              }

              const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
              const shouldInsertAfter = e.clientX > rect.left + rect.width / 2;
              setDropIndicator({ targetSheetId: sheet.id, position: shouldInsertAfter ? "after" : "before" });
            }}
            onDropOnTab={(e) => {
              stopAutoScroll();
              clearDragIndicators();
              const fromId = e.dataTransfer.getData("text/sheet-id") || e.dataTransfer.getData("text/plain");
              if (!fromId || fromId === sheet.id) return;

              const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
              const shouldInsertAfter = e.clientX > rect.left + rect.width / 2;
              void moveSheet(fromId, {
                kind: shouldInsertAfter ? "after" : "before",
                targetSheetId: sheet.id,
              }).catch(() => {
                // Best-effort: avoid unhandled rejections from async move handlers.
              });
            }}
          />
        ))}
      </div>

      <button
        type="button"
        className="sheet-add"
        data-testid="sheet-add"
        disabled={mutationsDisabled}
        onClick={() => {
          void (async () => {
            if (editingSheetId) {
              const ok = await commitRename(editingSheetId);
              if (!ok) return;
            }
            await onAddSheet();
          })().catch(() => {
            // Best-effort: avoid unhandled rejections from async sheet creation.
          });
        }}
        aria-label="Add sheet"
      >
        +
      </button>

      <button
        type="button"
        className="sheet-overflow"
        data-testid="sheet-overflow"
        aria-label="Show sheet list"
        onClick={() => {
          void openSheetPicker().catch(() => {
            // Best-effort: avoid unhandled rejections from the async sheet picker.
          });
        }}
      >
        ⋯
      </button>
    </>
  );
}

function SheetTab(props: {
  sheet: SheetMeta;
  active: boolean;
  editing: boolean;
  readOnly: boolean;
  dragging: boolean;
  dropPosition: "before" | "after" | null;
  draftName: string;
  renameError: string | null;
  renameInFlight: boolean;
  renameInputRef: React.RefObject<HTMLInputElement>;
  tabRef?: React.Ref<HTMLButtonElement>;
  onActivate: (event: React.MouseEvent<HTMLButtonElement>) => void;
  onBeginRename: () => void;
  onContextMenu: (e: React.MouseEvent<HTMLButtonElement>) => void;
  onCommitRename: () => void;
  onCancelRename: () => void;
  onDraftNameChange: (name: string) => void;
  onDragStart: () => void;
  onDragEnd: () => void;
  onDragOverTab: (e: React.DragEvent<HTMLButtonElement>) => void;
  onDropOnTab: (e: React.DragEvent<HTMLButtonElement>) => void;
}) {
  const { sheet, active, editing, draftName, renameError, renameInFlight } = props;
  const cancelBlurCommitRef = useRef(false);
  const tabColorCss = !editing ? (normalizeExcelColorToCss(sheet.tabColor) ?? null) : null;
  const ariaLabel = sheet.visibility === "visible" ? sheet.name : `${sheet.name} (${sheet.visibility})`;

  return (
    <button
      type="button"
      className="sheet-tab"
      role="tab"
      aria-selected={active}
      aria-label={ariaLabel}
      tabIndex={active ? 0 : -1}
      data-testid={`sheet-tab-${sheet.id}`}
      data-sheet-id={sheet.id}
      data-active={active ? "true" : "false"}
      data-tab-color={tabColorCss ?? undefined}
      data-dragging={props.dragging ? "true" : undefined}
      data-drop-position={props.dropPosition ?? undefined}
      draggable={!editing && !props.readOnly}
      ref={props.tabRef}
      onClick={(event) => {
        // React dispatches a separate `click` event for each click in a double-click.
        // Avoid treating the second click as another activation; `onDoubleClick` handles rename.
        if (editing) return;
        if (event.detail > 1) return;
        props.onActivate(event);
      }}
      onDoubleClick={() => {
        if (!editing) props.onBeginRename();
      }}
      onContextMenu={(e) => {
        if (editing) return;
        props.onContextMenu(e);
      }}
      onDragStart={(e) => {
        props.onDragStart();
        e.dataTransfer.setData("text/sheet-id", sheet.id);
        e.dataTransfer.setData("text/plain", sheet.id);
        e.dataTransfer.effectAllowed = "move";
      }}
      onDragOver={(e) => {
        if (!e.dataTransfer.types.includes("text/sheet-id") && !e.dataTransfer.types.includes("text/plain")) return;
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        props.onDragOverTab(e);
      }}
      onDrop={(e) => {
        if (!e.dataTransfer.types.includes("text/sheet-id") && !e.dataTransfer.types.includes("text/plain")) return;
        e.preventDefault();
        e.stopPropagation();
        props.onDropOnTab(e);
      }}
      onDragEnd={() => props.onDragEnd()}
    >
      {editing ? (
        <span className="sheet-tab__pill sheet-tab__pill--editing">
          <input
            ref={props.renameInputRef}
            className="sheet-tab__input"
            value={draftName}
            autoFocus
            readOnly={renameInFlight}
            aria-busy={renameInFlight ? true : undefined}
            aria-invalid={renameError ? true : undefined}
            title={renameError ?? undefined}
            onClick={(e) => e.stopPropagation()}
            onChange={(e) => props.onDraftNameChange(e.target.value)}
            onFocus={(e) => e.currentTarget.select()}
            onBlur={() => {
              if (renameInFlight) return;
              if (cancelBlurCommitRef.current) {
                cancelBlurCommitRef.current = false;
                return;
              }
              props.onCommitRename();
            }}
            onKeyDown={(e) => {
              e.stopPropagation();
              if (renameInFlight) return;
              if (e.key === "Enter") {
                e.preventDefault();
                props.onCommitRename();
              }
              if (e.key === "Escape") {
                e.preventDefault();
                cancelBlurCommitRef.current = true;
                props.onCancelRename();
              }
            }}
          />
        </span>
      ) : (
        <span className="sheet-tab__pill">
          <span className="sheet-tab__name">{sheet.name}</span>
          {tabColorCss ? <span className="sheet-tab__color" style={{ background: tabColorCss }} /> : null}
        </span>
      )}
    </button>
  );
}
