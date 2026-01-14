import type { CommandRegistry } from "../extensions/commandRegistry.js";
import type { LayoutController } from "../layout/layoutController.js";
import { getPanelPlacement } from "../layout/layoutState.js";
import { PanelIds } from "../panels/panelRegistry.js";
import { t } from "../i18n/index.js";
import { showCollabEditRejectedToast } from "../collab/editRejectionToast.js";

export const DATA_QUERIES_RIBBON_COMMANDS = {
  toggleQueriesConnections: "data.queriesConnections.queriesConnections",
  refreshAll: "data.queriesConnections.refreshAll",
  refreshAllRefresh: "data.queriesConnections.refreshAll.refresh",
  refreshAllConnections: "data.queriesConnections.refreshAll.refreshAllConnections",
  refreshAllQueries: "data.queriesConnections.refreshAll.refreshAllQueries",
} as const;

type ToastFn = (message: string, type?: "info" | "success" | "warning" | "error", options?: any) => void;
type NotifyFn = (options: { title: string; body: string }) => Promise<void> | void;

type PowerQueryServiceLike = {
  ready: Promise<void>;
  getQueries: () => unknown[];
  refreshAll: () => { promise: Promise<unknown> };
};

export function registerDataQueriesCommands(params: {
  commandRegistry: CommandRegistry;
  layoutController: LayoutController | null;
  /**
   * Optional spreadsheet edit-state predicate.
   *
   * When omitted, falls back to the desktop-shell-owned `globalThis.__formulaSpreadsheetIsEditing`
   * flag (when present).
   *
   * The desktop shell passes a custom predicate (`isSpreadsheetEditing`) that includes split-view
   * secondary editor state so command palette/keybindings cannot bypass ribbon disabling.
   */
  isEditing?: (() => boolean) | null;
  /**
   * Optional spreadsheet read-only predicate.
   *
   * When omitted, falls back to the SpreadsheetApp-owned `globalThis.__formulaSpreadsheetIsReadOnly`
   * flag (when present).
   *
   * The desktop ribbon disables refresh commands in read-only collab roles; guard execution so
   * command palette/keybindings cannot bypass that state.
   */
  isReadOnly?: (() => boolean) | null;
  getPowerQueryService: () => PowerQueryServiceLike | null;
  showToast: ToastFn;
  notify: NotifyFn;
  /**
   * Optional hook to keep ribbon UI-state (pressed toggles) in sync when a command
   * cannot execute (e.g. missing layout controller in non-desktop builds).
   */
  refreshRibbonUiState?: (() => void) | null;
  /**
   * Optional focus restoration hook (typically `SpreadsheetApp.focus()`).
   *
   * Some ribbon commands historically restored focus immediately so long-running async
   * work (like Power Query refresh) doesn't leave the ribbon as the focused surface.
   */
  focusAfterExecute?: (() => void) | null;
  /**
   * Optional time source (for tests).
   */
  now?: () => number;
}): void {
  const {
    commandRegistry,
    layoutController,
    isEditing = null,
    isReadOnly = null,
    getPowerQueryService,
    showToast,
    notify,
    refreshRibbonUiState = null,
    focusAfterExecute = null,
    now = () => Date.now(),
  } = params;
  const isEditingFn = isEditing ?? (() => (globalThis as any).__formulaSpreadsheetIsEditing === true);
  const isReadOnlyFn = isReadOnly ?? (() => (globalThis as any).__formulaSpreadsheetIsReadOnly === true);

  commandRegistry.registerBuiltinCommand(
    DATA_QUERIES_RIBBON_COMMANDS.toggleQueriesConnections,
    // Match the ribbon label so the command palette entry is discoverable by the same
    // name users see in the UI (while still keeping the canonical panel toggle command).
    "Queries & Connections",
    (next?: boolean) => {
      if (!layoutController) {
        showToast("Queries panel is not available (layout controller missing).", "error");
        // Ensure the ribbon toggle state reflects the actual panel placement.
        refreshRibbonUiState?.();
        focusAfterExecute?.();
        return;
      }

      const placement = getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES);
      const open = () => {
        layoutController.openPanel(PanelIds.DATA_QUERIES);
        // Floating panels can be minimized; opening should restore them (Excel-style behavior).
        try {
          const floating = (layoutController.layout as any)?.floating?.[PanelIds.DATA_QUERIES];
          if (floating?.minimized) {
            layoutController.setFloatingPanelMinimized(PanelIds.DATA_QUERIES, false);
          }
        } catch {
          // Best-effort: ignore layout shape mismatches.
        }
      };
      const close = () => layoutController.closePanel(PanelIds.DATA_QUERIES);

      if (typeof next === "boolean") {
        if (next) open();
        else close();
        focusAfterExecute?.();
        return;
      }

      // Toggle when no explicit state was provided (command palette / programmatic use).
      const isMinimizedFloating = (() => {
        if (placement.kind !== "floating") return false;
        try {
          return Boolean((layoutController.layout as any)?.floating?.[PanelIds.DATA_QUERIES]?.minimized);
        } catch {
          return false;
        }
      })();
      const isCollapsedDock = (() => {
        if (placement.kind !== "docked") return false;
        try {
          return Boolean((layoutController.layout as any)?.docks?.[placement.side]?.collapsed);
        } catch {
          return false;
        }
      })();
      if (placement.kind === "closed" || isMinimizedFloating || isCollapsedDock) open();
      else close();

      focusAfterExecute?.();
    },
    {
      category: t("commandCategory.data"),
      icon: null,
      keywords: ["data", "queries", "connections", "power query", "panel"],
    },
  );

  const refreshAll = () => {
    if (isEditingFn()) return;
    if (isReadOnlyFn()) {
      showCollabEditRejectedToast([{ rejectionKind: "dataQueriesRefresh", rejectionReason: "permission" }]);
      focusAfterExecute?.();
      return;
    }
    void (async () => {
      const service = getPowerQueryService();
      if (!service) {
        showToast("Queries service not available");
        return;
      }

      try {
        await service.ready;
      } catch (err) {
        console.error("Power Query service failed to initialize:", err);
        showToast("Queries service not available", "error");
        return;
      }

      const queries = service.getQueries();
      if (!queries.length) {
        showToast("No queries to refresh");
        return;
      }

      const startedAtMs = now();
      const queryCount = queries.length;

      const shouldNotifyInBackground = (): boolean => {
        try {
          if (typeof document === "undefined") return false;
          // Only notify when the app is not focused / not visible (user likely switched away).
          if ((document as any).hidden) return true;
          if (typeof document.hasFocus === "function") return !document.hasFocus();
          return false;
        } catch {
          return false;
        }
      };

      try {
        const handle = service.refreshAll();
        await handle.promise;
        const elapsedMs = now() - startedAtMs;
        // Avoid spamming notifications for extremely fast refreshes; only notify when the
        // user likely switched away or when the refresh took long enough to be meaningful.
        if (shouldNotifyInBackground() || elapsedMs >= 5_000) {
          const noun = queryCount === 1 ? "query" : "queries";
          void notify({ title: "Power Query refresh complete", body: `Refreshed ${queryCount} ${noun}.` });
        }
      } catch (err) {
        console.error("Failed to refresh all queries:", err);
        showToast(`Failed to refresh queries: ${String(err)}`, "error");
        if (shouldNotifyInBackground()) {
          void notify({ title: "Power Query refresh failed", body: "One or more queries failed to refresh." });
        }
      }
    })();

    // Don't wait for the refresh to complete; restore focus immediately so long-running
    // refresh jobs don't steal focus later when their promise settles.
    focusAfterExecute?.();
  };

  const refreshCommands: Array<{ id: string; title: string }> = [
    { id: DATA_QUERIES_RIBBON_COMMANDS.refreshAll, title: "Refresh All" },
    { id: DATA_QUERIES_RIBBON_COMMANDS.refreshAllRefresh, title: "Refresh" },
    { id: DATA_QUERIES_RIBBON_COMMANDS.refreshAllConnections, title: "Refresh All Connections" },
    { id: DATA_QUERIES_RIBBON_COMMANDS.refreshAllQueries, title: "Refresh All Queries" },
  ];

  for (const { id, title } of refreshCommands) {
    commandRegistry.registerBuiltinCommand(id, title, refreshAll, {
      category: t("commandCategory.data"),
      icon: null,
      description: "Refresh all Power Query queries and connections",
      keywords: ["refresh", "power query", "queries", "connections"],
    });
  }
}
