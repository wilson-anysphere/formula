import type { CommandRegistry } from "../extensions/commandRegistry.js";
import type { LayoutController } from "../layout/layoutController.js";
import { getPanelPlacement } from "../layout/layoutState.js";
import { PanelIds } from "../panels/panelRegistry.js";
import { t } from "../i18n/index.js";

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
  getPowerQueryService: () => PowerQueryServiceLike | null;
  showToast: ToastFn;
  notify: NotifyFn;
  /**
   * Optional time source (for tests).
   */
  now?: () => number;
}): void {
  const { commandRegistry, layoutController, getPowerQueryService, showToast, notify, now = () => Date.now() } = params;

  commandRegistry.registerBuiltinCommand(
    DATA_QUERIES_RIBBON_COMMANDS.toggleQueriesConnections,
    // Match the ribbon label so the command palette entry is discoverable by the same
    // name users see in the UI (while still keeping the canonical panel toggle command).
    "Queries & Connections",
    (next?: boolean) => {
      if (!layoutController) {
        showToast("Queries panel is not available (layout controller missing).", "error");
        return;
      }

      const placement = getPanelPlacement(layoutController.layout, PanelIds.DATA_QUERIES);
      const open = () => layoutController.openPanel(PanelIds.DATA_QUERIES);
      const close = () => layoutController.closePanel(PanelIds.DATA_QUERIES);

      if (typeof next === "boolean") {
        if (next) open();
        else close();
        return;
      }

      // Toggle when no explicit state was provided (command palette / programmatic use).
      if (placement.kind === "closed") open();
      else close();
    },
    {
      category: t("commandCategory.data"),
      icon: null,
      keywords: ["data", "queries", "connections", "power query", "panel"],
    },
  );

  const refreshAll = () => {
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
  };

  const refreshCommandIds: string[] = [
    DATA_QUERIES_RIBBON_COMMANDS.refreshAll,
    DATA_QUERIES_RIBBON_COMMANDS.refreshAllRefresh,
    DATA_QUERIES_RIBBON_COMMANDS.refreshAllConnections,
    DATA_QUERIES_RIBBON_COMMANDS.refreshAllQueries,
  ];

  for (const commandId of refreshCommandIds) {
    commandRegistry.registerBuiltinCommand(commandId, "Refresh All Queries", refreshAll, {
      category: t("commandCategory.data"),
      icon: null,
      keywords: ["refresh", "power query", "queries", "connections"],
    });
  }
}
