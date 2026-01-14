import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

import { CredentialManager, httpScope, normalizeScopes, oauth2Scope, randomId, type Query } from "@formula/power-query";

import { isLoopbackRedirectUrl, matchesRedirectUri, oauthBroker } from "../../power-query/oauthBroker.js";
import { createPowerQueryRefreshStateStore } from "../../power-query/refreshStateStore.js";
import { loadOAuth2ProviderConfigs, saveOAuth2ProviderConfigs, type OAuth2ProviderConfig } from "../../power-query/oauthProviders.ts";
import { deriveQueryListRows, reduceQueryRuntimeState, type QueryRuntimeState } from "../../power-query/queryRuntime.ts";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver";
import {
  DesktopPowerQueryService,
  getDesktopPowerQueryService,
  onDesktopPowerQueryServiceChanged,
} from "../../power-query/service.js";
import * as nativeDialogs from "../../tauri/nativeDialogs.js";
import { getTauriEventApiOrNull, hasTauri, hasTauriInvoke } from "../../tauri/api";
import { showInputBox } from "../../extensions/ui.js";

import { PanelIds } from "../panelRegistry.js";

type Props = {
  getDocumentController: () => any;
  workbookId?: string;
  sheetNameResolver?: SheetNameResolver | null;
  /**
   * Optional SpreadsheetApp-like object for UI state (read-only) detection.
   *
   * The Power Query service itself operates on a DocumentController, but some panel actions
   * (refresh/load) should be disabled in read-only roles for consistency with ribbon disabling.
   */
  app?: { isReadOnly?: () => boolean } | null;
};

type PendingPkce = { providerId: string; redirectUri: string };
type PendingDeviceCode = { providerId: string; code: string; verificationUri: string };

type StorageLike = { getItem(key: string): string | null; setItem(key: string, value: string): void };

function getLocalStorageOrNull(): StorageLike | null {
  try {
    const storage = (globalThis as any)?.localStorage as StorageLike | undefined;
    if (storage && typeof storage.getItem === "function" && typeof storage.setItem === "function") return storage;
  } catch {
    // ignore
  }
  return null;
}

function selectedQueryKey(workbookId: string): string {
  return `formula.desktop.powerQuery.selectedQuery:${workbookId}`;
}

function saveSelectedQueryId(workbookId: string, queryId: string): void {
  const storage = getLocalStorageOrNull();
  if (!storage) return;
  try {
    storage.setItem(selectedQueryKey(workbookId), queryId);
  } catch {
    // ignore
  }
}

const RECOMMENDED_DESKTOP_OAUTH_REDIRECT_URI = "formula://oauth/callback";
const EXAMPLE_LOOPBACK_OAUTH_REDIRECT_URI = "http://127.0.0.1:4242/oauth/callback";

function hasTauriEventApi(): boolean {
  return getTauriEventApiOrNull() != null;
}

function supportsDesktopOAuthRedirectCapture(redirectUri: string): boolean {
  if (!hasTauri() || !hasTauriEventApi()) return false;
  if (typeof redirectUri !== "string" || redirectUri.trim() === "") return false;
  try {
    const url = new URL(redirectUri);
    const protocol = url.protocol.toLowerCase();

    // Custom-scheme deep links (e.g. `formula://oauth/callback`).
    if (protocol === "formula:") return true;

    // Loopback redirect capture (RFC 8252) for providers that support it.
    if (isLoopbackRedirectUrl(url)) {
      return hasTauriInvoke();
    }

    return false;
  } catch {
    return false;
  }
}

function dispatchOpenPanel(detail: Record<string, unknown>) {
  try {
    if (typeof window === "undefined") return;
    window.dispatchEvent(new CustomEvent("formula:open-panel", { detail }));
  } catch {
    // ignore
  }
}

function formatTimestamp(ms: number | null | undefined): string {
  if (typeof ms !== "number" || !Number.isFinite(ms)) return "—";
  const date = new Date(ms);
  if (Number.isNaN(date.getTime())) return "—";
  return date.toLocaleString();
}

function sqlScopeKey(scope: any): string {
  if (!scope || typeof scope !== "object") return "<unknown>";
  const server = typeof (scope as any).server === "string" ? String((scope as any).server) : "";
  const database = typeof (scope as any).database === "string" ? String((scope as any).database) : "";
  const user = typeof (scope as any).user === "string" ? String((scope as any).user) : "";
  return `${server}|${database}|${user}`;
}

function httpScopeKey(scope: any): string {
  if (!scope || typeof scope !== "object") return "<unknown>";
  const origin = typeof (scope as any).origin === "string" ? String((scope as any).origin) : "";
  const realm = typeof (scope as any).realm === "string" ? String((scope as any).realm) : "";
  return `${origin}|${realm}`;
}

function safeHttpScope(url: string): any | null {
  try {
    return httpScope({ url, realm: null }) as any;
  } catch {
    return null;
  }
}

function safeHttpUrl(value: string): string | null {
  const text = typeof value === "string" ? value.trim() : "";
  if (!text) return null;
  try {
    const parsed = new URL(text);
    if (parsed.protocol !== "http:" && parsed.protocol !== "https:") return null;
    return parsed.toString();
  } catch {
    return null;
  }
}

function inferSqlUser(connection: unknown): string {
  if (typeof connection === "string") {
    try {
      const url = new URL(connection);
      return url.username || "";
    } catch {
      return "";
    }
  }
  if (!connection || typeof connection !== "object" || Array.isArray(connection)) return "";
  const record = connection as any;
  if (typeof record.user === "string") return record.user;
  if (typeof record.username === "string") return record.username;
  return "";
}

function createBlankQuery(args: { name?: string } = {}): Query {
  return {
    id: `q_${randomId(8)}`,
    name: args.name ?? "Query",
    source: { type: "range", range: { values: [["Value"], [1], [2], [3]], hasHeaders: true } },
    steps: [],
    refreshPolicy: { type: "manual" },
  };
}

function duplicateQueryDefinition(existing: Query): Query {
  const baseName = existing.name?.trim() || "Query";
  const copyId = `q_${randomId(8)}`;
  const copyName = `${baseName} (copy)`;
  // Avoid copying the destination so the duplicate doesn't overwrite the same output range.
  const { destination: _destination, ...rest } = existing as any;
  return { ...(rest as Query), id: copyId, name: copyName };
}

export function DataQueriesPanelContainer(props: Props) {
  const workbookId = props.workbookId ?? "default";
  const docController = props.getDocumentController();
  const app = props.app ?? null;

  const [isReadOnly, setIsReadOnly] = useState<boolean>(() => {
    if (!app || typeof app.isReadOnly !== "function") return false;
    try {
      return Boolean(app.isReadOnly());
    } catch {
      return false;
    }
  });

  const [isEditing, setIsEditing] = useState<boolean>(() => {
    const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
    return globalEditing === true;
  });

  const mutationsDisabled = isReadOnly || isEditing;

  useEffect(() => {
    if (typeof window === "undefined") return;
    const onReadOnlyChanged = (evt: Event) => {
      const detail = (evt as CustomEvent)?.detail as any;
      if (detail && typeof detail.readOnly === "boolean") {
        setIsReadOnly(detail.readOnly);
        return;
      }
      if (!app || typeof app.isReadOnly !== "function") return;
      try {
        setIsReadOnly(Boolean(app.isReadOnly()));
      } catch {
        // ignore
      }
    };
    window.addEventListener("formula:read-only-changed", onReadOnlyChanged as EventListener);
    return () => window.removeEventListener("formula:read-only-changed", onReadOnlyChanged as EventListener);
  }, [app]);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const onEditingChanged = (evt: Event) => {
      const detail = (evt as CustomEvent)?.detail as any;
      if (detail && typeof detail.isEditing === "boolean") {
        setIsEditing(detail.isEditing);
        return;
      }
      const globalEditing = (globalThis as any).__formulaSpreadsheetIsEditing;
      setIsEditing(globalEditing === true);
    };
    window.addEventListener("formula:spreadsheet-editing-changed", onEditingChanged as EventListener);
    return () => window.removeEventListener("formula:spreadsheet-editing-changed", onEditingChanged as EventListener);
  }, []);

  const [service, setService] = useState<DesktopPowerQueryService | null>(() => getDesktopPowerQueryService(workbookId));

  useEffect(() => {
    return onDesktopPowerQueryServiceChanged(workbookId, setService);
  }, [workbookId]);

  useEffect(() => {
    if (service) return;
    if (hasTauri()) return;

    const local = new DesktopPowerQueryService({
      workbookId,
      document: docController,
      concurrency: 1,
      batchSize: 1024,
    });

    setService(local);
    return () => local.dispose();
  }, [docController, service, workbookId]);

  const credentialStore = service?.credentialStore ?? null;
  const oauth2Manager = service?.oauth2Manager ?? null;

  const sqlCredentialManager = useMemo(() => {
    if (!credentialStore) return null;
    try {
      return new CredentialManager({ store: credentialStore as any });
    } catch {
      return null;
    }
  }, [credentialStore]);

  const resolveSqlScope = useCallback(
    (connection: unknown): any | null => {
      if (!sqlCredentialManager) return null;
      try {
        return sqlCredentialManager.resolveScope("sql", { connection });
      } catch {
        return null;
      }
    },
    [sqlCredentialManager],
  );

  const [queries, setQueries] = useState<Query[]>(() => service?.getQueries?.() ?? []);
  const [runtimeState, setRuntimeState] = useState<QueryRuntimeState>({});
  const runtimeRef = useRef<QueryRuntimeState>({});

  const [lastRunAtMsByQueryId, setLastRunAtMsByQueryId] = useState<Record<string, number>>({});

  const [oauthProviders, setOauthProviders] = useState<OAuth2ProviderConfig[]>(() => loadOAuth2ProviderConfigs(workbookId));
  const [editingProvider, setEditingProvider] = useState<OAuth2ProviderConfig | null>(null);
  const [pendingPkce, setPendingPkce] = useState<PendingPkce | null>(null);
  const [pendingDeviceCode, setPendingDeviceCode] = useState<PendingDeviceCode | null>(null);
  const [globalError, setGlobalError] = useState<string | null>(null);

  const normalizeOAuthScopes = useCallback(
    (providerId: string, scopes: string[] | undefined) => {
      // Keep our credential store keying consistent with OAuth2Manager behavior:
      // if the query does not explicitly specify scopes, fall back to the provider's
      // configured defaultScopes (if any). Otherwise the UI may incorrectly show
      // "signed out" and sign-out may delete the wrong credential entry.
      const provider = oauthProviders.find((p) => p.id === providerId);
      const resolvedScopes = scopes ?? (provider as any)?.defaultScopes;
      return normalizeScopes(resolvedScopes as any);
    },
    [oauthProviders],
  );

  const activeRefreshHandleByQueryId = useRef(new Map<string, { jobId: string; cancel: () => void }>());
  const [activeRefreshAll, setActiveRefreshAll] = useState<{ sessionId: string; cancel: () => void } | null>(null);

  useEffect(() => {
    setOauthProviders(loadOAuth2ProviderConfigs(workbookId));
  }, [workbookId]);

  // Keep oauth provider registry in sync with stored configs.
  useEffect(() => {
    if (!oauth2Manager) return;
    for (const provider of oauthProviders) {
      try {
        oauth2Manager.registerProvider(provider as any);
      } catch {
        // Ignore invalid provider configs so the panel doesn't crash.
      }
    }
  }, [oauth2Manager, oauthProviders]);

  // Wire broker handlers for OAuth flows.
  useEffect(() => {
    oauthBroker.setDeviceCodePromptHandler((code, verificationUri) => {
      setPendingDeviceCode((prev) => prev ?? { providerId: "<unknown>", code, verificationUri });
    });
    return () => {
      oauthBroker.setDeviceCodePromptHandler(null);
    };
  }, []);

  useEffect(() => {
    const store = createPowerQueryRefreshStateStore({ workbookId });
    let cancelled = false;
    store
      .load()
      .then((state) => {
        if (cancelled) return;
        const map: Record<string, number> = {};
        for (const [queryId, entry] of Object.entries(state ?? {})) {
          const lastRunAtMs = (entry as any)?.lastRunAtMs;
          if (typeof lastRunAtMs === "number" && Number.isFinite(lastRunAtMs)) {
            map[String(queryId)] = lastRunAtMs;
          }
        }
        setLastRunAtMsByQueryId(map);
      })
      .catch(() => {
        if (!cancelled) setLastRunAtMsByQueryId({});
      });

    return () => {
      cancelled = true;
    };
  }, [workbookId]);

  useEffect(() => {
    if (!service) {
      setQueries([]);
      return;
    }

    setQueries(service.getQueries());

    return service.onEvent((evt) => {
      if (evt?.type === "queries:changed") {
        const next = Array.isArray((evt as any).queries) ? (evt as any).queries : [];
        setQueries(next);
        return;
      }

      runtimeRef.current = reduceQueryRuntimeState(runtimeRef.current, evt);
      setRuntimeState(runtimeRef.current);

      const queryId = (evt as any)?.job?.queryId ?? (evt as any)?.queryId;
      if (typeof queryId === "string" && (evt as any)?.type === "completed") {
        const completedAt = (evt as any)?.job?.completedAt;
        const ms = completedAt instanceof Date && !Number.isNaN(completedAt.getTime()) ? completedAt.getTime() : Date.now();
        setLastRunAtMsByQueryId((prev) => ({ ...prev, [queryId]: ms }));
      }

      if (typeof queryId === "string") {
        const type = (evt as any)?.type;
        if (type === "completed") {
          const query = service.getQuery(queryId);
          const dest = query?.destination as any;
          const hasSheetDestination =
            dest &&
            typeof dest === "object" &&
            typeof dest.sheetId === "string" &&
            dest.start &&
            typeof dest.start === "object" &&
            typeof dest.start.row === "number" &&
            typeof dest.start.col === "number" &&
            typeof dest.includeHeader === "boolean";
          // If there's no destination, the refresh finishes at "completed" (no apply phase),
          // so clear the in-flight handle to allow subsequent manual refreshes.
          if (!hasSheetDestination) {
            activeRefreshHandleByQueryId.current.delete(queryId);
          }
        } else if (type === "error" || type === "cancelled" || type === "apply:completed" || type === "apply:error" || type === "apply:cancelled") {
          activeRefreshHandleByQueryId.current.delete(queryId);
        }
      }
    });
  }, [service]);

  const rows = useMemo(
    () => deriveQueryListRows(queries, runtimeState, lastRunAtMsByQueryId, { sheetNameResolver: props.sheetNameResolver }),
    [queries, runtimeState, lastRunAtMsByQueryId, props.sheetNameResolver],
  );

  const refreshQuery = useCallback(
    (queryId: string) => {
      setGlobalError(null);
      if (mutationsDisabled) return;
      if (!service) {
        setGlobalError("Power Query service not available.");
        return;
      }
      if (activeRefreshHandleByQueryId.current.has(queryId)) return;
      try {
        const handle = service.refreshWithDependencies(queryId);
        activeRefreshHandleByQueryId.current.set(queryId, { jobId: handle.id, cancel: handle.cancel });
      } catch (err) {
        setGlobalError(err instanceof Error ? err.message : String(err));
      }
    },
    [mutationsDisabled, service],
  );

  const cancelQueryRefresh = useCallback((queryId: string) => {
    activeRefreshHandleByQueryId.current.get(queryId)?.cancel();
  }, []);

  const refreshAll = useCallback(() => {
    setGlobalError(null);
    if (mutationsDisabled) return;
    if (!service) {
      setGlobalError("Power Query service not available.");
      return;
    }

    try {
      const handle = service.refreshAll();
      setActiveRefreshAll({ sessionId: handle.sessionId, cancel: handle.cancel });

      const cancelQuery = typeof (handle as any).cancelQuery === "function" ? ((handle as any).cancelQuery as (queryId: string) => void) : null;
      if (cancelQuery) {
        for (const query of queries) {
          activeRefreshHandleByQueryId.current.set(query.id, {
            jobId: handle.sessionId,
            cancel: () => cancelQuery(query.id),
          });
        }
      }

      handle.promise
        .finally(() => {
          setActiveRefreshAll(null);
          if (cancelQuery) {
            for (const query of queries) {
              const entry = activeRefreshHandleByQueryId.current.get(query.id);
              if (entry?.jobId === handle.sessionId) activeRefreshHandleByQueryId.current.delete(query.id);
            }
          }
        })
        .catch(() => {});
    } catch (err) {
      setGlobalError(err instanceof Error ? err.message : String(err));
    }
  }, [mutationsDisabled, queries, service]);

  const openInEditor = useCallback(
    (queryId: string) => {
      saveSelectedQueryId(workbookId, queryId);
      dispatchOpenPanel({ panelId: PanelIds.QUERY_EDITOR, queryId });
    },
    [workbookId],
  );

  const addNewQuery = useCallback(() => {
    setGlobalError(null);
    if (!service) {
      setGlobalError("Power Query service not available.");
      return;
    }
    const nextQuery = createBlankQuery({ name: `Query ${queries.length + 1}` });
    service.registerQuery(nextQuery);
  }, [queries.length, service]);

  const duplicateQuery = useCallback(
    (queryId: string) => {
      setGlobalError(null);
      if (!service) {
        setGlobalError("Power Query service not available.");
        return;
      }
      const existing = service.getQuery(queryId);
      if (!existing) return;
      service.registerQuery(duplicateQueryDefinition(existing));
    },
    [service],
  );

  const deleteQuery = useCallback(
    async (queryId: string) => {
      try {
        if (!service) {
          setGlobalError("Power Query service not available.");
          return;
        }
        const ok = await nativeDialogs.confirm("Delete this query?");
        if (!ok) return;
        setGlobalError(null);
        service.unregisterQuery(queryId);
      } catch (err: any) {
        setGlobalError(err?.message ?? String(err));
      }
    },
    [service],
  );

  const isOAuthSignedIn = useCallback(
    async (providerId: string, scopes: string[] | undefined): Promise<boolean> => {
      if (!credentialStore) return false;
      const normalized = normalizeOAuthScopes(providerId, scopes);
      const scope = oauth2Scope({ providerId, scopesHash: normalized.scopesHash });
      const entry = await credentialStore.get(scope as any);
      const secret = entry?.secret as any;
      return typeof secret?.refreshToken === "string" && secret.refreshToken.length > 0;
    },
    [credentialStore, normalizeOAuthScopes],
  );

  const [oauthSignedInByKey, setOauthSignedInByKey] = useState<Record<string, boolean>>({});
  const [sqlCredentialPresentByKey, setSqlCredentialPresentByKey] = useState<Record<string, boolean>>({});
  const [httpHeadersPresentByKey, setHttpHeadersPresentByKey] = useState<Record<string, boolean>>({});

  useEffect(() => {
    if (!credentialStore) {
      setOauthSignedInByKey({});
      setSqlCredentialPresentByKey({});
      setHttpHeadersPresentByKey({});
      return;
    }

    let cancelled = false;
    const run = async () => {
      const updates: Record<string, boolean> = {};
      const sqlUpdates: Record<string, boolean> = {};
      const httpUpdates: Record<string, boolean> = {};
      for (const query of queries) {
        const source = query.source as any;
        if (source?.type !== "api" || source?.auth?.type !== "oauth2") continue;
        const providerId = String(source.auth.providerId ?? "");
        if (!providerId) continue;
        const scopes = Array.isArray(source.auth.scopes) ? source.auth.scopes : undefined;
        const key = `${providerId}:${normalizeOAuthScopes(providerId, scopes).scopesHash}`;
        updates[key] = await isOAuthSignedIn(providerId, scopes);
      }
      for (const query of queries) {
        const source = query.source as any;
        if (source?.type !== "database") continue;
        const scope = resolveSqlScope(source.connection);
        if (!scope) continue;
        const key = sqlScopeKey(scope);
        try {
          sqlUpdates[key] = (await credentialStore.get(scope as any)) != null;
        } catch {
          sqlUpdates[key] = false;
        }
      }
      for (const query of queries) {
        const source = query.source as any;
        if (source?.type !== "api" || typeof source?.url !== "string") continue;
        const scope = safeHttpScope(source.url);
        if (!scope) continue;
        const key = httpScopeKey(scope);
        try {
          const entry = await credentialStore.get(scope as any);
          const secret = entry?.secret as any;
          const headers = secret && typeof secret === "object" && !Array.isArray(secret) ? (secret as any).headers : null;
          httpUpdates[key] = Boolean(headers && typeof headers === "object" && Object.keys(headers).length > 0);
        } catch {
          httpUpdates[key] = false;
        }
      }
      if (!cancelled) setOauthSignedInByKey(updates);
      if (!cancelled) setSqlCredentialPresentByKey(sqlUpdates);
      if (!cancelled) setHttpHeadersPresentByKey(httpUpdates);
    };

    void run();
    return () => {
      cancelled = true;
    };
  }, [credentialStore, isOAuthSignedIn, normalizeOAuthScopes, queries, resolveSqlScope]);

  const signOutOAuth = useCallback(
    async (providerId: string, scopes: string[] | undefined) => {
      if (!credentialStore) return;
      setGlobalError(null);
      const normalized = normalizeOAuthScopes(providerId, scopes);
      const scope = oauth2Scope({ providerId, scopesHash: normalized.scopesHash });
      try {
        await credentialStore.delete(scope as any);
      } catch (err) {
        setGlobalError(err instanceof Error ? err.message : String(err));
      }

      // Ensure sign-out takes effect immediately by evicting any cached tokens in
      // the in-memory OAuth2Manager. Otherwise the engine may continue using a
      // cached refresh/access token until the app reloads.
      try {
        const key = `${providerId}:${normalized.scopesHash}`;
        (oauth2Manager as any)?.cache?.delete?.(key);
        (oauth2Manager as any)?.inFlight?.delete?.(key);
      } catch {
        // ignore
      }

      setOauthSignedInByKey((prev) => {
        const next = { ...prev };
        delete next[`${providerId}:${normalized.scopesHash}`];
        return next;
      });
    },
    [credentialStore, normalizeOAuthScopes, oauth2Manager],
  );

  const setHttpHeaderCredential = useCallback(
    async (url: string) => {
      if (!credentialStore) return;
      setGlobalError(null);
      const headerName = (await showInputBox({ prompt: "Header name", value: "Authorization" }))?.trim();
      if (!headerName) return;
      const headerValue = (await showInputBox({ prompt: `Value for header '${headerName}'`, value: "" })) ?? "";
      if (!headerValue.trim()) return;
      try {
        const scope = safeHttpScope(url);
        if (!scope) throw new Error("Invalid URL");
        const key = httpScopeKey(scope);
        const existing = await credentialStore.get(scope as any);
        const existingSecret = existing?.secret as any;
        const nextSecret: any =
          existingSecret && typeof existingSecret === "object" && !Array.isArray(existingSecret) ? { ...existingSecret } : {};
        const existingHeaders =
          nextSecret.headers && typeof nextSecret.headers === "object" && !Array.isArray(nextSecret.headers)
            ? nextSecret.headers
            : {};
        nextSecret.headers = { ...existingHeaders, [headerName]: headerValue };
        await credentialStore.set(scope as any, nextSecret);
        setHttpHeadersPresentByKey((prev) => ({ ...prev, [key]: true }));
      } catch (err) {
        setGlobalError(err instanceof Error ? err.message : String(err));
      }
    },
    [credentialStore],
  );

  const clearHttpHeaderCredential = useCallback(
    async (url: string) => {
      if (!credentialStore) return;
      setGlobalError(null);
      try {
        const scope = safeHttpScope(url);
        if (!scope) return;
        const key = httpScopeKey(scope);
        const existing = await credentialStore.get(scope as any);
        const existingSecret = existing?.secret as any;
        if (existingSecret && typeof existingSecret === "object" && !Array.isArray(existingSecret)) {
          const nextSecret: any = { ...existingSecret };
          delete nextSecret.headers;
          if (Object.keys(nextSecret).length > 0) {
            await credentialStore.set(scope as any, nextSecret);
          } else {
            await credentialStore.delete(scope as any);
          }
        } else {
          await credentialStore.delete(scope as any);
        }
        setHttpHeadersPresentByKey((prev) => {
          const next = { ...prev };
          delete next[key];
          return next;
        });
      } catch (err) {
        setGlobalError(err instanceof Error ? err.message : String(err));
      }
    },
    [credentialStore],
  );

  const setSqlCredential = useCallback(
    async (connection: unknown) => {
      if (!credentialStore) return;
      setGlobalError(null);

      const scope = resolveSqlScope(connection);
      if (!scope) {
        setGlobalError("Unable to derive a stable credential scope for this database connection.");
        return;
      }

      const defaultUser = inferSqlUser(connection);
      const userValue = await showInputBox({ prompt: "Database user (optional)", value: defaultUser });
      if (userValue == null) return;
      const user = userValue.trim();

      const passwordValue = await showInputBox({ prompt: "Database password", value: "", type: "password" });
      if (passwordValue == null) return;
      const password = passwordValue;
      if (!password.trim()) return;

      try {
        await credentialStore.set(scope as any, user ? { user, password } : { password });
        const key = sqlScopeKey(scope);
        setSqlCredentialPresentByKey((prev) => ({ ...prev, [key]: true }));
      } catch (err) {
        setGlobalError(err instanceof Error ? err.message : String(err));
      }
    },
    [credentialStore, resolveSqlScope],
  );

  const clearSqlCredential = useCallback(
    async (connection: unknown) => {
      if (!credentialStore) return;
      setGlobalError(null);
      const scope = resolveSqlScope(connection);
      if (!scope) return;
      const key = sqlScopeKey(scope);
      try {
        await credentialStore.delete(scope as any);
        setSqlCredentialPresentByKey((prev) => {
          const next = { ...prev };
          delete next[key];
          return next;
        });
      } catch (err) {
        setGlobalError(err instanceof Error ? err.message : String(err));
      }
    },
    [credentialStore, resolveSqlScope],
  );

  const startOAuthSignIn = useCallback(
    async (providerId: string, scopes: string[] | undefined) => {
      if (!oauth2Manager) {
        setGlobalError("OAuth2 manager not available.");
        return;
      }

      setGlobalError(null);

      const config = oauthProviders.find((p) => p.id === providerId);
      if (!config) {
        setGlobalError(`OAuth provider '${providerId}' is not configured.`);
        return;
      }

      try {
        oauth2Manager.registerProvider(config as any);
      } catch (err) {
        setGlobalError(err instanceof Error ? err.message : String(err));
        return;
      }

      try {
        if (config.deviceAuthorizationEndpoint) {
          oauthBroker.setDeviceCodePromptHandler((code, verificationUri) => {
            setPendingDeviceCode({ providerId, code, verificationUri });
          });
          await oauth2Manager.authorizeWithDeviceCode({
            providerId,
            scopes,
            broker: oauthBroker as any,
          });
        } else if (config.authorizationEndpoint && config.redirectUri) {
          setPendingPkce({ providerId, redirectUri: config.redirectUri });
          const promise = oauth2Manager.authorizeWithPkce({
            providerId,
            scopes,
            broker: oauthBroker as any,
          });
          await promise;
        } else {
          setGlobalError(
            `OAuth provider '${providerId}' is missing deviceAuthorizationEndpoint or authorizationEndpoint+redirectUri.`,
          );
          return;
        }

        const normalized = normalizeScopes(scopes ?? (config as any)?.defaultScopes);
        setOauthSignedInByKey((prev) => ({ ...prev, [`${providerId}:${normalized.scopesHash}`]: true }));
      } catch (err) {
        if (err && typeof err === "object" && !Array.isArray(err) && (err as any).cancelled === true) {
          // User-cancelled PKCE wait; ignore.
        } else {
          setGlobalError(err instanceof Error ? err.message : String(err));
        }
      } finally {
        setPendingDeviceCode(null);
        setPendingPkce(null);
      }
    },
    [oauth2Manager, oauthProviders],
  );

  const saveProviderConfig = useCallback(
    (provider: OAuth2ProviderConfig) => {
      const next = [...oauthProviders.filter((p) => p.id !== provider.id), provider];
      setOauthProviders(next);
      saveOAuth2ProviderConfigs(workbookId, next);
      setEditingProvider(null);
    },
    [oauthProviders, workbookId],
  );

  const openProviderEditor = useCallback(
    (providerId: string) => {
      const existing = oauthProviders.find((p) => p.id === providerId);
      const draft =
        existing ??
        ({
          id: providerId,
          clientId: "",
          tokenEndpoint: "",
          ...(hasTauri() ? { redirectUri: RECOMMENDED_DESKTOP_OAUTH_REDIRECT_URI } : {}),
        } satisfies OAuth2ProviderConfig);
      setEditingProvider(draft);
    },
    [oauthProviders],
  );

  const closeProviderEditor = useCallback(() => setEditingProvider(null), []);

  const resolvePkceRedirect = useCallback(() => {
    void (async () => {
      if (!pendingPkce) return;
      if (supportsDesktopOAuthRedirectCapture(pendingPkce.redirectUri)) return;
      const redirectUrl = (await showInputBox({
        prompt: `Paste the full redirect URL (starts with ${pendingPkce.redirectUri})`,
        value: "",
        type: "textarea",
      }))?.trim();
      if (!redirectUrl) return;
      if (!matchesRedirectUri(pendingPkce.redirectUri, redirectUrl)) {
        setGlobalError(`Redirect URL does not match expected redirect URI (${pendingPkce.redirectUri}).`);
        return;
      }
      oauthBroker.resolveRedirect(pendingPkce.redirectUri, redirectUrl);
    })().catch(() => {});
  }, [pendingPkce]);

  const cancelPkceRedirect = useCallback(() => {
    if (!pendingPkce) return;
    oauthBroker.rejectRedirect(pendingPkce.redirectUri, { cancelled: true });
    setPendingPkce(null);
  }, [pendingPkce]);

  const renderProviderEditor = () => {
    if (!editingProvider) return null;
    const provider = editingProvider;

    return (
      <div className="data-queries-provider-editor">
        <div className="data-queries-provider-editor__title">Configure OAuth provider</div>
        <div className="data-queries-provider-editor__form">
          <label className="data-queries-provider-editor__label">
            Provider id
            <input
              value={provider.id}
              onChange={(e) => setEditingProvider({ ...provider, id: e.target.value })}
              className="data-queries-provider-editor__input"
              disabled={oauthProviders.some((p) => p.id === provider.id)}
            />
          </label>
          <label className="data-queries-provider-editor__label">
            Client ID
            <input
              value={provider.clientId}
              onChange={(e) => setEditingProvider({ ...provider, clientId: e.target.value })}
              className="data-queries-provider-editor__input"
            />
          </label>
          <label className="data-queries-provider-editor__label">
            Token endpoint
            <input
              value={provider.tokenEndpoint}
              onChange={(e) => setEditingProvider({ ...provider, tokenEndpoint: e.target.value })}
              className="data-queries-provider-editor__input"
            />
          </label>
          <label className="data-queries-provider-editor__label">
            Device authorization endpoint (optional)
            <input
              value={provider.deviceAuthorizationEndpoint ?? ""}
              onChange={(e) =>
                setEditingProvider({
                  ...provider,
                  deviceAuthorizationEndpoint: e.target.value || undefined,
                })
              }
              className="data-queries-provider-editor__input"
            />
          </label>
          <label className="data-queries-provider-editor__label">
            Authorization endpoint (optional, for PKCE)
            <input
              value={provider.authorizationEndpoint ?? ""}
              onChange={(e) =>
                setEditingProvider({
                  ...provider,
                  authorizationEndpoint: e.target.value || undefined,
                })
              }
              className="data-queries-provider-editor__input"
            />
          </label>
          <label className="data-queries-provider-editor__label">
            Redirect URI (optional, for PKCE)
            <input
              value={provider.redirectUri ?? ""}
              onChange={(e) => setEditingProvider({ ...provider, redirectUri: e.target.value || undefined })}
              className="data-queries-provider-editor__input"
            />
            {hasTauri() ? (
              <div className="data-queries-provider-editor__tip">
                Tip: Use <code>{RECOMMENDED_DESKTOP_OAUTH_REDIRECT_URI}</code> (preferred) or a loopback redirect like{" "}
                <code>{EXAMPLE_LOOPBACK_OAUTH_REDIRECT_URI}</code> to enable automatic redirect capture.
              </div>
            ) : null}
          </label>
          <div className="data-queries-provider-editor__buttons">
            <button
              type="button"
              onClick={() => saveProviderConfig(provider)}
              disabled={!provider.id.trim() || !provider.clientId.trim() || !provider.tokenEndpoint.trim()}
            >
              Save provider
            </button>
            <button type="button" onClick={closeProviderEditor}>
              Cancel
            </button>
          </div>
        </div>
      </div>
    );
  };

  if (!service) {
    return <div className="data-queries-panel-container__message">Power Query service not available.</div>;
  }

  const engineError = service.engineError;

  return (
    <div className="data-queries-panel-container">
      {engineError ? (
        <div className="data-queries-panel-container__engine-warning">Power Query engine running in fallback mode: {engineError}</div>
      ) : null}

      {renderProviderEditor()}

      <div className="data-queries-panel-container__toolbar">
        <button type="button" onClick={addNewQuery}>
          New query
        </button>
        <button type="button" onClick={refreshAll} disabled={queries.length === 0 || mutationsDisabled}>
          Refresh all
        </button>
        {activeRefreshAll ? (
          <button type="button" onClick={activeRefreshAll.cancel}>
            Cancel refresh all
          </button>
        ) : null}
        <div className="data-queries-panel-container__toolbar-spacer" />
        <button type="button" onClick={() => dispatchOpenPanel({ panelId: PanelIds.QUERY_EDITOR })}>
          Open Query Editor
        </button>
      </div>

      {pendingDeviceCode ? (
        <div className="data-queries-panel-container__banner">
          <div className="data-queries-panel-container__banner-title">OAuth device code</div>
          <div className="data-queries-panel-container__banner-text">
            {(() => {
              const href = safeHttpUrl(pendingDeviceCode.verificationUri);
              return (
                <>
                  Code: <code>{pendingDeviceCode.code}</code> • URL:{" "}
                  {href ? (
                    <a
                      href={href}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="data-queries-panel-container__mono-link"
                    >
                      {pendingDeviceCode.verificationUri}
                    </a>
                  ) : (
                    <code>{pendingDeviceCode.verificationUri}</code>
                  )}
                </>
              );
            })()}
          </div>
        </div>
      ) : null}

      {pendingPkce ? (
        <div className="data-queries-panel-container__banner">
          {supportsDesktopOAuthRedirectCapture(pendingPkce.redirectUri) ? (
            <>
              <div className="data-queries-panel-container__banner-title">Awaiting OAuth redirect…</div>
              <div className="data-queries-panel-container__banner-text">
                Complete sign-in in your browser. Formula will continue automatically after the redirect.
              </div>
              <div className="data-queries-panel-container__banner-actions data-queries-panel-container__banner-actions--mt8">
                <button type="button" onClick={cancelPkceRedirect}>
                  Cancel sign-in
                </button>
              </div>
            </>
          ) : (
            <>
              <div className="data-queries-panel-container__banner-title">OAuth redirect required</div>
              <div className="data-queries-panel-container__banner-text data-queries-panel-container__banner-text--mb8">
                After authenticating, copy the full redirect URL and paste it to complete sign-in.
              </div>
              <div className="data-queries-panel-container__banner-actions">
                <button type="button" onClick={resolvePkceRedirect}>
                  Paste redirect URL…
                </button>
                <button type="button" onClick={cancelPkceRedirect}>
                  Cancel sign-in
                </button>
              </div>
            </>
          )}
        </div>
      ) : null}

      {globalError ? (
        <div className="data-queries-panel-container__error">{globalError}</div>
      ) : null}

      <div className="data-queries-panel-container__table-wrap">
        <table className="data-queries-table">
          <thead>
            <tr>
              <th className="data-queries-table__th">Query</th>
              <th className="data-queries-table__th">Destination</th>
              <th className="data-queries-table__th">Last refresh</th>
              <th className="data-queries-table__th">Status</th>
              <th className="data-queries-table__th">Auth</th>
              <th className="data-queries-table__th">Error</th>
              <th className="data-queries-table__th">Actions</th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan={7} className="data-queries-table__empty">
                  No queries yet.
                </td>
              </tr>
            ) : (
              rows.map((row) => {
                const query = queries.find((q) => q.id === row.id) as Query | undefined;
                const source = query?.source as any;
                const apiUrl = source?.type === "api" && typeof source?.url === "string" ? String(source.url) : null;
                const httpScopeObj = apiUrl ? safeHttpScope(apiUrl) : null;
                const httpKey = httpScopeObj ? httpScopeKey(httpScopeObj) : null;
                const httpHasHeaders = httpKey ? Boolean(httpHeadersPresentByKey[httpKey]) : false;
                const databaseConnection = source?.type === "database" ? source.connection : null;
                const databaseScope = databaseConnection != null ? resolveSqlScope(databaseConnection) : null;
                const databaseKey = databaseScope ? sqlScopeKey(databaseScope) : null;
                const databaseHasCredential = databaseKey ? sqlCredentialPresentByKey[databaseKey] : false;
                const oauth =
                  source?.type === "api" && source?.auth?.type === "oauth2"
                    ? {
                        providerId: String(source.auth.providerId ?? ""),
                        scopes: Array.isArray(source.auth.scopes) ? source.auth.scopes : undefined,
                      }
                    : null;

                const oauthKey = oauth ? `${oauth.providerId}:${normalizeOAuthScopes(oauth.providerId, oauth.scopes).scopesHash}` : null;
                const signedIn = oauthKey ? oauthSignedInByKey[oauthKey] : false;

                const canCancel = row.status === "queued" || row.status === "refreshing" || row.status === "applying";
                const statusLabel =
                  row.status === "applying" && typeof row.rowsWritten === "number" ? `Applying (${row.rowsWritten} rows…)` : row.status;

                return (
                  <tr key={row.id} className="data-queries-table__row">
                    <td className="data-queries-table__td data-queries-table__query-name">{row.name}</td>
                    <td className="data-queries-table__td data-queries-table__td--mono data-queries-table__td--small">
                      {row.destination}
                    </td>
                    <td className="data-queries-table__td data-queries-table__td--small">{formatTimestamp(row.lastRefreshAtMs)}</td>
                    <td className="data-queries-table__td data-queries-table__td--small">{statusLabel}</td>
                    <td className="data-queries-table__td data-queries-table__td--small">
                      {oauth ? (
                        <div className="data-queries-auth">
                          <div>{row.authLabel}</div>
                          <div className="data-queries-auth__buttons">
                            {signedIn ? (
                              <button type="button" onClick={() => void signOutOAuth(oauth.providerId, oauth.scopes)}>
                                Sign out
                              </button>
                            ) : (
                              <button type="button" onClick={() => void startOAuthSignIn(oauth.providerId, oauth.scopes)}>
                                Sign in
                              </button>
                            )}
                            <button type="button" onClick={() => openProviderEditor(oauth.providerId)}>
                              Configure
                            </button>
                          </div>
                        </div>
                      ) : apiUrl ? (
                        <div className="data-queries-auth">
                          <div className="data-queries-table__muted">HTTP headers</div>
                          <div className="data-queries-auth__buttons">
                             <button type="button" onClick={() => void setHttpHeaderCredential(apiUrl)}>
                               {httpHasHeaders ? "Edit…" : "Set…"}
                             </button>
                             <button type="button" onClick={() => void clearHttpHeaderCredential(apiUrl)} disabled={!httpHasHeaders}>
                                Clear
                              </button>
                            </div>
                          </div>
                      ) : databaseConnection != null ? (
                        <div className="data-queries-auth">
                          <div className="data-queries-table__muted">{row.authLabel ?? "Database"}</div>
                          {databaseScope ? (
                            <div className="data-queries-auth__buttons">
                                {databaseHasCredential ? (
                                  <button type="button" onClick={() => void clearSqlCredential(databaseConnection)}>
                                    Clear
                                  </button>
                                ) : (
                                  <button type="button" onClick={() => void setSqlCredential(databaseConnection)}>
                                    Set…
                                  </button>
                                )}
                            </div>
                          ) : (
                            <div className="data-queries-table__muted">Credentials: managed on refresh</div>
                          )}
                        </div>
                      ) : row.authRequired ? (
                        row.authLabel
                      ) : (
                        "—"
                      )}
                    </td>
                    <td
                      className={[
                        "data-queries-table__td",
                        "data-queries-table__td--small",
                        "data-queries-table__error",
                        row.errorSummary ? "data-queries-table__error--present" : null,
                      ]
                        .filter(Boolean)
                        .join(" ")}
                    >
                      {row.errorSummary ?? "—"}
                    </td>
                    <td className="data-queries-table__td">
                      <div className="data-queries-actions">
                        <button type="button" onClick={() => refreshQuery(row.id)} disabled={canCancel || mutationsDisabled}>
                          Refresh
                        </button>
                        {canCancel ? (
                          <button type="button" onClick={() => cancelQueryRefresh(row.id)}>
                            Cancel
                          </button>
                        ) : null}
                        <button type="button" onClick={() => openInEditor(row.id)}>
                          Open
                        </button>
                        <button type="button" onClick={() => duplicateQuery(row.id)}>
                          Duplicate
                        </button>
                        <button type="button" onClick={() => deleteQuery(row.id)}>
                          Delete
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
