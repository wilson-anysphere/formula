import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";

import type { Query } from "../../../../../packages/power-query/src/model.js";
import { randomId } from "../../../../../packages/power-query/src/credentials/utils.js";
import { httpScope, oauth2Scope } from "../../../../../packages/power-query/src/credentials/scopes.js";
import { normalizeScopes } from "../../../../../packages/power-query/src/oauth2/tokenStore.js";

import { oauthBroker } from "../../power-query/oauthBroker.js";
import { createPowerQueryRefreshStateStore } from "../../power-query/refreshStateStore.js";
import { loadOAuth2ProviderConfigs, saveOAuth2ProviderConfigs, type OAuth2ProviderConfig } from "../../power-query/oauthProviders.ts";
import { deriveQueryListRows, reduceQueryRuntimeState, type QueryRuntimeState } from "../../power-query/queryRuntime.ts";
import {
  DesktopPowerQueryService,
  getDesktopPowerQueryService,
  onDesktopPowerQueryServiceChanged,
} from "../../power-query/service.js";

import { PanelIds } from "../panelRegistry.js";

type Props = {
  getDocumentController: () => any;
  workbookId?: string;
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

function hasTauri(): boolean {
  return Boolean((globalThis as any).__TAURI__);
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

  const [queries, setQueries] = useState<Query[]>(() => service?.getQueries?.() ?? []);
  const [runtimeState, setRuntimeState] = useState<QueryRuntimeState>({});
  const runtimeRef = useRef<QueryRuntimeState>({});

  const [lastRunAtMsByQueryId, setLastRunAtMsByQueryId] = useState<Record<string, number>>({});

  const [oauthProviders, setOauthProviders] = useState<OAuth2ProviderConfig[]>(() => loadOAuth2ProviderConfigs(workbookId));
  const [editingProvider, setEditingProvider] = useState<OAuth2ProviderConfig | null>(null);
  const [pendingPkce, setPendingPkce] = useState<PendingPkce | null>(null);
  const [pendingDeviceCode, setPendingDeviceCode] = useState<PendingDeviceCode | null>(null);
  const [globalError, setGlobalError] = useState<string | null>(null);

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

      if (
        typeof queryId === "string" &&
        ((evt as any)?.type === "error" ||
          (evt as any)?.type === "cancelled" ||
          (evt as any)?.type === "apply:completed" ||
          (evt as any)?.type === "apply:error" ||
          (evt as any)?.type === "apply:cancelled")
      ) {
        activeRefreshHandleByQueryId.current.delete(queryId);
      }
    });
  }, [service]);

  const rows = useMemo(
    () => deriveQueryListRows(queries, runtimeState, lastRunAtMsByQueryId),
    [queries, runtimeState, lastRunAtMsByQueryId],
  );

  const refreshQuery = useCallback(
    (queryId: string) => {
      setGlobalError(null);
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
    [service],
  );

  const cancelQueryRefresh = useCallback((queryId: string) => {
    activeRefreshHandleByQueryId.current.get(queryId)?.cancel();
  }, []);

  const refreshAll = useCallback(() => {
    setGlobalError(null);
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
  }, [queries, service]);

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
    (queryId: string) => {
      if (!service) {
        setGlobalError("Power Query service not available.");
        return;
      }
      if (typeof window !== "undefined" && typeof window.confirm === "function") {
        if (!window.confirm("Delete this query?")) return;
      }
      setGlobalError(null);
      service.unregisterQuery(queryId);
    },
    [service],
  );

  const isOAuthSignedIn = useCallback(
    async (providerId: string, scopes: string[] | undefined): Promise<boolean> => {
      if (!credentialStore) return false;
      const normalized = normalizeScopes(scopes);
      const scope = oauth2Scope({ providerId, scopesHash: normalized.scopesHash });
      const entry = await credentialStore.get(scope as any);
      const secret = entry?.secret as any;
      return typeof secret?.refreshToken === "string" && secret.refreshToken.length > 0;
    },
    [credentialStore],
  );

  const [oauthSignedInByKey, setOauthSignedInByKey] = useState<Record<string, boolean>>({});

  useEffect(() => {
    if (!credentialStore) {
      setOauthSignedInByKey({});
      return;
    }

    let cancelled = false;
    const run = async () => {
      const updates: Record<string, boolean> = {};
      for (const query of queries) {
        const source = query.source as any;
        if (source?.type !== "api" || source?.auth?.type !== "oauth2") continue;
        const providerId = String(source.auth.providerId ?? "");
        if (!providerId) continue;
        const scopes = Array.isArray(source.auth.scopes) ? source.auth.scopes : undefined;
        const key = `${providerId}:${normalizeScopes(scopes).scopesHash}`;
        updates[key] = await isOAuthSignedIn(providerId, scopes);
      }
      if (!cancelled) setOauthSignedInByKey(updates);
    };

    void run();
    return () => {
      cancelled = true;
    };
  }, [credentialStore, isOAuthSignedIn, queries]);

  const signOutOAuth = useCallback(
    async (providerId: string, scopes: string[] | undefined) => {
      if (!credentialStore) return;
      setGlobalError(null);
      const normalized = normalizeScopes(scopes);
      const scope = oauth2Scope({ providerId, scopesHash: normalized.scopesHash });
      try {
        await credentialStore.delete(scope as any);
      } catch (err) {
        setGlobalError(err instanceof Error ? err.message : String(err));
      }
      setOauthSignedInByKey((prev) => {
        const next = { ...prev };
        delete next[`${providerId}:${normalized.scopesHash}`];
        return next;
      });
    },
    [credentialStore],
  );

  const setHttpHeaderCredential = useCallback(
    async (url: string) => {
      if (!credentialStore) return;
      if (typeof window === "undefined" || typeof window.prompt !== "function") return;
      setGlobalError(null);
      const headerName = window.prompt("Header name", "Authorization")?.trim();
      if (!headerName) return;
      const headerValue = window.prompt(`Value for header '${headerName}'`, "") ?? "";
      if (!headerValue.trim()) return;
      try {
        const scope = httpScope({ url, realm: null });
        await credentialStore.set(scope as any, { headers: { [headerName]: headerValue } });
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
        const scope = httpScope({ url, realm: null });
        await credentialStore.delete(scope as any);
      } catch (err) {
        setGlobalError(err instanceof Error ? err.message : String(err));
      }
    },
    [credentialStore],
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

        const normalized = normalizeScopes(scopes);
        setOauthSignedInByKey((prev) => ({ ...prev, [`${providerId}:${normalized.scopesHash}`]: true }));
      } catch (err) {
        setGlobalError(err instanceof Error ? err.message : String(err));
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
      const existing =
        oauthProviders.find((p) => p.id === providerId) ??
        ({
          id: providerId,
          clientId: "",
          tokenEndpoint: "",
        } satisfies OAuth2ProviderConfig);
      setEditingProvider(existing);
    },
    [oauthProviders],
  );

  const closeProviderEditor = useCallback(() => setEditingProvider(null), []);

  const resolvePkceRedirect = useCallback(() => {
    if (!pendingPkce) return;
    if (typeof window === "undefined" || typeof window.prompt !== "function") return;
    const redirectUrl = window.prompt(`Paste the full redirect URL (starts with ${pendingPkce.redirectUri})`, "");
    if (!redirectUrl) return;
    oauthBroker.resolveRedirect(pendingPkce.redirectUri, redirectUrl);
  }, [pendingPkce]);

  const renderProviderEditor = () => {
    if (!editingProvider) return null;
    const provider = editingProvider;

    return (
      <div style={{ padding: 12, borderBottom: "1px solid var(--border)", background: "var(--bg-secondary)" }}>
        <div style={{ fontWeight: 600, marginBottom: 8 }}>Configure OAuth provider</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          <label style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12 }}>
            Provider id
            <input
              value={provider.id}
              onChange={(e) => setEditingProvider({ ...provider, id: e.target.value })}
              style={{ padding: 6 }}
              disabled={oauthProviders.some((p) => p.id === provider.id)}
            />
          </label>
          <label style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12 }}>
            Client ID
            <input
              value={provider.clientId}
              onChange={(e) => setEditingProvider({ ...provider, clientId: e.target.value })}
              style={{ padding: 6 }}
            />
          </label>
          <label style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12 }}>
            Token endpoint
            <input
              value={provider.tokenEndpoint}
              onChange={(e) => setEditingProvider({ ...provider, tokenEndpoint: e.target.value })}
              style={{ padding: 6 }}
            />
          </label>
          <label style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12 }}>
            Device authorization endpoint (optional)
            <input
              value={provider.deviceAuthorizationEndpoint ?? ""}
              onChange={(e) =>
                setEditingProvider({
                  ...provider,
                  deviceAuthorizationEndpoint: e.target.value || undefined,
                })
              }
              style={{ padding: 6 }}
            />
          </label>
          <label style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12 }}>
            Authorization endpoint (optional, for PKCE)
            <input
              value={provider.authorizationEndpoint ?? ""}
              onChange={(e) =>
                setEditingProvider({
                  ...provider,
                  authorizationEndpoint: e.target.value || undefined,
                })
              }
              style={{ padding: 6 }}
            />
          </label>
          <label style={{ display: "flex", flexDirection: "column", gap: 4, fontSize: 12 }}>
            Redirect URI (optional, for PKCE)
            <input
              value={provider.redirectUri ?? ""}
              onChange={(e) => setEditingProvider({ ...provider, redirectUri: e.target.value || undefined })}
              style={{ padding: 6 }}
            />
          </label>
          <div style={{ display: "flex", gap: 8 }}>
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
    return <div style={{ padding: 12, color: "var(--text-muted)" }}>Power Query service not available.</div>;
  }

  const engineError = service.engineError;

  return (
    <div style={{ flex: 1, minHeight: 0, display: "flex", flexDirection: "column" }}>
      {engineError ? (
        <div style={{ padding: 12, color: "var(--text-muted)", borderBottom: "1px solid var(--border)" }}>
          Power Query engine running in fallback mode: {engineError}
        </div>
      ) : null}

      {renderProviderEditor()}

      <div style={{ padding: 12, borderBottom: "1px solid var(--border)", display: "flex", flexWrap: "wrap", gap: 8 }}>
        <button type="button" onClick={addNewQuery}>
          New query
        </button>
        <button type="button" onClick={refreshAll} disabled={queries.length === 0}>
          Refresh all
        </button>
        {activeRefreshAll ? (
          <button type="button" onClick={activeRefreshAll.cancel}>
            Cancel refresh all
          </button>
        ) : null}
        <div style={{ flex: 1 }} />
        <button type="button" onClick={() => dispatchOpenPanel({ panelId: PanelIds.QUERY_EDITOR })}>
          Open Query Editor
        </button>
      </div>

      {pendingDeviceCode ? (
        <div style={{ padding: 12, borderBottom: "1px solid var(--border)", background: "var(--bg-secondary)" }}>
          <div style={{ fontWeight: 600, marginBottom: 6 }}>OAuth device code</div>
          <div style={{ fontSize: 12, opacity: 0.85 }}>
            Code: <code>{pendingDeviceCode.code}</code> • URL: <code>{pendingDeviceCode.verificationUri}</code>
          </div>
        </div>
      ) : null}

      {pendingPkce ? (
        <div style={{ padding: 12, borderBottom: "1px solid var(--border)", background: "var(--bg-secondary)" }}>
          <div style={{ fontWeight: 600, marginBottom: 6 }}>OAuth redirect required</div>
          <div style={{ fontSize: 12, opacity: 0.85, marginBottom: 8 }}>
            After authenticating, copy the full redirect URL and paste it to complete sign-in.
          </div>
          <button type="button" onClick={resolvePkceRedirect}>
            Paste redirect URL…
          </button>
        </div>
      ) : null}

      {globalError ? (
        <div style={{ padding: 12, color: "var(--error)", borderBottom: "1px solid var(--border)" }}>{globalError}</div>
      ) : null}

      <div style={{ flex: 1, minHeight: 0, overflow: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead>
            <tr style={{ textAlign: "left", borderBottom: "1px solid var(--border)" }}>
              <th style={{ padding: 8 }}>Query</th>
              <th style={{ padding: 8 }}>Destination</th>
              <th style={{ padding: 8 }}>Last refresh</th>
              <th style={{ padding: 8 }}>Status</th>
              <th style={{ padding: 8 }}>Auth</th>
              <th style={{ padding: 8 }}>Error</th>
              <th style={{ padding: 8 }}>Actions</th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan={7} style={{ padding: 12, color: "var(--text-muted)" }}>
                  No queries yet.
                </td>
              </tr>
            ) : (
              rows.map((row) => {
                const query = queries.find((q) => q.id === row.id) as Query | undefined;
                const source = query?.source as any;
                const apiUrl = source?.type === "api" && typeof source?.url === "string" ? String(source.url) : null;
                const oauth =
                  source?.type === "api" && source?.auth?.type === "oauth2"
                    ? {
                        providerId: String(source.auth.providerId ?? ""),
                        scopes: Array.isArray(source.auth.scopes) ? source.auth.scopes : undefined,
                      }
                    : null;

                const oauthKey = oauth ? `${oauth.providerId}:${normalizeScopes(oauth.scopes).scopesHash}` : null;
                const signedIn = oauthKey ? oauthSignedInByKey[oauthKey] : false;

                const canCancel = row.status === "queued" || row.status === "refreshing" || row.status === "applying";
                const statusLabel =
                  row.status === "applying" && typeof row.rowsWritten === "number" ? `Applying (${row.rowsWritten} rows…)` : row.status;

                return (
                  <tr key={row.id} style={{ borderBottom: "1px solid var(--border)" }}>
                    <td style={{ padding: 8, fontWeight: 600 }}>{row.name}</td>
                    <td style={{ padding: 8, fontFamily: "monospace", fontSize: 12 }}>{row.destination}</td>
                    <td style={{ padding: 8, fontSize: 12 }}>{formatTimestamp(row.lastRefreshAtMs)}</td>
                    <td style={{ padding: 8, fontSize: 12 }}>{statusLabel}</td>
                    <td style={{ padding: 8, fontSize: 12 }}>
                      {oauth ? (
                        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                          <div>{row.authLabel}</div>
                          <div style={{ display: "flex", gap: 6 }}>
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
                        <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                          <div style={{ color: "var(--text-muted)" }}>HTTP headers</div>
                          <div style={{ display: "flex", gap: 6 }}>
                            <button type="button" onClick={() => void setHttpHeaderCredential(apiUrl)}>
                              Set…
                            </button>
                            <button type="button" onClick={() => void clearHttpHeaderCredential(apiUrl)}>
                              Clear
                            </button>
                          </div>
                        </div>
                      ) : row.authRequired ? (
                        row.authLabel
                      ) : (
                        "—"
                      )}
                    </td>
                    <td style={{ padding: 8, fontSize: 12, color: row.errorSummary ? "var(--error)" : "var(--text-muted)" }}>
                      {row.errorSummary ?? "—"}
                    </td>
                    <td style={{ padding: 8 }}>
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                        <button type="button" onClick={() => refreshQuery(row.id)} disabled={canCancel}>
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
