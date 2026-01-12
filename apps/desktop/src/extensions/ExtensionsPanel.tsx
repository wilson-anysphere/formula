import React from "react";

import type { DesktopExtensionHostManager } from "./extensionHostManager.js";
import { buildCommandKeybindingDisplayIndex, getPrimaryCommandKeybindingDisplay } from "./keybindings.js";
import { showInputBox, showToast } from "./ui.js";

type NetworkPolicy = {
  mode: "full" | "deny" | "allowlist" | string;
  hosts?: string[];
};

type GrantedPermissions = Record<string, true | NetworkPolicy>;

type ContributedCommand = {
  extensionId: string;
  command: string;
  title: string;
  category: string | null;
};

type ContributedPanel = {
  extensionId: string;
  id: string;
  title: string;
};

function groupByExtension<T extends { extensionId: string }>(items: T[]): Map<string, T[]> {
  const map = new Map<string, T[]>();
  for (const item of items) {
    const list = map.get(item.extensionId) ?? [];
    list.push(item);
    map.set(item.extensionId, list);
  }
  return map;
}

function normalizeStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return value
    .filter((entry) => typeof entry === "string")
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0);
}

export function ExtensionsPanel({
  manager,
  onExecuteCommand,
  onOpenPanel,
}: {
  manager: DesktopExtensionHostManager;
  onExecuteCommand: (commandId: string, ...args: any[]) => Promise<unknown> | void;
  onOpenPanel: (panelId: string) => void;
}) {
  const [, bump] = React.useState(0);
  React.useEffect(() => manager.subscribe(() => bump((v) => v + 1)), [manager]);

  const [permissionsByExtension, setPermissionsByExtension] = React.useState<Record<string, GrantedPermissions>>({});
  const [permissionsError, setPermissionsError] = React.useState<string | null>(null);
  const [busy, setBusy] = React.useState<string | null>(null);

  const extensions = manager.host.listExtensions();
  const commands = manager.getContributedCommands() as ContributedCommand[];
  const panels = manager.getContributedPanels() as ContributedPanel[];
  const keybindings = manager.getContributedKeybindings() as Array<{ command: string; key: string; mac?: string | null }>;
  const platform = typeof navigator !== "undefined" && /Mac|iPhone|iPad|iPod/.test(navigator.platform) ? "mac" : "other";
  const keybindingIndex = buildCommandKeybindingDisplayIndex({ platform, contributed: keybindings });

  const commandsByExt = groupByExtension(commands);
  const panelsByExt = groupByExtension(panels);

  const refreshPermissions = React.useCallback(async () => {
    if (!manager.ready || manager.error) return;
    setPermissionsError(null);
    try {
      const exts = manager.host.listExtensions();
      const entries = await Promise.all(
        exts.map(async (ext: any) => [ext.id, (await manager.getGrantedPermissions(ext.id)) as GrantedPermissions] as const),
      );
      setPermissionsByExtension(Object.fromEntries(entries));
    } catch (err) {
      setPermissionsError(String((err as any)?.message ?? err));
    }
  }, [manager]);

  const refreshPermissionsForExtension = React.useCallback(
    async (extensionId: string) => {
      if (!manager.ready || manager.error) return;
      setPermissionsError(null);
      try {
        const data = (await manager.getGrantedPermissions(extensionId)) as GrantedPermissions;
        setPermissionsByExtension((prev) => ({ ...prev, [extensionId]: data }));
      } catch (err) {
        setPermissionsError(String((err as any)?.message ?? err));
      }
    },
    [manager],
  );

  const extensionsKey = extensions.map((ext: any) => String(ext.id)).join("|");

  React.useEffect(() => {
    void refreshPermissions();
  }, [refreshPermissions, manager.ready, manager.error, extensionsKey]);

  const executeCommandAndRefreshPermissions = React.useCallback(
    async (extensionId: string, commandId: string, args: any[] = []) => {
      try {
        await Promise.resolve(onExecuteCommand(commandId, ...args));
      } finally {
        await refreshPermissionsForExtension(extensionId);
      }
    },
    [onExecuteCommand, refreshPermissionsForExtension],
  );

  const runCommandWithArgs = React.useCallback(
    (extensionId: string, commandId: string) => {
      void (async () => {
        const raw = await showInputBox({
          prompt: "Command arguments (JSON array)",
          value: "[]",
          placeHolder: 'Example: ["https://example.com/"]',
        });
        if (raw == null) return;

        let parsed: unknown;
        const trimmed = raw.trim();
        try {
          parsed = trimmed.length === 0 ? [] : JSON.parse(trimmed);
        } catch (err) {
          showToast(`Invalid JSON: ${String((err as any)?.message ?? err)}`, "error");
          return;
        }

        const args = Array.isArray(parsed) ? parsed : [parsed];
        await executeCommandAndRefreshPermissions(extensionId, commandId, args);
      })().catch((err) => {
        showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
      });
    },
    [executeCommandAndRefreshPermissions],
  );

  if (!manager.ready) {
    return <div>Loading extensions…</div>;
  }

  if (manager.error) {
    return <div>Failed to load extensions: {String((manager.error as any)?.message ?? manager.error)}</div>;
  }

  if (extensions.length === 0) {
    return <div>No extensions installed.</div>;
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: "14px" }}>
      <div style={{ display: "flex", justifyContent: "flex-end", gap: "8px" }}>
        <button
          type="button"
          data-testid="reset-all-extension-permissions"
          disabled={busy === "reset-all"}
          onClick={() => {
            void (async () => {
              setBusy("reset-all");
              try {
                await manager.resetAllPermissions();
                await refreshPermissions();
              } catch (err) {
                showToast(`Failed to reset permissions: ${String((err as any)?.message ?? err)}`, "error");
              } finally {
                setBusy(null);
              }
            })().catch(() => {
              // ignore
            });
          }}
          style={{
            padding: "8px 10px",
            borderRadius: "10px",
            border: "1px solid var(--border)",
            background: "var(--bg-secondary)",
            color: "var(--text-primary)",
            cursor: "pointer",
          }}
        >
          Reset all extensions permissions
        </button>
        <button
          type="button"
          data-testid="refresh-extension-permissions"
          disabled={busy === "refresh"}
          onClick={() => {
            void (async () => {
              setBusy("refresh");
              try {
                await refreshPermissions();
              } catch (err) {
                showToast(`Failed to refresh permissions: ${String((err as any)?.message ?? err)}`, "error");
              } finally {
                setBusy(null);
              }
            })().catch(() => {
              // ignore
            });
          }}
          style={{
            padding: "8px 10px",
            borderRadius: "10px",
            border: "1px solid var(--border)",
            background: "var(--bg-secondary)",
            color: "var(--text-primary)",
            cursor: "pointer",
          }}
        >
          Refresh
        </button>
      </div>

      {permissionsError ? (
        <div style={{ color: "var(--text-secondary)", fontSize: "12px" }}>Failed to load permissions: {permissionsError}</div>
      ) : null}

      {extensions.map((ext: any) => {
        const extCommands = commandsByExt.get(ext.id) ?? [];
        const extPanels = panelsByExt.get(ext.id) ?? [];

        const declaredPermissions = normalizeStringArray(ext.manifest?.permissions);
        const declaredSet = new Set(declaredPermissions);

        const granted = permissionsByExtension[ext.id] ?? {};
        const networkPolicy = (granted as any).network ?? null;
        const networkMode = typeof networkPolicy?.mode === "string" ? String(networkPolicy.mode) : null;
        const networkHosts = normalizeStringArray(networkPolicy?.hosts);

        const allPermissions = (() => {
          const set = new Set<string>();
          for (const perm of declaredPermissions) set.add(perm);
          for (const perm of Object.keys(granted ?? {})) set.add(perm);
          return [...set].sort();
        })();

        return (
          <div
            key={ext.id}
            data-testid={`extension-card-${ext.id}`}
            style={{
              border: "1px solid var(--panel-border)",
              borderRadius: "10px",
              background: "var(--bg-primary)",
              padding: "12px",
            }}
          >
            <div style={{ fontWeight: 700, marginBottom: "8px" }}>{ext.manifest?.displayName ?? ext.id}</div>
            <div style={{ fontSize: "12px", color: "var(--text-secondary)", marginBottom: "10px" }}>{ext.id}</div>

            <div style={{ marginBottom: "12px" }}>
              <div style={{ fontWeight: 600, marginBottom: "6px" }}>Permissions</div>

              <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                {allPermissions.length === 0 ? (
                  <div style={{ color: "var(--text-secondary)" }}>No permissions declared.</div>
                ) : (
                  allPermissions.map((perm) => {
                    const declared = declaredSet.has(perm);
                    const isNetwork = perm === "network";
                    const grantedValue = isNetwork
                      ? Boolean(networkPolicy) && networkMode !== "deny"
                      : (granted as any)?.[perm] === true;

                    const subtitle = isNetwork
                      ? networkPolicy
                        ? `mode: ${networkMode ?? "unknown"}${
                            networkMode === "allowlist" && networkHosts.length > 0 ? ` • hosts: ${networkHosts.join(", ")}` : ""
                          }`
                        : "not granted"
                      : grantedValue
                        ? "granted"
                        : "not granted";

                    return (
                      <div
                        key={perm}
                        data-testid={`permission-row-${ext.id}-${perm}`}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "space-between",
                          gap: "10px",
                          padding: "10px 12px",
                          borderRadius: "10px",
                          border: "1px solid var(--border)",
                          background: "var(--bg-secondary)",
                        }}
                      >
                        <div style={{ minWidth: 0 }}>
                          <div style={{ fontWeight: 600, display: "flex", alignItems: "center", gap: "8px" }}>
                            <span>{perm}</span>
                            <span style={{ fontSize: "11px", color: "var(--text-secondary)" }}>
                              {declared ? "declared" : "not declared"}
                            </span>
                          </div>
                          <div
                            style={{
                              fontSize: "12px",
                              color: "var(--text-secondary)",
                              overflow: "hidden",
                              textOverflow: "ellipsis",
                            }}
                          >
                            {subtitle}
                          </div>
                        </div>

                        <button
                          type="button"
                          data-testid={`revoke-permission-${ext.id}-${perm}`}
                          disabled={!grantedValue || busy === `revoke:${ext.id}:${perm}`}
                          onClick={() => {
                            void (async () => {
                              const key = `revoke:${ext.id}:${perm}`;
                              setBusy(key);
                              try {
                                await manager.revokePermission(ext.id, perm);
                                await refreshPermissionsForExtension(ext.id);
                              } catch (err) {
                                showToast(
                                  `Failed to revoke permission '${perm}': ${String((err as any)?.message ?? err)}`,
                                  "error",
                                );
                              } finally {
                                setBusy(null);
                              }
                            })().catch(() => {
                              // ignore
                            });
                          }}
                          style={{
                            flex: "0 0 auto",
                            padding: "8px 10px",
                            borderRadius: "10px",
                            border: "1px solid var(--border)",
                            background: "var(--bg-primary)",
                            color: grantedValue ? "var(--text-primary)" : "var(--text-secondary)",
                            cursor: grantedValue ? "pointer" : "not-allowed",
                          }}
                        >
                          Revoke
                        </button>
                      </div>
                    );
                  })
                )}

                <div style={{ display: "flex", justifyContent: "flex-end", marginTop: "6px" }}>
                  <button
                    type="button"
                    data-testid={`reset-extension-permissions-${ext.id}`}
                    disabled={busy === `reset:${ext.id}`}
                    onClick={() => {
                      void (async () => {
                        const key = `reset:${ext.id}`;
                        setBusy(key);
                        try {
                          await manager.resetPermissionsForExtension(ext.id);
                          await refreshPermissionsForExtension(ext.id);
                        } catch (err) {
                          showToast(
                            `Failed to reset permissions for ${String(ext.manifest?.displayName ?? ext.id)}: ${String(
                              (err as any)?.message ?? err,
                            )}`,
                            "error",
                          );
                        } finally {
                          setBusy(null);
                        }
                      })().catch(() => {
                        // ignore
                      });
                    }}
                    style={{
                      padding: "8px 10px",
                      borderRadius: "10px",
                      border: "1px solid var(--border)",
                      background: "var(--bg-primary)",
                      color: "var(--text-primary)",
                      cursor: "pointer",
                    }}
                  >
                    Revoke all permissions
                  </button>
                </div>
              </div>
            </div>

            {extCommands.length > 0 ? (
              <>
                <div style={{ fontWeight: 600, marginBottom: "6px" }}>Commands</div>
                <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                  {extCommands.map((cmd) => (
                    <div key={cmd.command} style={{ display: "flex", gap: "8px", alignItems: "stretch" }}>
                      <button
                        type="button"
                        data-testid={`run-command-${cmd.command}`}
                        onClick={() => {
                          void executeCommandAndRefreshPermissions(String(ext.id), cmd.command).catch((err) => {
                            showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
                          });
                        }}
                        style={{
                          textAlign: "left",
                          padding: "10px 12px",
                          borderRadius: "10px",
                          border: "1px solid var(--border)",
                          background: "var(--bg-secondary)",
                          color: "var(--text-primary)",
                          cursor: "pointer",
                          flex: "1 1 auto",
                        }}
                      >
                        <div
                          style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: "12px" }}
                        >
                          <div style={{ fontWeight: 600, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis" }}>
                            {cmd.category ? `${cmd.category}: ` : ""}
                            {cmd.title}
                          </div>
                          {(() => {
                            const shortcut = getPrimaryCommandKeybindingDisplay(cmd.command, keybindingIndex);
                            if (!shortcut) return null;
                            return (
                              <div
                                aria-hidden="true"
                                style={{
                                  fontSize: "12px",
                                  color: "var(--text-secondary)",
                                  whiteSpace: "nowrap",
                                  fontFamily:
                                    "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono', 'Courier New', monospace",
                                }}
                              >
                                {shortcut}
                              </div>
                            );
                          })()}
                        </div>
                        <div style={{ fontSize: "12px", color: "var(--text-secondary)" }}>{cmd.command}</div>
                      </button>
                      <button
                        type="button"
                        data-testid={`run-command-with-args-${cmd.command}`}
                        onClick={() => runCommandWithArgs(String(ext.id), cmd.command)}
                        style={{
                          padding: "10px 12px",
                          borderRadius: "10px",
                          border: "1px solid var(--border)",
                          background: "var(--bg-secondary)",
                          color: "var(--text-primary)",
                          cursor: "pointer",
                          whiteSpace: "nowrap",
                        }}
                      >
                        Run…
                      </button>
                    </div>
                  ))}
                </div>
              </>
            ) : (
              <div style={{ color: "var(--text-secondary)" }}>No commands contributed.</div>
            )}

            {extPanels.length > 0 ? (
              <>
                <div style={{ fontWeight: 600, marginTop: "12px", marginBottom: "6px" }}>Panels</div>
                <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                  {extPanels.map((panel) => (
                    <button
                      key={panel.id}
                      type="button"
                      data-testid={`open-panel-${panel.id}`}
                      onClick={() => onOpenPanel(panel.id)}
                      style={{
                        textAlign: "left",
                        padding: "10px 12px",
                        borderRadius: "10px",
                        border: "1px solid var(--border)",
                        background: "var(--bg-secondary)",
                        color: "var(--text-primary)",
                        cursor: "pointer",
                      }}
                    >
                      <div style={{ fontWeight: 600 }}>{panel.title}</div>
                      <div style={{ fontSize: "12px", color: "var(--text-secondary)" }}>{panel.id}</div>
                    </button>
                  ))}
                </div>
              </>
            ) : null}
          </div>
        );
      })}
    </div>
  );
}
