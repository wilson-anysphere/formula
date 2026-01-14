import React from "react";

import type { WebExtensionManager } from "@formula/extension-marketplace";
import type { DesktopExtensionHostManager } from "./extensionHostManager.js";
import {
  buildCommandKeybindingDisplayIndex,
  getPrimaryCommandKeybindingDisplay,
  parseKeybinding,
  type ParsedKeybinding,
} from "./keybindings.js";
import { DEFAULT_RESERVED_EXTENSION_SHORTCUTS } from "./keybindingService.js";
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

const RESERVED_EXTENSION_KEYBINDINGS: ParsedKeybinding[] = DEFAULT_RESERVED_EXTENSION_SHORTCUTS.map((binding) =>
  parseKeybinding("__reserved__", binding, null),
).filter((binding): binding is ParsedKeybinding => binding != null);

function isReservedExtensionKeybinding(binding: ParsedKeybinding): boolean {
  return RESERVED_EXTENSION_KEYBINDINGS.some(
    (reserved) =>
      reserved.ctrl === binding.ctrl &&
      reserved.shift === binding.shift &&
      reserved.alt === binding.alt &&
      reserved.meta === binding.meta &&
      reserved.key === binding.key,
  );
}

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

function collectDeclaredPermissions(value: unknown): string[] {
  const list = Array.isArray(value) ? value : [];
  const out = new Set<string>();
  for (const entry of list) {
    if (typeof entry === "string") {
      const trimmed = entry.trim();
      if (trimmed) out.add(trimmed);
      continue;
    }
    if (entry && typeof entry === "object" && !Array.isArray(entry)) {
      for (const key of Object.keys(entry)) {
        const trimmed = String(key).trim();
        if (trimmed) out.add(trimmed);
      }
    }
  }
  return [...out].sort();
}

export function ExtensionsPanel({
  manager,
  webExtensionManager,
  onSyncExtensions,
  onExecuteCommand,
  onOpenPanel,
}: {
  manager: DesktopExtensionHostManager;
  webExtensionManager?: WebExtensionManager | null;
  onSyncExtensions?: (() => void) | null;
  onExecuteCommand: (commandId: string, ...args: any[]) => Promise<unknown> | void;
  onOpenPanel: (panelId: string) => void;
}) {
  const [, bump] = React.useState(0);
  React.useEffect(() => manager.subscribe(() => bump((v) => v + 1)), [manager]);

  // The desktop app defers loading built-in extensions until a user action triggers it (to avoid
  // spawning extra Workers during startup). If the Extensions panel is already open due to layout
  // persistence, we still want it to load extensions so the panel isn't stuck in the "Loading…"
  // state after reload.
  React.useEffect(() => {
    if (manager.ready) return;
    void manager.loadBuiltInExtensions().catch(() => {
      // Errors are surfaced via `manager.error` once the load attempt completes.
    });
  }, [manager]);

  const [permissionsByExtension, setPermissionsByExtension] = React.useState<Record<string, GrantedPermissions>>({});
  const [permissionsError, setPermissionsError] = React.useState<string | null>(null);
  const [busy, setBusy] = React.useState<string | null>(null);

  const [installed, setInstalled] = React.useState<Array<{
    id: string;
    version: string;
    corrupted?: boolean;
    corruptedReason?: string;
    incompatible?: boolean;
    incompatibleReason?: string;
  }> | null>(null);
  const [installError, setInstallError] = React.useState<string | null>(null);
  const [repairingId, setRepairingId] = React.useState<string | null>(null);

  const refreshInstalled = React.useCallback(async () => {
    if (!webExtensionManager) return;
    try {
      setInstallError(null);
      await webExtensionManager.verifyAllInstalled();
      const next = await webExtensionManager.listInstalled();
      setInstalled(next);
    } catch (err: any) {
      setInstallError(String(err?.message ?? err));
    }
  }, [webExtensionManager]);

  React.useEffect(() => {
    if (!webExtensionManager) return;
    void refreshInstalled();
  }, [refreshInstalled, webExtensionManager]);

  // Keep the marketplace-installed list in sync when other parts of the app mutate installs
  // (Marketplace panel install/uninstall/update). Those flows ultimately call
  // `DesktopExtensionHostManager.notifyDidChange()`, so subscribing here avoids requiring a reload.
  React.useEffect(() => {
    if (!webExtensionManager) return;
    return manager.subscribe(() => {
      void refreshInstalled();
    });
  }, [manager, refreshInstalled, webExtensionManager]);

  const extensions = manager.host.listExtensions();
  const commands = manager.getContributedCommands() as ContributedCommand[];
  const panels = manager.getContributedPanels() as ContributedPanel[];
  const platform = typeof navigator !== "undefined" && /Mac|iPhone|iPad|iPod/.test(navigator.platform) ? "mac" : "other";
  const keybindings = manager.getContributedKeybindings() as Array<{
    command: string;
    key: string;
    mac?: string | null;
    when?: string | null;
  }>;
  const filteredKeybindings = keybindings.filter((kb) => {
    const raw = platform === "mac" && kb.mac ? kb.mac : kb.key;
    const parsed = parseKeybinding(kb.command, raw, kb.when ?? null);
    if (!parsed) return false;
    return !isReservedExtensionKeybinding(parsed);
  });
  const keybindingIndex = buildCommandKeybindingDisplayIndex({ platform, contributed: filteredKeybindings });

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
    void refreshPermissions().catch(() => {});
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

  return (
    <div className="extensions-panel">
      {webExtensionManager ? (
          <div className="extensions-panel__card">
            <div className="extensions-panel__card-title">Installed (IndexedDB)</div>
            {installError ? <div className="extensions-panel__error">Installed extensions error: {installError}</div> : null}
            {installed && installed.length > 0 ? (
              <div className="extensions-panel__installed-list">
                {installed.map((item) => {
                  const isCorrupted = Boolean(item.corrupted);
                  const isIncompatible = Boolean(item.incompatible);
                  return (
                    <div key={item.id} data-testid={`installed-extension-${item.id}`} className="extensions-panel__installed-item">
                      <div className="extensions-panel__installed-item-id">{item.id}</div>
                      <div className="extensions-panel__installed-item-version">v{item.version}</div>
                      <div
                        data-testid={`installed-extension-status-${item.id}`}
                        className={
                          isCorrupted
                            ? "extensions-panel__installed-item-status extensions-panel__installed-item-status--corrupted"
                            : isIncompatible
                              ? "extensions-panel__installed-item-status extensions-panel__installed-item-status--incompatible"
                              : "extensions-panel__installed-item-status"
                        }
                      >
                        {isCorrupted
                          ? `Corrupted: ${item.corruptedReason ?? "unknown reason"}`
                          : isIncompatible
                            ? `Incompatible: ${item.incompatibleReason ?? "unknown reason"}`
                            : "OK"}
                      </div>
                      {isCorrupted || isIncompatible ? (
                        <button
                          type="button"
                          data-testid={`repair-extension-${item.id}`}
                          disabled={repairingId === item.id}
                          onClick={() => {
                            void (async () => {
                              if (!webExtensionManager) return;
                              setRepairingId(item.id);
                              try {
                                let shouldAttemptLoad = true;
                                if (isIncompatible) {
                                  const installedVersion = String(item.version ?? "");
                                  const reason = String(item.incompatibleReason ?? "");
                                  const isEngineMismatch = reason.toLowerCase().includes("engine mismatch");
                                  try {
                                    const updated = await webExtensionManager.update(item.id);
                                    const updatedVersion = String(updated?.version ?? "");

                                    const didUpdate = updatedVersion.length > 0 && updatedVersion !== installedVersion;
                                    if (!didUpdate) {
                                      // If the extension is quarantined due to a corrupted/invalid stored
                                      // manifest (not an engine mismatch), reinstalling the current version
                                      // can repair it even when no update is available.
                                      if (!isEngineMismatch) {
                                        await webExtensionManager.repair(item.id);
                                      } else {
                                        shouldAttemptLoad = false;
                                        try {
                                          showToast("No compatible update", "warning");
                                        } catch {
                                          // ignore missing toast root
                                        }
                                      }
                                    }
                                  } catch (error) {
                                    const msg = String((error as any)?.message ?? error);
                                    if (msg.toLowerCase().includes("engine mismatch")) {
                                      if (!isEngineMismatch) {
                                        // The latest version is incompatible with this engine, but the
                                        // stored incompatible quarantine may be caused by a corrupted
                                        // manifest. Reinstall the current version as a recovery path.
                                        await webExtensionManager.repair(item.id);
                                      } else {
                                        shouldAttemptLoad = false;
                                        try {
                                          showToast("No compatible update", "warning");
                                        } catch {
                                          // ignore missing toast root
                                        }
                                      }
                                    } else {
                                      throw error;
                                    }
                                  }
                                } else {
                                  await webExtensionManager.repair(item.id);
                                }
                                if (shouldAttemptLoad) {
                                  await webExtensionManager.loadInstalled(item.id).catch(() => {});
                                }
                                await refreshInstalled();
                                onSyncExtensions?.();
                                bump((v) => v + 1);
                              } catch (error: any) {
                                setInstallError(String(error?.message ?? error));
                              } finally {
                                setRepairingId((prev) => (prev === item.id ? null : prev));
                              }
                            })().catch(() => {});
                          }}
                          className="extensions-panel__button extensions-panel__reset-button extensions-panel__repair-button"
                        >
                          {repairingId === item.id
                            ? isIncompatible
                              ? "Updating…"
                              : "Repairing…"
                            : isIncompatible
                              ? "Update"
                              : "Repair"}
                        </button>
                      ) : null}
                    </div>
                  );
                })}
              </div>
            ) : installed ? (
              <div className="extensions-panel__empty">No marketplace extensions installed.</div>
            ) : installError ? null : (
              <div className="extensions-panel__empty">Loading installed extensions…</div>
            )}
          </div>
        ) : null}

      <div className="extensions-panel__toolbar">
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
          className="extensions-panel__button extensions-panel__toolbar-button"
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
          className="extensions-panel__button extensions-panel__toolbar-button"
        >
          Refresh
        </button>
      </div>

      {permissionsError ? (
        <div className="extensions-panel__error">Failed to load permissions: {permissionsError}</div>
      ) : null}

      {extensions.map((ext: any) => {
        const extCommands = commandsByExt.get(ext.id) ?? [];
        const extPanels = panelsByExt.get(ext.id) ?? [];

        const declaredPermissions = collectDeclaredPermissions(ext.manifest?.permissions);
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
          <div key={ext.id} data-testid={`extension-card-${ext.id}`} className="extensions-panel__card">
            <div className="extensions-panel__card-title">{ext.manifest?.displayName ?? ext.id}</div>
            <div className="extensions-panel__card-id">{ext.id}</div>

            <div className="extensions-panel__section">
              <div className="extensions-panel__section-title">Permissions</div>

              <div className="extensions-panel__permissions">
                {allPermissions.length === 0 ? (
                  <div className="extensions-panel__empty">No permissions declared.</div>
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
                      <div key={perm} data-testid={`permission-row-${ext.id}-${perm}`} className="extensions-panel__permission-row">
                        <div className="extensions-panel__permission-meta">
                          <div className="extensions-panel__permission-name">
                            <span>{perm}</span>
                            <span className="extensions-panel__permission-declared">{declared ? "declared" : "not declared"}</span>
                          </div>
                          <div className="extensions-panel__permission-subtitle" data-testid={`permission-${ext.id}-${perm}`}>
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
                          className="extensions-panel__button extensions-panel__revoke-button"
                        >
                          Revoke
                        </button>
                      </div>
                    );
                  })
                )}

                <div className="extensions-panel__reset-row">
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
                    className="extensions-panel__button extensions-panel__reset-button"
                  >
                    Revoke all permissions
                  </button>
                </div>
              </div>
            </div>

            {extCommands.length > 0 ? (
              <>
                <div className="extensions-panel__section-title">Commands</div>
                <div className="extensions-panel__commands">
                  {extCommands.map((cmd) => (
                    <div key={cmd.command} className="extensions-panel__command-row">
                      <button
                        type="button"
                        data-testid={`run-command-${cmd.command}`}
                        onClick={() => {
                          void executeCommandAndRefreshPermissions(String(ext.id), cmd.command).catch((err) => {
                            showToast(`Command failed: ${String((err as any)?.message ?? err)}`, "error");
                          });
                        }}
                        className="extensions-panel__button extensions-panel__command-button"
                      >
                        <div className="extensions-panel__command-header">
                          <div className="extensions-panel__command-title">
                            {cmd.category ? `${cmd.category}: ` : ""}
                            {cmd.title}
                          </div>
                          {(() => {
                            const shortcut = getPrimaryCommandKeybindingDisplay(cmd.command, keybindingIndex);
                            if (!shortcut) return null;
                            return (
                              <div aria-hidden="true" className="extensions-panel__command-shortcut">
                                {shortcut}
                              </div>
                            );
                          })()}
                        </div>
                        <div className="extensions-panel__command-id">{cmd.command}</div>
                      </button>
                      <button
                        type="button"
                        data-testid={`run-command-with-args-${cmd.command}`}
                        onClick={() => runCommandWithArgs(String(ext.id), cmd.command)}
                        className="extensions-panel__button extensions-panel__command-args-button"
                      >
                        Run…
                      </button>
                    </div>
                  ))}
                </div>
              </>
            ) : (
              <div className="extensions-panel__empty">No commands contributed.</div>
            )}

            {extPanels.length > 0 ? (
              <>
                <div className="extensions-panel__section-title extensions-panel__section-title--spaced">Panels</div>
                <div className="extensions-panel__panels">
                  {extPanels.map((panel) => (
                    <button
                      key={panel.id}
                      type="button"
                      data-testid={`open-panel-${panel.id}`}
                      onClick={() => onOpenPanel(panel.id)}
                      className="extensions-panel__button extensions-panel__panel-button"
                    >
                      <div className="extensions-panel__panel-title">{panel.title}</div>
                      <div className="extensions-panel__panel-id">{panel.id}</div>
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
