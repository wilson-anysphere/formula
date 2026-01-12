import React from "react";

import type { DesktopExtensionHostManager } from "./extensionHostManager.js";

type NetworkPolicy = {
  mode: "full" | "deny" | "allowlist" | string;
  hosts?: string[];
};

type GrantedPermissions = Record<string, true | NetworkPolicy>;

type PermissionState =
  | { status: "loading" }
  | { status: "error"; error: string }
  | { status: "ready"; data: GrantedPermissions };

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

export function ExtensionsPanel({
  manager,
  onExecuteCommand,
  onOpenPanel,
}: {
  manager: DesktopExtensionHostManager;
  onExecuteCommand: (commandId: string) => void;
  onOpenPanel: (panelId: string) => void;
}) {
  const [, bump] = React.useState(0);
  React.useEffect(() => manager.subscribe(() => bump((v) => v + 1)), [manager]);

  const extensions = manager.host.listExtensions();
  const commands = manager.getContributedCommands() as ContributedCommand[];
  const panels = manager.getContributedPanels() as ContributedPanel[];

  const commandsByExt = groupByExtension(commands);
  const panelsByExt = groupByExtension(panels);

  const [permissionsByExt, setPermissionsByExt] = React.useState<Record<string, PermissionState>>({});
  const [permissionsVersion, setPermissionsVersion] = React.useState(0);

  const fetchPermissions = React.useCallback(
    async (extensionId: string): Promise<PermissionState> => {
      try {
        const data = (await manager.host.getGrantedPermissions(extensionId)) as GrantedPermissions;
        return { status: "ready", data };
      } catch (err) {
        return { status: "error", error: String((err as any)?.message ?? err) };
      }
    },
    [manager],
  );

  const loadPermissionsForExtension = React.useCallback(
    async (extensionId: string) => {
      const id = String(extensionId);
      setPermissionsByExt((prev) => ({ ...prev, [id]: { status: "loading" } }));
      const next = await fetchPermissions(id);
      setPermissionsByExt((prev) => ({ ...prev, [id]: next }));
    },
    [fetchPermissions],
  );

  const refreshAllPermissions = React.useCallback(
    async (extensionIds: string[]) => {
      const ids = extensionIds.map(String);
      setPermissionsByExt((prev) => {
        const next = { ...prev };
        for (const id of ids) next[id] = { status: "loading" };
        return next;
      });

      const entries = await Promise.all(ids.map(async (id) => [id, await fetchPermissions(id)] as const));
      setPermissionsByExt((prev) => ({ ...prev, ...Object.fromEntries(entries) }));
    },
    [fetchPermissions],
  );

  const extensionsKey = extensions.map((ext: any) => String(ext.id)).join("|");

  React.useEffect(() => {
    if (!manager.ready || manager.error) return;
    const ids = extensions.map((ext: any) => String(ext.id));
    void refreshAllPermissions(ids);
    // Intentionally depend on a stable key instead of the `extensions` array identity.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [manager.ready, manager.error, permissionsVersion, extensionsKey, refreshAllPermissions]);

  const resetAllPermissions = React.useCallback(() => {
    void (async () => {
      await manager.host.resetAllPermissions();
      setPermissionsVersion((v) => v + 1);
    })().catch(() => {
      // ignore
    });
  }, [manager]);

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
      <div style={{ display: "flex", justifyContent: "flex-end" }}>
        <button
          type="button"
          data-testid="reset-all-extension-permissions"
          onClick={resetAllPermissions}
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
      </div>
      {extensions.map((ext: any) => {
        const extCommands = commandsByExt.get(ext.id) ?? [];
        const extPanels = panelsByExt.get(ext.id) ?? [];
        const permState = permissionsByExt[String(ext.id)] ?? ({ status: "loading" } as const);
        const granted = permState.status === "ready" ? permState.data : null;

        const booleanPerms =
          granted != null
            ? Object.entries(granted)
                .filter(([key, value]) => key !== "network" && value === true)
                .map(([key]) => key)
                .sort((a, b) => a.localeCompare(b))
            : [];
        const networkPolicy =
          granted && typeof (granted as any).network === "object" ? ((granted as any).network as NetworkPolicy) : null;
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

            {extCommands.length > 0 ? (
              <>
                <div style={{ fontWeight: 600, marginBottom: "6px" }}>Commands</div>
                <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                  {extCommands.map((cmd) => (
                    <button
                      key={cmd.command}
                      type="button"
                      data-testid={`run-command-${cmd.command}`}
                      onClick={() => onExecuteCommand(cmd.command)}
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
                      <div style={{ fontWeight: 600 }}>
                        {cmd.category ? `${cmd.category}: ` : ""}
                        {cmd.title}
                      </div>
                      <div style={{ fontSize: "12px", color: "var(--text-secondary)" }}>{cmd.command}</div>
                    </button>
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

            <div style={{ fontWeight: 600, marginTop: "12px", marginBottom: "6px" }}>Permissions</div>
            {permState.status === "loading" ? (
              <div style={{ color: "var(--text-secondary)" }}>Loading permissions…</div>
            ) : permState.status === "error" ? (
              <div style={{ color: "var(--text-secondary)" }}>Failed to load permissions: {permState.error}</div>
            ) : booleanPerms.length === 0 && !networkPolicy ? (
              <div data-testid={`permissions-empty-${ext.id}`} style={{ color: "var(--text-secondary)" }}>
                No permissions granted.
              </div>
            ) : (
              <div
                data-testid={`permissions-list-${ext.id}`}
                style={{ display: "flex", flexDirection: "column", gap: "6px" }}
              >
                {booleanPerms.map((perm) => (
                  <div
                    key={perm}
                    data-testid={`permission-${ext.id}-${perm}`}
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      gap: "8px",
                      padding: "6px 8px",
                      borderRadius: "8px",
                      border: "1px solid var(--border)",
                      background: "var(--bg-secondary)",
                    }}
                  >
                    <div style={{ minWidth: 0 }}>
                      <div style={{ fontFamily: "monospace", fontSize: "12px", wordBreak: "break-all" }}>{perm}</div>
                      <div style={{ fontSize: "12px", color: "var(--text-secondary)" }}>granted</div>
                    </div>
                    <button
                      type="button"
                      data-testid={`revoke-permission-${ext.id}-${perm}`}
                      onClick={() => {
                        void (async () => {
                          await manager.host.revokePermissions(ext.id, [perm]);
                          await loadPermissionsForExtension(ext.id);
                        })().catch(() => {
                          // ignore
                        });
                      }}
                      style={{
                        padding: "6px 8px",
                        borderRadius: "10px",
                        border: "1px solid var(--border)",
                        background: "var(--bg-primary)",
                        color: "var(--text-primary)",
                        cursor: "pointer",
                        flex: "0 0 auto",
                      }}
                    >
                      Revoke
                    </button>
                  </div>
                ))}

                {networkPolicy ? (
                  <div
                    data-testid={`permission-${ext.id}-network`}
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      gap: "8px",
                      padding: "6px 8px",
                      borderRadius: "8px",
                      border: "1px solid var(--border)",
                      background: "var(--bg-secondary)",
                    }}
                  >
                    <div style={{ minWidth: 0 }}>
                      <div style={{ fontFamily: "monospace", fontSize: "12px" }}>network</div>
                      <div style={{ fontSize: "12px", color: "var(--text-secondary)" }}>
                        mode: {String(networkPolicy.mode)}
                        {Array.isArray(networkPolicy.hosts) && networkPolicy.hosts.length > 0
                          ? `, hosts: ${networkPolicy.hosts.join(", ")}`
                          : ""}
                      </div>
                    </div>
                    <button
                      type="button"
                      data-testid={`revoke-permission-${ext.id}-network`}
                      onClick={() => {
                        void (async () => {
                          await manager.host.revokePermissions(ext.id, ["network"]);
                          await loadPermissionsForExtension(ext.id);
                        })().catch(() => {
                          // ignore
                        });
                      }}
                      style={{
                        padding: "6px 8px",
                        borderRadius: "10px",
                        border: "1px solid var(--border)",
                        background: "var(--bg-primary)",
                        color: "var(--text-primary)",
                        cursor: "pointer",
                        flex: "0 0 auto",
                      }}
                    >
                      Revoke
                    </button>
                  </div>
                ) : null}
              </div>
            )}

            <div style={{ display: "flex", gap: "8px", marginTop: "10px" }}>
              <button
                type="button"
                data-testid={`revoke-all-permissions-${ext.id}`}
                onClick={() => {
                  void (async () => {
                    await manager.host.revokePermissions(ext.id);
                    await loadPermissionsForExtension(ext.id);
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
                Revoke all permissions
              </button>
            </div>
          </div>
        );
      })}
    </div>
  );
}
