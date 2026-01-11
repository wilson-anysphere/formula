import React from "react";

import type { DesktopExtensionHostManager } from "./extensionHostManager.js";

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
      {extensions.map((ext: any) => {
        const extCommands = commandsByExt.get(ext.id) ?? [];
        const extPanels = panelsByExt.get(ext.id) ?? [];
        return (
          <div
            key={ext.id}
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
          </div>
        );
      })}
    </div>
  );
}
