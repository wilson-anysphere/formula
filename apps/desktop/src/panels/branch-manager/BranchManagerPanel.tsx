import React, { useEffect, useMemo, useState } from "react";

import { t, tWithVars } from "../../i18n/index.js";
import * as nativeDialogs from "../../tauri/nativeDialogs.js";

/**
 * The desktop app wires this panel to the real document controller + branch
 * service. The implementation here is intentionally small and focused on the
 * workflow:
 * - create / rename / delete / switch branches
 * - launch merge workflow
 */

export type Branch = {
  id: string;
  name: string;
  description: string | null;
  createdBy: string;
  createdAt: number;
  headCommitId: string;
};

export type Actor = { userId: string; role: "owner" | "admin" | "editor" | "commenter" | "viewer" };

export type BranchService = {
  listBranches(): Promise<Branch[]>;
  createBranch(actor: Actor, input: { name: string; description?: string }): Promise<Branch>;
  renameBranch(actor: Actor, input: { oldName: string; newName: string }): Promise<void>;
  deleteBranch(actor: Actor, input: { name: string }): Promise<void>;
  checkoutBranch(actor: Actor, input: { name: string }): Promise<unknown>;
  previewMerge(
    actor: Actor,
    input: { sourceBranch: string }
  ): Promise<{ conflicts: unknown[]; merged: unknown }>;
};

export function BranchManagerPanel({
  actor,
  branchService,
  onStartMerge
}: {
  actor: Actor;
  branchService: BranchService;
  onStartMerge: (sourceBranch: string) => void;
}) {
  const [branches, setBranches] = useState<Branch[]>([]);
  const [newBranchName, setNewBranchName] = useState("");
  const [error, setError] = useState<string | null>(null);

  const reload = async () => {
    try {
      setError(null);
      setBranches(await branchService.listBranches());
    } catch (e) {
      setError((e as Error).message);
    }
  };

  useEffect(() => {
    void reload();
  }, []);

  const canManage = useMemo(() => actor.role === "owner" || actor.role === "admin", [actor.role]);

  return (
    <div style={{ padding: 12, fontFamily: "system-ui, sans-serif" }}>
      <h3>{t("branchManager.title")}</h3>
      {!canManage && (
        <div style={{ color: "var(--text-secondary)", marginBottom: 8 }}>
          {t("branchManager.permissionWarning")}
        </div>
      )}
      {error && (
        <div style={{ color: "var(--error)", marginBottom: 8 }}>
          {error}
        </div>
      )}

      <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
        <input
          value={newBranchName}
          onChange={(e) => setNewBranchName(e.target.value)}
          placeholder={t("branchManager.newBranch.placeholder")}
          disabled={!canManage}
        />
        <button
          disabled={!canManage || !newBranchName.trim()}
          onClick={async () => {
            try {
              await branchService.createBranch(actor, { name: newBranchName.trim() });
              setNewBranchName("");
              await reload();
            } catch (e) {
              setError((e as Error).message);
            }
          }}
        >
          {t("branchManager.newBranch.create")}
        </button>
      </div>

      <ul style={{ listStyle: "none", padding: 0, margin: 0 }}>
        {branches.map((b) => (
          <li
            key={b.id}
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              padding: "6px 0",
              borderBottom: "1px solid var(--border)"
            }}
          >
            <div>
              <div style={{ fontWeight: 600 }}>{b.name}</div>
              {b.description ? <div style={{ color: "var(--text-secondary)" }}>{b.description}</div> : null}
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              <button
                disabled={!canManage}
                onClick={async () => {
                  const newName = window.prompt(t("branchManager.prompt.rename"), b.name);
                  if (!newName || newName.trim() === b.name) return;
                  try {
                    await branchService.renameBranch(actor, { oldName: b.name, newName: newName.trim() });
                    await reload();
                  } catch (e) {
                    setError((e as Error).message);
                  }
                }}
              >
                {t("branchManager.actions.rename")}
              </button>
              <button
                disabled={!canManage}
                onClick={async () => {
                  try {
                    await branchService.checkoutBranch(actor, { name: b.name });
                    await reload();
                  } catch (e) {
                    setError((e as Error).message);
                  }
                }}
              >
                {t("branchManager.actions.switch")}
              </button>
              <button
                disabled={!canManage || b.name === "main"}
                onClick={async () => {
                  const ok = await nativeDialogs.confirm(tWithVars("branchManager.confirm.delete", { name: b.name }));
                  if (!ok) return;
                  try {
                    await branchService.deleteBranch(actor, { name: b.name });
                    await reload();
                  } catch (e) {
                    setError((e as Error).message);
                  }
                }}
              >
                {t("branchManager.actions.delete")}
              </button>
              <button
                disabled={!canManage}
                onClick={() => onStartMerge(b.name)}
              >
                {t("branchManager.actions.merge")}
              </button>
            </div>
          </li>
        ))}
      </ul>
    </div>
  );
}
