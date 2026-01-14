import React, { useEffect, useMemo, useState } from "react";

import { t, tWithVars } from "../../i18n/index.js";
import * as nativeDialogs from "../../tauri/nativeDialogs.js";
import { showInputBox } from "../../extensions/ui.js";

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
  getCurrentBranchName(): Promise<string>;
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
  onStartMerge,
  mutationsDisabled = false,
}: {
  actor: Actor;
  branchService: BranchService;
  onStartMerge: (sourceBranch: string) => void;
  mutationsDisabled?: boolean;
}) {
  const [branches, setBranches] = useState<Branch[]>([]);
  const [currentBranchName, setCurrentBranchName] = useState<string | null>(null);
  const [newBranchName, setNewBranchName] = useState("");
  const [error, setError] = useState<string | null>(null);

  const reload = async () => {
    try {
      setError(null);
      const [nextBranches, current] = await Promise.all([
        branchService.listBranches(),
        branchService.getCurrentBranchName(),
      ]);
      setBranches(nextBranches);
      setCurrentBranchName(current);
    } catch (e) {
      setError((e as Error).message);
    }
  };

  useEffect(() => {
    void reload().catch(() => {
      // Best-effort: avoid unhandled rejections from fire-and-forget effect calls.
    });
  }, [branchService]);

  const canManage = useMemo(() => actor.role === "owner" || actor.role === "admin", [actor.role]);

  return (
    <div className="branch-manager">
      <h3 className="branch-manager__title">{t("branchManager.title")}</h3>
      {!canManage && (
        <div className="branch-manager__permission-warning">
          {t("branchManager.permissionWarning")}
        </div>
      )}
      {error && (
        <div className="branch-manager__error">
          {error}
        </div>
      )}

      <div className="branch-manager__new-branch">
        <input
          value={newBranchName}
          onChange={(e) => setNewBranchName(e.target.value)}
          placeholder={t("branchManager.newBranch.placeholder")}
          disabled={!canManage || mutationsDisabled}
        />
        <button
          disabled={!canManage || mutationsDisabled || !newBranchName.trim()}
          onClick={() => {
            void (async () => {
              try {
                await branchService.createBranch(actor, { name: newBranchName.trim() });
                setNewBranchName("");
                await reload();
              } catch (e) {
                setError((e as Error).message);
              }
            })().catch((e) => {
              setError((e as Error)?.message ?? String(e));
            });
          }}
        >
          {t("branchManager.newBranch.create")}
        </button>
      </div>

      <ul className="branch-manager__list">
        {branches.map((b) => {
          const isCurrent = b.name === currentBranchName;
          return (
          <li
            key={b.id}
            className={isCurrent ? "branch-manager__item branch-manager__item--current" : "branch-manager__item"}
          >
            <div className="branch-manager__item-content">
              <div className="branch-manager__item-title">
                {b.name}
                {isCurrent ? <span className="branch-manager__current-badge">{t("branchManager.current")}</span> : null}
              </div>
              {b.description ? <div className="branch-manager__item-description">{b.description}</div> : null}
            </div>
            <div className="branch-manager__item-actions">
              <button
                disabled={!canManage || mutationsDisabled}
                onClick={() => {
                  void (async () => {
                    const newName = await showInputBox({ prompt: t("branchManager.prompt.rename"), value: b.name });
                    const trimmed = newName?.trim();
                    if (!trimmed || trimmed === b.name) return;
                    try {
                      await branchService.renameBranch(actor, { oldName: b.name, newName: trimmed });
                      await reload();
                    } catch (e) {
                      setError((e as Error).message);
                    }
                  })().catch((e) => {
                    setError((e as Error)?.message ?? String(e));
                  });
                }}
              >
                {t("branchManager.actions.rename")}
              </button>
              <button
                disabled={!canManage || mutationsDisabled || isCurrent}
                onClick={() => {
                  void (async () => {
                    try {
                      await branchService.checkoutBranch(actor, { name: b.name });
                      await reload();
                    } catch (e) {
                      setError((e as Error).message);
                    }
                  })().catch((e) => {
                    setError((e as Error)?.message ?? String(e));
                  });
                }}
              >
                {t("branchManager.actions.switch")}
              </button>
              <button
                disabled={!canManage || mutationsDisabled || b.name === "main" || isCurrent}
                onClick={() => {
                  void (async () => {
                    const ok = await nativeDialogs.confirm(tWithVars("branchManager.confirm.delete", { name: b.name }));
                    if (!ok) return;
                    try {
                      await branchService.deleteBranch(actor, { name: b.name });
                      await reload();
                    } catch (e) {
                      setError((e as Error).message);
                    }
                  })().catch((e) => {
                    setError((e as Error)?.message ?? String(e));
                  });
                }}
              >
                {t("branchManager.actions.delete")}
              </button>
              <button
                disabled={!canManage || mutationsDisabled || isCurrent}
                onClick={() => onStartMerge(b.name)}
              >
                {t("branchManager.actions.merge")}
              </button>
            </div>
          </li>
        );
        })}
      </ul>
    </div>
  );
}
