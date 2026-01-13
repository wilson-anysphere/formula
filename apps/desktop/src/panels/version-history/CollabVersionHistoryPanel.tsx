import React, { useEffect, useMemo, useState } from "react";

import type { CollabSession } from "@formula/collab-session";

import { buildVersionHistoryItems } from "./index.js";
import { t, tWithVars } from "../../i18n/index.js";
import * as nativeDialogs from "../../tauri/nativeDialogs.js";

function formatVersionTimestamp(timestampMs: number): string {
  try {
    return new Date(timestampMs).toLocaleString();
  } catch {
    return String(timestampMs);
  }
}

export function CollabVersionHistoryPanel({ session }: { session: CollabSession }) {
  // `@formula/collab-versioning` depends on the core versioning subsystem, which can pull in
  // Node-only modules (e.g. `node:events`). Avoid importing it at desktop shell startup so
  // split-view/grid e2e can boot without requiring those polyfills; load it lazily when the
  // panel is actually opened.
  const [collabVersioning, setCollabVersioning] = useState<any | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [versions, setVersions] = useState<any[]>([]);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const [checkpointName, setCheckpointName] = useState("");
  const [checkpointAnnotations, setCheckpointAnnotations] = useState("");
  const [checkpointLocked, setCheckpointLocked] = useState(false);

  useEffect(() => {
    let disposed = false;
    let instance: any | null = null;

    void (async () => {
      try {
        setLoadError(null);
        setCollabVersioning(null);
        const mod = await import("../../../../../packages/collab/versioning/src/index.js");
        if (disposed) return;
        const localPresence = session.presence?.localPresence ?? null;
        instance = mod.createCollabVersioning({
          session,
          user: localPresence ? { userId: localPresence.id, userName: localPresence.name } : undefined,
        });
        setCollabVersioning(instance);
      } catch (e) {
        if (disposed) return;
        setLoadError((e as Error).message);
      }
    })();

    return () => {
      disposed = true;
      instance?.destroy();
    };
  }, [session]);

  const refresh = async () => {
    try {
      setError(null);
      const manager = collabVersioning;
      if (!manager) return;
      const next = await manager.listVersions();
      setVersions(next);
      if (selectedId && !next.some((v: any) => v.id === selectedId)) setSelectedId(null);
    } catch (e) {
      setError((e as Error).message);
    }
  };

  useEffect(() => {
    if (!collabVersioning) return;
    void refresh();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [collabVersioning]);

  const items = useMemo(() => buildVersionHistoryItems(versions as any), [versions]);

  const kindLabel = (kind: string) =>
    kind === "checkpoint"
      ? t("versionHistory.checkpoint")
      : kind === "snapshot"
        ? t("versionHistory.autoSave")
        : kind === "restore"
          ? t("versionHistory.restore")
          : kind;

  const selectedVersion = useMemo(() => {
    if (!selectedId) return null;
    return versions.find((v) => v.id === selectedId) ?? null;
  }, [selectedId, versions]);

  const selectedItem = useMemo(() => {
    if (!selectedId) return null;
    return items.find((v) => v.id === selectedId) ?? null;
  }, [items, selectedId]);

  const selectedIsCheckpoint = selectedVersion?.kind === "checkpoint";
  const selectedLocked = selectedIsCheckpoint ? Boolean(selectedVersion?.checkpointLocked) : false;
  const selectedAnnotations = selectedIsCheckpoint ? (selectedVersion?.checkpointAnnotations ?? "") : "";

  const deleteDisabled = useMemo(() => {
    if (busy) return true;
    if (!selectedId) return true;
    if (selectedIsCheckpoint && selectedLocked) return true;
    return false;
  }, [busy, selectedId, selectedIsCheckpoint, selectedLocked]);

  if (loadError) {
    return (
      <div className="collab-panel__message collab-panel__message--error">
        {tWithVars("versionHistory.panel.unavailableWithMessage", { message: loadError })}
      </div>
    );
  }

  if (!collabVersioning) {
    return <div className="collab-panel__message">{t("versionHistory.panel.loading")}</div>;
  }

  return (
    <div className="collab-version-history">
      <h3 className="collab-version-history__title">{t("panels.versionHistory.title")}</h3>

      {error ? <div className="collab-version-history__error">{error}</div> : null}

      <div className="collab-version-history__create">
        <div className="collab-version-history__create-title">{t("versionHistory.actions.createCheckpoint")}</div>
        <div className="collab-version-history__form">
          <label className="collab-version-history__field">
            <div className="collab-version-history__label">{t("versionHistory.prompt.checkpointName")}</div>
            <input
              className="collab-version-history__input"
              value={checkpointName}
              onChange={(e) => setCheckpointName(e.target.value)}
              placeholder={t("versionHistory.prompt.checkpointName")}
            />
          </label>

          <label className="collab-version-history__field">
            <div className="collab-version-history__label">Annotations (optional)</div>
            <textarea
              className="collab-version-history__textarea"
              rows={3}
              value={checkpointAnnotations}
              onChange={(e) => setCheckpointAnnotations(e.target.value)}
              placeholder="Notes about this checkpoint…"
            />
          </label>

          <label className="collab-version-history__checkbox">
            <input
              type="checkbox"
              checked={checkpointLocked}
              onChange={(e) => setCheckpointLocked(e.target.checked)}
            />
            Locked
          </label>

          <div className="collab-version-history__create-actions">
            <button
              disabled={busy || !checkpointName.trim()}
              onClick={async () => {
                const name = checkpointName.trim();
                if (!name) {
                  setError("Checkpoint name is required.");
                  return;
                }
                try {
                  setBusy(true);
                  setError(null);
                  const created = await collabVersioning.createCheckpoint({
                    name,
                    annotations: checkpointAnnotations.trim() ? checkpointAnnotations.trim() : undefined,
                    locked: checkpointLocked,
                  });
                  setCheckpointName("");
                  setCheckpointAnnotations("");
                  setCheckpointLocked(false);
                  await refresh();
                  setSelectedId(created?.id ?? null);
                } catch (e) {
                  setError((e as Error).message);
                } finally {
                  setBusy(false);
                }
              }}
            >
              {t("versionHistory.actions.createCheckpoint")}
            </button>
          </div>
        </div>
      </div>

      <div className="collab-version-history__actions">
        <button
          disabled={busy || !selectedId}
          onClick={async () => {
            const id = selectedId;
            if (!id) return;
            const ok = await nativeDialogs.confirm(t("versionHistory.confirm.restoreOverwrite"));
            if (!ok) return;
            try {
              setBusy(true);
              setError(null);
              await collabVersioning.restoreVersion(id);
              await refresh();
            } catch (e) {
              setError((e as Error).message);
            } finally {
              setBusy(false);
            }
          }}
        >
          {t("versionHistory.actions.restoreSelected")}
        </button>

        {selectedIsCheckpoint ? (
          <button
            disabled={busy || !selectedId}
            onClick={async () => {
              const id = selectedId;
              if (!id) return;
              try {
                setBusy(true);
                setError(null);
                await collabVersioning.setCheckpointLocked(id, !selectedLocked);
                await refresh();
              } catch (e) {
                setError((e as Error).message);
              } finally {
                setBusy(false);
              }
            }}
          >
            {selectedLocked ? t("versionHistory.actions.unlock") : t("versionHistory.actions.lock")}
          </button>
        ) : null}

        <button
          disabled={deleteDisabled}
          onClick={async () => {
            const id = selectedId;
            if (!id) return;
            const ok = await nativeDialogs.confirm("Delete this version? This cannot be undone.");
            if (!ok) return;
            try {
              setBusy(true);
              setError(null);
              await collabVersioning.deleteVersion(id);
              await refresh();
            } catch (e) {
              setError((e as Error).message);
            } finally {
              setBusy(false);
            }
          }}
        >
          {t("versionHistory.actions.delete")} selected
        </button>

        <button disabled={busy} onClick={() => void refresh()}>
          {t("versionHistory.actions.refresh")}
        </button>
      </div>

      {selectedIsCheckpoint && selectedLocked ? (
        <div className="collab-version-history__hint">Locked checkpoints must be unlocked before deleting.</div>
      ) : null}

      {items.length === 0 ? (
        <div className="collab-version-history__empty">{t("versionHistory.panel.empty")}</div>
      ) : (
        <ul className="collab-version-history__list">
          {items.map((item) => {
            const selected = item.id === selectedId;
            const formattedKind = kindLabel(item.kind);
            return (
              <li
                key={item.id}
                className={
                  selected
                    ? "collab-version-history__item collab-version-history__item--selected"
                    : "collab-version-history__item"
                }
                onClick={() => setSelectedId(item.id)}
              >
                <input type="radio" checked={selected} onChange={() => setSelectedId(item.id)} />
                <div className="collab-version-history__item-content">
                  <div className="collab-version-history__item-title">
                    {item.title}
                    {item.locked ? <span className="collab-version-history__badge">{t("versionHistory.meta.locked")}</span> : null}
                  </div>
                  <div className="collab-version-history__item-meta">
                    {formatVersionTimestamp(item.timestampMs)} • {formattedKind}
                  </div>
                </div>
              </li>
            );
          })}
        </ul>
      )}

      {selectedItem ? (
        <div className="collab-version-history__details">
          <div className="collab-version-history__details-title">{t("versionHistory.diff.selected")}</div>
          <div className="collab-version-history__details-meta">
            {selectedItem.title} • {formatVersionTimestamp(selectedItem.timestampMs)} • {kindLabel(selectedItem.kind)}
            {selectedIsCheckpoint && selectedLocked ? ` • ${t("versionHistory.meta.locked")}` : ""}
          </div>

          {selectedIsCheckpoint && selectedAnnotations.trim() ? (
            <div className="collab-version-history__details-annotations">{selectedAnnotations}</div>
          ) : null}
        </div>
      ) : null}
    </div>
  );
}
