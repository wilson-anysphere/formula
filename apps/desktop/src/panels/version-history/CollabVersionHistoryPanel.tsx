import React, { useEffect, useMemo, useState } from "react";

import type { CollabSession } from "@formula/collab-session";

import { buildVersionHistoryItems } from "./VersionHistoryPanel.js";
import { VersionHistoryCompareSection } from "./VersionHistoryCompare.js";
import { t, tWithVars } from "../../i18n/index.js";
import * as nativeDialogs from "../../tauri/nativeDialogs.js";
import { clearReservedRootGuardError, useReservedRootGuardError } from "../collabReservedRootGuard.react.js";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver";
import { useCollabSessionSyncState } from "../useCollabSessionSyncState.js";
import { createCollabVersioningForPanel, type CreateVersionStore } from "./createCollabVersioningForPanel.js";
import type { VersionStore } from "../../../../../packages/collab/versioning/src/index.ts";

function formatVersionTimestamp(timestampMs: number): string {
  try {
    return new Date(timestampMs).toLocaleString();
  } catch {
    return String(timestampMs);
  }
}

export function CollabVersionHistoryPanel({
  session,
  sheetNameResolver = null,
  createVersionStore,
  versionStore,
}: {
  session: CollabSession;
  sheetNameResolver?: SheetNameResolver | null;
  /**
   * Optional VersionStore provider. Use this to inject an out-of-doc store so the
   * panel does not write to reserved Yjs roots (`versions*`).
   */
  createVersionStore?: CreateVersionStore;
  /**
   * Optional pre-constructed VersionStore instance. Prefer {@link createVersionStore}
   * to keep store construction lazy.
   */
  versionStore?: VersionStore;
}) {
  const syncState = useCollabSessionSyncState(session);
  const hasInjectedStore = Boolean(createVersionStore || versionStore);
  const reservedRootGuardError = useReservedRootGuardError((session as any)?.provider ?? null);
  // Reserved root guard disconnects are sticky (we remember them per provider) so
  // panels opened later can show the banner. When using an out-of-doc store, we
  // only need to disable actions if the provider is currently disconnected.
  const mutationsDisabled = Boolean(reservedRootGuardError) && (!hasInjectedStore || !syncState.connected);
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

  const banner = mutationsDisabled && reservedRootGuardError ? (
    <div className="collab-panel__message collab-panel__message--error" data-testid="reserved-root-guard-error">
      <div>{reservedRootGuardError}</div>
      <button
        type="button"
        data-testid="reserved-root-guard-retry"
        disabled={busy}
        onClick={() => {
          try {
            collabVersioning?.destroy?.();
          } catch {
            // ignore
          }
          clearReservedRootGuardError((session as any)?.provider ?? null);
          try {
            (session as any)?.provider?.connect?.();
          } catch {
            // ignore
          }
          setError(null);
          setLoadError(null);
          setCollabVersioning(null);
        }}
      >
        {t("collab.retry")}
      </button>
    </div>
  ) : null;

  const [checkpointName, setCheckpointName] = useState("");
  const [checkpointAnnotations, setCheckpointAnnotations] = useState("");
  const [checkpointLocked, setCheckpointLocked] = useState(false);

  useEffect(() => {
    if (!mutationsDisabled) return;
    try {
      // Stop auto snapshot timers so we don't keep mutating the in-doc version store
      // after the sync server has rejected reserved root updates.
      collabVersioning?.destroy?.();
    } catch {
      // ignore
    }
  }, [mutationsDisabled, collabVersioning]);

  useEffect(() => {
    if (mutationsDisabled) return;
    let disposed = false;
    let instance: any | null = null;

    void (async () => {
      try {
        setLoadError(null);
        setCollabVersioning(null);
        if (disposed) return;
        instance = await createCollabVersioningForPanel({ session, store: versionStore, createVersionStore });
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
  }, [session, mutationsDisabled, createVersionStore, versionStore]);

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
    if (mutationsDisabled) return;
    void refresh();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [collabVersioning, mutationsDisabled]);

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

  const versioningReady = Boolean(collabVersioning);

  const deleteDisabled = useMemo(() => {
    if (busy) return true;
    if (mutationsDisabled) return true;
    if (!versioningReady) return true;
    if (!selectedId) return true;
    if (selectedIsCheckpoint && selectedLocked) return true;
    return false;
  }, [busy, mutationsDisabled, versioningReady, selectedId, selectedIsCheckpoint, selectedLocked]);

  return (
    <div className="collab-version-history">
      <h3 className="collab-version-history__title">{t("panels.versionHistory.title")}</h3>

      {banner}

      {loadError ? (
        <div className="collab-panel__message collab-panel__message--error">
          {tWithVars("versionHistory.panel.unavailableWithMessage", { message: loadError })}
        </div>
      ) : null}

      {!versioningReady && !mutationsDisabled && !loadError ? (
        <div role="status" className="collab-panel__message">
          {t("versionHistory.panel.loading")}
        </div>
      ) : null}
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
              disabled={busy || mutationsDisabled}
            />
          </label>

          <label className="collab-version-history__field">
            <div className="collab-version-history__label">{t("versionHistory.prompt.annotationsOptional")}</div>
            <textarea
              className="collab-version-history__textarea"
              rows={3}
              value={checkpointAnnotations}
              onChange={(e) => setCheckpointAnnotations(e.target.value)}
              placeholder={t("versionHistory.prompt.annotationsPlaceholder")}
              disabled={busy || mutationsDisabled}
            />
          </label>

          <label className="collab-version-history__checkbox">
            <input
              type="checkbox"
              checked={checkpointLocked}
              onChange={(e) => setCheckpointLocked(e.target.checked)}
              disabled={busy || mutationsDisabled}
            />
            {t("versionHistory.meta.locked")}
          </label>

          <div className="collab-version-history__create-actions">
            <button
              disabled={busy || mutationsDisabled || !versioningReady || !checkpointName.trim()}
              onClick={async () => {
                if (mutationsDisabled) return;
                if (!collabVersioning) return;
                const name = checkpointName.trim();
                if (!name) {
                  setError(t("versionHistory.errors.checkpointNameRequired"));
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
          disabled={busy || mutationsDisabled || !versioningReady || !selectedId}
          onClick={async () => {
            if (mutationsDisabled) return;
            if (!collabVersioning) return;
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
            disabled={busy || mutationsDisabled || !versioningReady || !selectedId}
            onClick={async () => {
              if (mutationsDisabled) return;
              if (!collabVersioning) return;
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
            if (mutationsDisabled) return;
            if (!collabVersioning) return;
            const id = selectedId;
            if (!id) return;
            const ok = await nativeDialogs.confirm(t("versionHistory.confirm.deleteIrreversible"));
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
          {t("versionHistory.actions.deleteSelected")}
        </button>

        <button disabled={busy || mutationsDisabled || !versioningReady} onClick={() => void refresh()}>
          {t("versionHistory.actions.refresh")}
        </button>
      </div>

      {selectedIsCheckpoint && selectedLocked ? (
        <div className="collab-version-history__hint">{t("versionHistory.hint.unlockToDelete")}</div>
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

      <VersionHistoryCompareSection
        versionId={selectedId}
        // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
        versionManager={(collabVersioning as any)?.manager ?? null}
        sheetNameResolver={sheetNameResolver}
      />
    </div>
  );
}
