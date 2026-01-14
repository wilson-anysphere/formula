/**
 * Minimal, dependency-free "in-app" marketplace panel.
 *
 * The real desktop app will render this inside the application's panel system
 * (Tauri/WebView). This module focuses on wiring: search → install/update/uninstall.
 */

import {
  getDefaultSeedStoreStorage,
  removeSeedPanelsForExtension,
  setSeedPanelsForExtension,
} from "../../extensions/contributedPanelsSeedStore.js";
import { showQuickPick, showToast } from "../../extensions/ui.js";
import { getTauriDialogConfirmOrNull } from "../../tauri/api.js";
import * as nativeDialogs from "../../tauri/nativeDialogs.js";

function tryShowToast(message, type = "info") {
  try {
    showToast(String(message ?? ""), type);
  } catch {
    // ignore missing toast root
  }
}

function tryNotifyExtensionsChanged() {
  try {
    window.dispatchEvent(new Event("formula:extensions-changed"));
  } catch {
    // ignore
  }
}

function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  for (const [key, value] of Object.entries(attrs)) {
    if (value === undefined || value === null) continue;
    if (key === "className") node.className = value;
    else if (key === "dataset" && value && typeof value === "object") {
      for (const [k, v] of Object.entries(value)) node.dataset[k] = String(v);
    }
    else if (key.startsWith("on") && typeof value === "function") node.addEventListener(key.slice(2).toLowerCase(), value);
    else node.setAttribute(key, String(value));
  }
  for (const child of children) node.append(child);
  return node;
}

function updateContributedPanelSeedsFromHost(extensionHostManager, extensionId) {
  const storage = getDefaultSeedStoreStorage();
  if (!storage) return;

  if (!extensionHostManager || typeof extensionHostManager.listContributions !== "function") return;

  try {
    const contributed = extensionHostManager.listContributions()?.panels ?? [];
    const panels = contributed
      .filter((p) => p && typeof p === "object" && p.extensionId === extensionId)
      .map((p) => ({ id: p.id, title: p.title, icon: p.icon ?? null }));

    const ok = setSeedPanelsForExtension(storage, extensionId, panels, {
      onError: (message) => {
        // eslint-disable-next-line no-console
        console.error(message);
        tryShowToast(message, "error");
      },
    });

    if (!ok) {
      // setSeedPanelsForExtension already surfaced the error.
      return;
    }
  } catch (error) {
    // eslint-disable-next-line no-console
    console.error("Failed to update contributed panel seed store:", error);
  }
}

function badge(text, { tone = "neutral", title = null } = {}) {
  const bgByTone = {
    neutral: "var(--bg-tertiary)",
    good: "var(--success-bg)",
    warn: "var(--warning-bg)",
    bad: "var(--error-bg)",
  };
  const fgByTone = {
    neutral: "var(--text-secondary)",
    good: "var(--success)",
    warn: "var(--warning)",
    bad: "var(--error)",
  };
  const borderByTone = {
    neutral: "var(--border)",
    good: "var(--success)",
    warn: "var(--warning)",
    bad: "var(--error)",
  };
  return el(
    "span",
    {
      className: "marketplace-badge",
      style: [
        "display:inline-flex",
        "align-items:center",
        "padding:2px 8px",
        "border-radius:999px",
        "font-size:11px",
        "font-weight:600",
        `background:${bgByTone[tone] || bgByTone.neutral}`,
        `border:1px solid ${borderByTone[tone] || borderByTone.neutral}`,
        `color:${fgByTone[tone] || fgByTone.neutral}`,
      ].join(";"),
      title: title || undefined,
    },
    [document.createTextNode(text)],
  );
}

function normalizeScanStatus(value) {
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;
  return trimmed.toLowerCase();
}

function getLatestVersionScanStatus(details) {
  if (!details || typeof details !== "object") return null;
  const latest = details.latestVersion;
  if (typeof latest !== "string" || !latest) return null;
  const versions = Array.isArray(details.versions) ? details.versions : [];
  const record = versions.find((v) => v && String(v.version) === String(latest)) || null;
  return normalizeScanStatus(record && record.scanStatus);
}

function describePolicy(details, scanPolicy) {
  if (!details) return { blocked: false, reason: null, warning: null };

  if (details.blocked) return { blocked: true, reason: "Blocked by marketplace", warning: null };
  if (details.malicious) return { blocked: true, reason: "Flagged as malicious", warning: null };
  if (details.publisherRevoked) return { blocked: true, reason: "Publisher revoked", warning: null };

  const warnings = [];
  const policy = scanPolicy === "allow" || scanPolicy === "ignore" || scanPolicy === "enforce" ? scanPolicy : "enforce";

  const scanStatus = getLatestVersionScanStatus(details);
  if (policy !== "ignore") {
    if (!scanStatus) {
      if (policy === "enforce") return { blocked: true, reason: "Missing security scan status", warning: null };
      warnings.push("Missing security scan status");
    } else if (scanStatus !== "passed") {
      if (policy === "enforce") return { blocked: true, reason: `Security scan not passed (${scanStatus})`, warning: null };
      warnings.push(`Security scan not passed (${scanStatus})`);
    }
  }

  if (details.deprecated) warnings.push("Deprecated");

  return { blocked: false, reason: null, warning: warnings.length > 0 ? warnings.join(" · ") : null };
}

async function renderSearchResults({
  container,
  marketplaceClient,
  extensionManager,
  extensionHostManager,
  query,
  transientStatusById,
}) {
  container.textContent = "Searching…";
  const results = await marketplaceClient.search({ q: query, limit: 25, offset: 0 });

  const list = el("div", { className: "marketplace-results", dataset: { testid: "marketplace-results" } });

  const detailsById = new Map();
  await Promise.all(
    results.results.map(async (item) => {
      try {
        detailsById.set(item.id, await marketplaceClient.getExtension(item.id));
      } catch {
        detailsById.set(item.id, null);
      }
    }),
  );

  for (const item of results.results) {
    let installed = null;
    try {
      installed = await extensionManager.getInstalled(item.id);
    } catch {
      installed = null;
    }
    const details = detailsById.get(item.id) || null;
    const transientStatus = transientStatusById?.get(item.id) ?? null;

    const latestVersion = details?.latestVersion || item.latestVersion || null;
    const latestScanStatusRaw =
      latestVersion && Array.isArray(details?.versions)
        ? details.versions.find((v) => String(v.version) === String(latestVersion))?.scanStatus || null
        : null;

    const badges = el("div", { className: "badges", style: "display:flex; gap:6px; flex-wrap:wrap; margin-top:6px;" });
    const verified = Boolean(item.verified ?? details?.verified);
    const featured = Boolean(item.featured ?? details?.featured);
    const deprecated = Boolean(details?.deprecated ?? item.deprecated);
    const blocked = Boolean(details?.blocked ?? item.blocked);
    const malicious = Boolean(details?.malicious ?? item.malicious);
    const publisherRevoked = Boolean(details?.publisherRevoked);

    if (verified) badges.append(badge("verified", { tone: "good" }));
    if (featured) badges.append(badge("featured", { tone: "good" }));
    if (deprecated) badges.append(badge("deprecated", { tone: "warn" }));
    if (blocked) badges.append(badge("blocked", { tone: "bad" }));
    if (malicious) badges.append(badge("malicious", { tone: "bad" }));
    if (publisherRevoked) badges.append(badge("revoked", { tone: "bad" }));
    if (installed?.corrupted) {
      const reason =
        installed.corruptedReason && typeof installed.corruptedReason === "string" ? installed.corruptedReason : "Unknown reason";
      const at = installed.corruptedAt && typeof installed.corruptedAt === "string" ? installed.corruptedAt : null;
      badges.append(
        badge("corrupted", {
          tone: "bad",
          title: at ? `Corrupted at ${at}: ${reason}` : `Corrupted: ${reason}`,
        }),
      );
    }
    if (installed?.incompatible) {
      const reason =
        installed.incompatibleReason && typeof installed.incompatibleReason === "string"
          ? installed.incompatibleReason
          : "Unknown reason";
      const at = installed.incompatibleAt && typeof installed.incompatibleAt === "string" ? installed.incompatibleAt : null;
      badges.append(
        badge("incompatible", {
          tone: "warn",
          title: at ? `Marked incompatible at ${at}: ${reason}` : `Marked incompatible: ${reason}`,
        }),
      );
    }
    if (latestScanStatusRaw) {
      const normalized = String(latestScanStatusRaw).trim().toLowerCase();
      const tone =
        normalized === "passed"
          ? "good"
          : normalized === "pending" || normalized === "unknown"
            ? "warn"
            : "bad";
      badges.append(badge(`scan: ${latestScanStatusRaw}`, { tone }));
    }

    const policy = describePolicy(details, extensionManager?.scanPolicy);

    const row = el("div", { className: "marketplace-result", dataset: { testid: `marketplace-result-${item.id}` } }, [
      el("div", { className: "title" }, [document.createTextNode(`${item.displayName} (${item.id})`)]),
      el("div", { className: "desc" }, [document.createTextNode(item.description || "")]),
      badges,
    ]);

    if (policy.reason || policy.warning) {
      const pieces = [];
      if (policy.reason) pieces.push(policy.reason);
      if (policy.warning) pieces.push(policy.warning);
      row.append(el("div", { className: "policy" }, [document.createTextNode(pieces.join(" · "))]));
    }

    const actions = el("div", { className: "actions" });
    if (!installed) {
      if (typeof transientStatus === "string" && transientStatus.trim()) {
        actions.append(
          el("div", { className: "installed-meta" }, [document.createTextNode(String(transientStatus))]),
        );
      }
        actions.append(
          el(
            "button",
            {
              disabled: policy.blocked || blocked || malicious || publisherRevoked ? "true" : undefined,
              dataset: { testid: `marketplace-install-${item.id}` },
              onClick: async () => {
                actions.textContent = "Installing…";
                try {
                const record = await extensionManager.install(item.id, null, {
                  confirm: async (warning) => {
                    const message = `${warning.message}\n\nProceed with install?`;

                    // Prefer native dialogs in desktop builds when available.
                    if (getTauriDialogConfirmOrNull()) {
                      return await nativeDialogs.confirm(message, { title: "Install extension" });
                    }

                    // Web builds: use the non-blocking <dialog>-based picker.
                    if (typeof document === "undefined" || !document.body) return true;
                    const ok = await showQuickPick(
                      [
                        { label: "Proceed", value: true },
                        { label: "Cancel", value: false },
                      ],
                      { placeHolder: message },
                    );
                    return ok ?? false;
                  },
                });
                if (Array.isArray(record?.warnings)) {
                  for (const warning of record.warnings) {
                    if (!warning || typeof warning.message !== "string") continue;
                    tryShowToast(warning.message, "warning");
                  }
                }

                if (extensionHostManager?.syncInstalledExtensions) {
                  await extensionHostManager.syncInstalledExtensions();
                } else if (extensionHostManager) {
                  await extensionHostManager.reloadExtension(item.id);
                }
                updateContributedPanelSeedsFromHost(extensionHostManager, item.id);
                actions.textContent = "Installed";
                tryNotifyExtensionsChanged();
                await renderSearchResults({
                  container,
                  marketplaceClient,
                  extensionManager,
                  extensionHostManager,
                  query,
                  transientStatusById,
                });
              } catch (error) {
                // eslint-disable-next-line no-console
                console.error(error);
                tryShowToast(String(error?.message ?? error), "error");
                actions.textContent = `Error: ${String(error?.message ?? error)}`;
                await renderSearchResults({
                  container,
                  marketplaceClient,
                  extensionManager,
                  extensionHostManager,
                  query,
                  transientStatusById,
                }).catch(() => {});
              }
            },
          },
          [document.createTextNode("Install")],
        ),
      );
    } else {
      const metaParts = ["Installed"];
      if (installed.version) metaParts.push(`v${installed.version}`);
      if (installed.scanStatus) metaParts.push(`scan=${installed.scanStatus}`);
      if (installed.signingKeyId) metaParts.push(`key=${installed.signingKeyId}`);
      if (metaParts.length > 0) {
        actions.append(el("div", { className: "installed-meta" }, [document.createTextNode(metaParts.join(" · "))]));
      }
      if (transientStatusById && transientStatusById.has(item.id)) {
        // Any "recently uninstalled" state is no longer relevant if the extension is installed again.
        transientStatusById.delete(item.id);
      }

      if (installed.corrupted || installed.incompatible) {
        const installedVersion = installed?.version ? String(installed.version) : null;
        const shouldTryUpdate = Boolean(installed?.incompatible);
        const incompatibleReason =
          installed?.incompatibleReason && typeof installed.incompatibleReason === "string"
            ? installed.incompatibleReason
            : "";
        const isEngineMismatch = incompatibleReason.toLowerCase().includes("engine mismatch");
        actions.append(
          el(
            "button",
            {
              dataset: { testid: `marketplace-repair-${item.id}` },
              onClick: async () => {
                actions.textContent = shouldTryUpdate ? "Updating…" : "Repairing…";
                try {
                  if (extensionHostManager) {
                    await extensionHostManager.unloadExtension(item.id);
                    await extensionHostManager.resetExtensionState?.(item.id);
                  }
                } catch {
                  // ignore
                }

                try {
                  let record;
                  if (shouldTryUpdate && typeof extensionManager.update === "function") {
                    try {
                      record = await extensionManager.update(item.id);
                    } catch (error) {
                      const msg = String(error?.message ?? error);
                      if (/engine mismatch/i.test(msg)) {
                        // If the extension is incompatible due to an engine mismatch, there is no
                        // recovery path without a compatible update (or changing engine version).
                        if (isEngineMismatch) {
                          actions.textContent = "No compatible update";
                          tryShowToast("No compatible update", "warning");
                          tryNotifyExtensionsChanged();
                          await renderSearchResults({
                            container,
                            marketplaceClient,
                            extensionManager,
                            extensionHostManager,
                            query,
                            transientStatusById,
                          });
                          return;
                        }

                        // If the extension is marked incompatible for a reason other than engine
                        // mismatch (eg: corrupted stored manifest), but the latest version cannot
                        // be installed due to an engine mismatch, fall back to reinstalling the
                        // currently-installed version so users still have a recovery path.
                        if (typeof extensionManager.repair === "function") {
                          record = await extensionManager.repair(item.id);
                        } else {
                          throw error;
                        }
                      } else {
                        throw error;
                      }
                    }
                    // If the update is a no-op (already on the latest version), fall back to a
                    // repair/reinstall so users still have a recovery path when an incompatible
                    // quarantine is caused by corrupted manifest metadata rather than an engine
                    // mismatch.
                    if (
                      installedVersion &&
                      record &&
                      String(record.version ?? "") === installedVersion &&
                      typeof extensionManager.repair === "function" &&
                      !isEngineMismatch
                    ) {
                      record = await extensionManager.repair(item.id);
                    }
                  } else if (typeof extensionManager.repair === "function") {
                    record = await extensionManager.repair(item.id);
                  } else {
                    record = await extensionManager.install(item.id);
                  }

                  if (Array.isArray(record?.warnings)) {
                    for (const warning of record.warnings) {
                      if (!warning || typeof warning.message !== "string") continue;
                      tryShowToast(warning.message, "warning");
                    }
                  }

                  if (extensionHostManager?.syncInstalledExtensions) {
                    await extensionHostManager.syncInstalledExtensions();
                  } else if (extensionHostManager) {
                    await extensionHostManager.reloadExtension(item.id);
                  }

                  updateContributedPanelSeedsFromHost(extensionHostManager, item.id);

                  const recordVersion = record?.version != null ? String(record.version) : "";
                  const didUpdate = Boolean(installedVersion && recordVersion && recordVersion !== installedVersion);
                  const wasNoOpUpdate = Boolean(shouldTryUpdate && installedVersion && recordVersion === installedVersion);
                  actions.textContent = isEngineMismatch && wasNoOpUpdate ? "No compatible update" : didUpdate ? "Updated" : "Repaired";
                  if (isEngineMismatch && wasNoOpUpdate) {
                    tryShowToast("No compatible update", "warning");
                  }

                  tryNotifyExtensionsChanged();
                  await renderSearchResults({
                    container,
                    marketplaceClient,
                    extensionManager,
                    extensionHostManager,
                    query,
                    transientStatusById,
                  });
                } catch (error) {
                  // eslint-disable-next-line no-console
                  console.error(error);
                  tryShowToast(String(error?.message ?? error), "error");
                  actions.textContent = `Error: ${String(error?.message ?? error)}`;
                  await renderSearchResults({
                    container,
                    marketplaceClient,
                    extensionManager,
                    extensionHostManager,
                    query,
                    transientStatusById,
                  }).catch(() => {});
                }
              },
            },
            [document.createTextNode("Repair")],
          ),
        );
      }

      actions.append(
        el(
          "button",
          {
            dataset: { testid: `marketplace-uninstall-${item.id}` },
            onClick: async () => {
              actions.textContent = "Uninstalling…";
              try {
                if (extensionHostManager) {
                  await extensionHostManager.unloadExtension(item.id);
                  await extensionHostManager.resetExtensionState?.(item.id);
                }
                await extensionManager.uninstall(item.id);
                if (extensionHostManager?.syncInstalledExtensions) {
                  await extensionHostManager.syncInstalledExtensions();
                }

                const storage = getDefaultSeedStoreStorage();
                if (storage) removeSeedPanelsForExtension(storage, item.id);
                if (transientStatusById) transientStatusById.set(item.id, "Uninstalled");

                actions.textContent = "Uninstalled";
                tryNotifyExtensionsChanged();
                await renderSearchResults({
                  container,
                  marketplaceClient,
                  extensionManager,
                  extensionHostManager,
                  query,
                  transientStatusById,
                });
              } catch (error) {
                // eslint-disable-next-line no-console
                console.error(error);
                tryShowToast(String(error?.message ?? error), "error");
                actions.textContent = `Error: ${String(error?.message ?? error)}`;
                await renderSearchResults({
                  container,
                  marketplaceClient,
                  extensionManager,
                  extensionHostManager,
                  query,
                  transientStatusById,
                }).catch(() => {});
              }
            },
          },
          [document.createTextNode("Uninstall")],
        ),
      );

      actions.append(
        el(
          "button",
          {
            dataset: { testid: `marketplace-update-${item.id}` },
            onClick: async () => {
              actions.textContent = "Checking…";
              try {
                const updates = await extensionManager.checkForUpdates();
                const update = updates.find((u) => u.id === item.id);
                if (!update) {
                  actions.textContent = "Up to date";
                  tryShowToast("Up to date", "info");
                  await renderSearchResults({
                    container,
                    marketplaceClient,
                    extensionManager,
                    extensionHostManager,
                    query,
                    transientStatusById,
                  });
                  return;
                }
                actions.textContent = `Updating to ${update.latestVersion}…`;

                // Terminate the running extension before mutating its package in IndexedDB.
                if (extensionHostManager) {
                  await extensionHostManager.unloadExtension(item.id);
                }

                const record = await extensionManager.update(item.id);
                if (Array.isArray(record?.warnings)) {
                  for (const warning of record.warnings) {
                    if (!warning || typeof warning.message !== "string") continue;
                    tryShowToast(warning.message, "warning");
                  }
                }

                if (extensionHostManager?.syncInstalledExtensions) {
                  await extensionHostManager.syncInstalledExtensions();
                } else if (extensionHostManager) {
                  await extensionHostManager.reloadExtension(item.id);
                }

                updateContributedPanelSeedsFromHost(extensionHostManager, item.id);
                actions.textContent = "Updated";
                tryNotifyExtensionsChanged();
                await renderSearchResults({
                  container,
                  marketplaceClient,
                  extensionManager,
                  extensionHostManager,
                  query,
                  transientStatusById,
                });
              } catch (error) {
                // eslint-disable-next-line no-console
                const msg = String(error?.message ?? error);
                if (/engine mismatch/i.test(msg)) {
                  actions.textContent = "No compatible update";
                  tryShowToast("No compatible update", "warning");
                } else {
                  console.error(error);
                  tryShowToast(msg, "error");
                  actions.textContent = `Error: ${msg}`;
                }
                await renderSearchResults({
                  container,
                  marketplaceClient,
                  extensionManager,
                  extensionHostManager,
                  query,
                  transientStatusById,
                });
              }
            },
          },
          [document.createTextNode("Update")],
        ),
      );
    }

    row.append(actions);
    list.append(row);
  }

  container.textContent = "";
  container.append(list);
}

export function createMarketplacePanel({
  container,
  marketplaceClient,
  extensionManager,
  extensionHostManager,
}) {
  // Ephemeral UI state for showing a status in the search results after actions like uninstall.
  // Cleared whenever the user manually triggers a new search.
  const transientStatusById = new Map();
  const queryInput = el("input", {
    type: "search",
    placeholder: "Search extensions…",
    dataset: { testid: "marketplace-search-input" },
  });
  const resultsContainer = el("div", { className: "results" });

  const searchButton = el(
    "button",
    {
      dataset: { testid: "marketplace-search-button" },
      onClick: async () => {
        transientStatusById.clear();
        try {
          await renderSearchResults({
            container: resultsContainer,
            marketplaceClient,
            extensionManager,
            extensionHostManager,
            query: queryInput.value,
            transientStatusById,
          });
        } catch (error) {
          // eslint-disable-next-line no-console
          console.error(error);
          const msg = String(error?.message ?? error);
          tryShowToast(msg, "error");
          resultsContainer.textContent = `Error: ${msg}`;
        }
      },
    },
    [document.createTextNode("Search")],
  );

  container.append(el("div", { className: "marketplace-panel" }, [queryInput, searchButton, resultsContainer]));

  return {
    dispose() {
      container.textContent = "";
    },
  };
}
