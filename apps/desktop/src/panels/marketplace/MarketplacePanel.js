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
import { showToast } from "../../extensions/ui.js";

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

async function renderSearchResults({ container, marketplaceClient, extensionManager, extensionHostManager, query }) {
  container.textContent = "Searching…";
  const results = await marketplaceClient.search({ q: query, limit: 25, offset: 0 });

  const list = el("div", { className: "marketplace-results" });

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
    const installed = await extensionManager.getInstalled(item.id);
    const details = detailsById.get(item.id) || null;

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

    const row = el("div", { className: "marketplace-result" }, [
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
      actions.append(
        el(
          "button",
          {
            disabled: policy.blocked ? "true" : undefined,
            onClick: async () => {
              actions.textContent = "Installing…";
              try {
                const record = await extensionManager.install(item.id, null, {
                  confirm: async (warning) => {
                    // Best-effort: use a browser confirm prompt (some environments may not allow it).
                    try {
                      if (typeof window?.confirm === "function") {
                        return window.confirm(`${warning.message}\n\nProceed with install?`);
                      }
                    } catch {
                      // ignore
                    }
                    return true;
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
              } catch (error) {
                // eslint-disable-next-line no-console
                console.error(error);
                tryShowToast(String(error?.message ?? error), "error");
                actions.textContent = `Error: ${String(error?.message ?? error)}`;
              }
            },
          },
          [document.createTextNode("Install")],
        ),
      );
    } else {
      const metaParts = [];
      if (installed.version) metaParts.push(`v${installed.version}`);
      if (installed.scanStatus) metaParts.push(`scan=${installed.scanStatus}`);
      if (installed.signingKeyId) metaParts.push(`key=${installed.signingKeyId}`);
      if (metaParts.length > 0) {
        actions.append(el("div", { className: "installed-meta" }, [document.createTextNode(metaParts.join(" · "))]));
      }

      actions.append(
        el(
          "button",
          {
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

                actions.textContent = "Uninstalled";
                tryNotifyExtensionsChanged();
              } catch (error) {
                // eslint-disable-next-line no-console
                console.error(error);
                tryShowToast(String(error?.message ?? error), "error");
                actions.textContent = `Error: ${String(error?.message ?? error)}`;
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
            onClick: async () => {
              actions.textContent = "Checking…";
              try {
                const updates = await extensionManager.checkForUpdates();
                const update = updates.find((u) => u.id === item.id);
                if (!update) {
                  actions.textContent = "Up to date";
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
              } catch (error) {
                // eslint-disable-next-line no-console
                console.error(error);
                tryShowToast(String(error?.message ?? error), "error");
                actions.textContent = `Error: ${String(error?.message ?? error)}`;
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
  const queryInput = el("input", { type: "search", placeholder: "Search extensions…" });
  const resultsContainer = el("div", { className: "results" });

  const searchButton = el(
    "button",
    {
      onClick: async () => {
        await renderSearchResults({
          container: resultsContainer,
          marketplaceClient,
          extensionManager,
          extensionHostManager,
          query: queryInput.value,
        });
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
