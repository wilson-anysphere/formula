import { showToast } from "../extensions/ui.js";
import { t, tWithVars } from "../i18n/index.js";
import { markKeybindingBarrier } from "../keybindingBarrier.js";

import { requestAppRestart } from "./appQuit";
import { notify } from "./notifications";
import { getTauriEventApiOrNull, getTauriWindowHandleOrNull, type TauriListen } from "./api";
import { shellOpen } from "./shellOpen";
import { installUpdateAndRestart } from "./updater";

export const FORMULA_RELEASES_URL = "https://github.com/wilson-anysphere/formula/releases";

export const UPDATER_DISMISSED_VERSION_KEY = "formula.updater.dismissedVersion";
export const UPDATER_DISMISSED_AT_KEY = "formula.updater.dismissedAt";
const UPDATER_DISMISSAL_TTL_MS = 7 * 24 * 60 * 60 * 1000;

function getTauriGlobalOrNull(): any | null {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (globalThis as any).__TAURI__ ?? null;
  } catch {
    // Some hardened host environments (or tests) may define `__TAURI__` with a throwing getter.
    // Treat that as "unavailable" so best-effort callsites can fall back cleanly.
    return null;
  }
}

function safeGetProp(obj: any, prop: string): any | undefined {
  if (!obj) return undefined;
  try {
    return obj[prop];
  } catch {
    return undefined;
  }
}

type UpdaterEventName =
  | "update-check-already-running"
  | "update-check-started"
  | "update-not-available"
  | "update-check-error"
  | "update-available"
  | "update-download-started"
  | "update-download-progress"
  | "update-downloaded"
  | "update-download-error";

type UpdaterEventPayload = {
  source?: string;
  version?: string;
  body?: string | null;
  message?: string;
  error?: string;
  // Optional background download progress metadata.
  downloaded?: number;
  total?: number | null;
  percent?: number | null;
  chunkLength?: number;
  // Optional manual download metadata (may be added to updater payloads in the future).
  releaseUrl?: string;
  release_url?: string;
  homepage?: string;
  homepageUrl?: string;
  homepage_url?: string;
  url?: string;
  downloadUrl?: string;
  download_url?: string;
};

type StorageLike = Pick<Storage, "getItem" | "setItem" | "removeItem">;

type DismissalRecord = { version: string; dismissedAtMs: number };

type UpdaterDownloadProgress = {
  downloaded?: number;
  total?: number;
  percent?: number;
  current?: number;
  chunkLength?: number;
  contentLength?: number;
  chunk_length?: number;
  content_length?: number;
  // Tauri v2 updater plugin download events are sent as `{ event, data }` objects.
  // We intentionally keep this type loose and extract numbers defensively.
  data?: {
    chunkLength?: number;
    contentLength?: number;
    chunk_length?: number;
    content_length?: number;
  };
};

type UpdaterUpdate = {
  version: string;
  body?: string | null;
  download(onProgress?: (progress: UpdaterDownloadProgress) => void): Promise<void>;
  install(): Promise<void>;
};

type DialogElements = {
  dialog: HTMLDialogElement;
  title: HTMLElement;
  version: HTMLElement;
  releaseNotesTitle: HTMLElement;
  body: HTMLElement;
  status: HTMLElement;
  progressWrap: HTMLElement;
  progressBar: HTMLProgressElement;
  progressText: HTMLElement;
  downloadBtn: HTMLButtonElement;
  laterBtn: HTMLButtonElement;
  viewVersionsBtn: HTMLButtonElement;
  restartBtn: HTMLButtonElement;
};

let updateDialog: DialogElements | null = null;
let updateInfo: { version: string; body: string | null; manualDownloadUrl: string } | null = null;
let downloadedUpdate: UpdaterUpdate | null = null;
let downloadInFlight = false;
let lastUpdateError: string | null = null;

let progressDownloaded = 0;
let progressTotal: number | null = null;
let progressPercent: number | null = null;

let updateReadyToastShownForVersion: string | null = null;
let backendDownloadedVersion: string | null = null;
let backendDownloadInFlightVersion: string | null = null;
function getLocalStorageOrNull(): StorageLike | null {
  try {
    // Prefer `window.localStorage` when available (jsdom/webview), but fall back to
    // `globalThis.localStorage`. Node 22+ ships an experimental localStorage accessor that throws
    // unless started with `--localstorage-file`, so probe before returning.
    const storage =
      (typeof window !== "undefined" ? ((window as any).localStorage as StorageLike | undefined) : undefined) ??
      ((globalThis as any).localStorage as StorageLike | undefined) ??
      null;
    if (!storage) return null;
    storage.getItem("formula.updater.storageProbe");
    return storage;
  } catch {
    return null;
  }
}

function readUpdaterDismissal(storage: StorageLike): DismissalRecord | null {
  try {
    const version = storage.getItem(UPDATER_DISMISSED_VERSION_KEY);
    if (typeof version !== "string" || version.trim() === "") return null;
    const dismissedAtRaw = storage.getItem(UPDATER_DISMISSED_AT_KEY);
    const dismissedAtMs = Number(dismissedAtRaw);
    if (!Number.isFinite(dismissedAtMs)) return null;
    return { version: version.trim(), dismissedAtMs };
  } catch {
    return null;
  }
}

function clearUpdaterDismissal(storage: StorageLike): void {
  try {
    storage.removeItem(UPDATER_DISMISSED_VERSION_KEY);
    storage.removeItem(UPDATER_DISMISSED_AT_KEY);
  } catch {
    // ignore
  }
}

function setUpdaterDismissal(storage: StorageLike, version: string): void {
  if (!version.trim()) return;
  try {
    storage.setItem(UPDATER_DISMISSED_VERSION_KEY, version.trim());
    storage.setItem(UPDATER_DISMISSED_AT_KEY, String(Date.now()));
  } catch {
    // ignore
  }
}

function shouldSuppressStartupUpdatePrompt(version: string): boolean {
  const storage = getLocalStorageOrNull();
  if (!storage) return false;

  const record = readUpdaterDismissal(storage);
  if (!record) return false;

  // New version => reset suppression.
  if (record.version !== version) {
    clearUpdaterDismissal(storage);
    return false;
  }

  const ageMs = Date.now() - record.dismissedAtMs;
  if (!Number.isFinite(ageMs) || ageMs < 0) {
    clearUpdaterDismissal(storage);
    return false;
  }

  if (ageMs < UPDATER_DISMISSAL_TTL_MS) return true;

  clearUpdaterDismissal(storage);
  return false;
}

function persistDismissalForCurrentUpdate(): void {
  const version = updateInfo?.version ?? "";
  if (!version.trim()) return;
  const storage = getLocalStorageOrNull();
  if (!storage) return;
  setUpdaterDismissal(storage, version);
}

function clearDismissalOnUpdateInitiated(): void {
  const storage = getLocalStorageOrNull();
  if (!storage) return;
  clearUpdaterDismissal(storage);
}

// If the user triggers a manual check while a background (startup) check is already in-flight,
// the backend emits `update-check-already-running`. Track that a manual request is "waiting" so
// we can surface the eventual completion event (no-update / error) even if it comes from a
// startup-sourced check.
let manualUpdateCheckFollowUp = false;
let manualUpdateCheckFollowUpTimeout: ReturnType<typeof setTimeout> | null = null;

function setManualUpdateCheckFollowUp(active: boolean): void {
  manualUpdateCheckFollowUp = active;
  if (manualUpdateCheckFollowUpTimeout) {
    clearTimeout(manualUpdateCheckFollowUpTimeout);
    manualUpdateCheckFollowUpTimeout = null;
  }
  if (active) {
    // Best-effort reset so a stuck/failed updater check doesn't cause us to surface future
    // automatic checks as if they were manual.
    manualUpdateCheckFollowUpTimeout = setTimeout(() => {
      manualUpdateCheckFollowUp = false;
      manualUpdateCheckFollowUpTimeout = null;
    }, 120_000);
  }
}
function getTauriListen(): TauriListen | null {
  return getTauriEventApiOrNull()?.listen ?? null;
}

function getTauriWindowHandle(): any | null {
  return getTauriWindowHandleOrNull();
}

async function showMainWindowBestEffort(): Promise<void> {
  const win = getTauriWindowHandle();
  if (!win) return;

  try {
    if (typeof win.show === "function") {
      await win.show();
    }
  } catch {
    // Best-effort.
  }

  try {
    if (typeof win.setFocus === "function") {
      await win.setFocus();
    }
  } catch {
    // Best-effort.
  }
}

async function openExternalUrl(url: string): Promise<void> {
  try {
    await shellOpen(url);
  } catch {
    // Best-effort.
  }
}

function isHttpUrl(value: string): boolean {
  try {
    const parsed = new URL(value);
    return parsed.protocol === "http:" || parsed.protocol === "https:";
  } catch {
    return false;
  }
}

/**
 * Best-effort resolver for a human-friendly "manual download" URL.
 *
 * If updater metadata includes a homepage/release URL, we prefer that. Otherwise we
 * fall back to the GitHub releases page.
 */
export function resolveUpdateReleaseUrl(update: unknown): string {
  if (typeof update === "string") {
    const trimmed = update.trim();
    if (trimmed && isHttpUrl(trimmed)) return trimmed;
    return FORMULA_RELEASES_URL;
  }

  if (!update || typeof update !== "object") return FORMULA_RELEASES_URL;
  const record = update as Record<string, unknown>;
  const candidates: unknown[] = [
    record.releaseUrl,
    record.release_url,
    record.homepage,
    record.homepageUrl,
    record.homepage_url,
    // Some updater manifests include `url` / `download_url` which might be a direct
    // artifact download; still useful as a manual escape hatch.
    record.url,
    record.downloadUrl,
    record.download_url,
  ];

  for (const candidate of candidates) {
    if (typeof candidate !== "string") continue;
    const trimmed = candidate.trim();
    if (!trimmed) continue;
    if (!isHttpUrl(trimmed)) continue;
    return trimmed;
  }

  return FORMULA_RELEASES_URL;
}

export async function openUpdateReleasePage(update: unknown): Promise<void> {
  const url = resolveUpdateReleaseUrl(update);
  await shellOpen(url);
}

function getUpdaterUpdateOrNull(raw: unknown): UpdaterUpdate | null {
  if (!raw || typeof raw !== "object") return null;
  const obj = raw as any;
  const version = typeof obj.version === "string" ? obj.version : "";
  if (!version.trim()) return null;
  const body = typeof obj.body === "string" ? obj.body : null;

  const download = obj.download as ((cb?: unknown) => Promise<void>) | undefined;
  const install = obj.install as (() => Promise<void>) | undefined;
  if (typeof download !== "function" || typeof install !== "function") return null;

  return {
    version,
    body,
    download: async (onProgress) => {
      await download.call(obj, onProgress);
    },
    install: async () => {
      await install.call(obj);
    },
  };
}

function getUpdaterApiOrNull(): { check: () => Promise<UpdaterUpdate | null> } | null {
  const tauri = getTauriGlobalOrNull();
  const plugin = safeGetProp(tauri, "plugin");
  const plugins = safeGetProp(tauri, "plugins");
  const updater = safeGetProp(tauri, "updater") ?? safeGetProp(plugin, "updater") ?? safeGetProp(plugins, "updater") ?? null;
  if (!updater) return null;

  const check = safeGetProp(updater, "check") as (() => Promise<unknown>) | undefined;
  if (typeof check !== "function") return null;

  return {
    check: async () => getUpdaterUpdateOrNull(await check.call(updater)),
  };
}

function safeShowModal(dialog: HTMLDialogElement): void {
  // `showModal()` throws if called on an already-open dialog.
  if ((dialog as any).open === true || dialog.hasAttribute("open")) return;
  const showModal = (dialog as any).showModal as (() => void) | undefined;
  if (typeof showModal === "function") showModal.call(dialog);
  else dialog.setAttribute("open", "");
}

function safeClose(dialog: HTMLDialogElement, returnValue = ""): void {
  const close = (dialog as any).close as ((returnValue?: string) => void) | undefined;
  if (typeof close === "function") close.call(dialog, returnValue);
  else dialog.removeAttribute("open");
}

function clampPercent(value: number): number {
  if (!Number.isFinite(value)) return 0;
  if (value <= 0) return 0;
  if (value >= 100) return 100;
  return value;
}

function extractNumber(value: unknown): number | null {
  const num = typeof value === "number" ? value : typeof value === "string" ? Number(value) : NaN;
  return Number.isFinite(num) ? num : null;
}

function errorMessage(err: unknown): string {
  if (typeof err === "string") return err;
  if (err instanceof Error) return err.message;
  if (err && typeof err === "object" && "message" in err) {
    try {
      return String((err as any).message);
    } catch {
      // fall through
    }
  }
  try {
    return String(err);
  } catch {
    return t("updater.unknownError");
  }
}

function ensureUpdateDialog(): DialogElements {
  if (updateDialog) {
    // Tests (and potentially custom host integrations) may remove the dialog from the DOM.
    // If it exists but is detached, reattach so update-available events can still surface UI.
    if (!updateDialog.dialog.isConnected && document.body) {
      try {
        document.body.appendChild(updateDialog.dialog);
      } catch {
        // Best-effort.
      }
    }
    return updateDialog;
  }

  const dialog = document.createElement("dialog");
  dialog.className = "dialog updater-dialog";
  dialog.dataset.testid = "updater-dialog";
  markKeybindingBarrier(dialog);

  const title = document.createElement("div");
  title.className = "dialog__title";
  title.textContent = t("updater.updateAvailableTitle");

  const version = document.createElement("div");
  version.className = "updater-dialog__version";
  version.dataset.testid = "updater-version";

  const releaseNotesTitle = document.createElement("div");
  releaseNotesTitle.className = "updater-dialog__release-notes-title";
  releaseNotesTitle.dataset.testid = "updater-release-notes-title";
  releaseNotesTitle.textContent = t("updater.releaseNotes");

  const body = document.createElement("pre");
  body.className = "updater-dialog__body";
  body.dataset.testid = "updater-body";

  const status = document.createElement("div");
  status.className = "updater-dialog__status";
  status.dataset.testid = "updater-status";

  const progressWrap = document.createElement("div");
  progressWrap.className = "updater-dialog__progress-wrap";
  progressWrap.dataset.testid = "updater-progress-wrap";
  progressWrap.hidden = true;

  const progressBar = document.createElement("progress");
  progressBar.className = "updater-dialog__progress";
  progressBar.dataset.testid = "updater-progress";
  progressBar.max = 100;
  progressBar.value = 0;

  const progressText = document.createElement("div");
  progressText.className = "updater-dialog__progress-text";
  progressText.dataset.testid = "updater-progress-text";

  progressWrap.appendChild(progressBar);
  progressWrap.appendChild(progressText);

  const controls = document.createElement("div");
  controls.className = "dialog__controls updater-dialog__controls";

  const laterBtn = document.createElement("button");
  laterBtn.type = "button";
  laterBtn.textContent = t("updater.later");
  laterBtn.dataset.testid = "updater-later";

  const viewVersionsBtn = document.createElement("button");
  viewVersionsBtn.type = "button";
  viewVersionsBtn.textContent = t("updater.openReleasePage");
  viewVersionsBtn.dataset.testid = "updater-view-versions";

  const downloadBtn = document.createElement("button");
  downloadBtn.type = "button";
  downloadBtn.textContent = t("updater.download");
  downloadBtn.dataset.testid = "updater-download";

  const restartBtn = document.createElement("button");
  restartBtn.type = "button";
  restartBtn.textContent = t("updater.restartNow");
  restartBtn.dataset.testid = "updater-restart";
  restartBtn.hidden = true;

  controls.appendChild(laterBtn);
  controls.appendChild(viewVersionsBtn);
  controls.appendChild(downloadBtn);
  controls.appendChild(restartBtn);

  dialog.appendChild(title);
  dialog.appendChild(version);
  dialog.appendChild(releaseNotesTitle);
  dialog.appendChild(body);
  dialog.appendChild(status);
  dialog.appendChild(progressWrap);
  dialog.appendChild(controls);

  dialog.addEventListener("cancel", (e) => {
    if (downloadInFlight) {
      e.preventDefault();
      return;
    }
    // Allow ESC to behave like "Later".
    e.preventDefault();
    persistDismissalForCurrentUpdate();
    safeClose(dialog, "later");
  });

  laterBtn.addEventListener("click", () => {
    if (downloadInFlight) return;
    persistDismissalForCurrentUpdate();
    safeClose(dialog, "later");
  });

  viewVersionsBtn.addEventListener("click", () => {
    const url = updateInfo?.manualDownloadUrl ?? FORMULA_RELEASES_URL;
    void openExternalUrl(url);
    if (!downloadInFlight) safeClose(dialog, "versions");
  });

  downloadBtn.addEventListener("click", () => {
    clearDismissalOnUpdateInitiated();
    void startUpdateDownload();
  });

  restartBtn.addEventListener("click", () => {
    void (async () => {
      clearDismissalOnUpdateInitiated();
      lastUpdateError = null;
      renderUpdateDialog();
      // Keep the dialog open if the restart was cancelled (e.g. user hit "Cancel"
      // on the unsaved-changes prompt).
      const didRestart = await restartToInstallUpdate();
      if (didRestart) safeClose(dialog, "restart");
      else if (lastUpdateError) {
        renderUpdateDialog();
        try {
          viewVersionsBtn.focus();
        } catch {
          // Best-effort focus.
        }
      }
    })().catch(() => {
      // Best-effort: avoid unhandled rejections from the async restart handler.
    });
  });

  document.body.appendChild(dialog);

  updateDialog = {
    dialog,
    title,
    version,
    releaseNotesTitle,
    body,
    status,
    progressWrap,
    progressBar,
    progressText,
    downloadBtn,
    laterBtn,
    viewVersionsBtn,
    restartBtn,
  };

  return updateDialog;
}

function renderUpdateDialog(): void {
  const els = ensureUpdateDialog();
  const info = updateInfo;
  const backendDownloadInFlight =
    backendDownloadInFlightVersion != null && backendDownloadInFlightVersion === info?.version;
  const anyDownloadInFlight = downloadInFlight || backendDownloadInFlight;

  els.title.textContent = t("updater.updateAvailableTitle");
  els.version.textContent = info ? tWithVars("updater.updateAvailableMessage", { version: info.version }) : "";
  els.body.textContent = info?.body ? info.body : "";
  els.releaseNotesTitle.textContent = t("updater.releaseNotes");
  els.releaseNotesTitle.hidden = !info?.body;

  els.laterBtn.disabled = anyDownloadInFlight;
  // Keep "Open release page" enabled even while a download is in flight so users can
  // inspect the release notes/versions without interrupting the download. The click
  // handler already avoids closing the dialog while `downloadInFlight` is true.
  els.viewVersionsBtn.disabled = false;
  els.downloadBtn.disabled =
    anyDownloadInFlight ||
    !!downloadedUpdate ||
    (backendDownloadedVersion != null && backendDownloadedVersion === info?.version);
  els.laterBtn.textContent = t("updater.later");
  els.viewVersionsBtn.textContent = t("updater.openReleasePage");
  els.downloadBtn.textContent = anyDownloadInFlight ? t("updater.downloading") : t("updater.download");

  const readyToInstall =
    !anyDownloadInFlight &&
    (!!downloadedUpdate || (backendDownloadedVersion != null && backendDownloadedVersion === info?.version));
  els.restartBtn.hidden = !readyToInstall;
  els.restartBtn.disabled = false;

  if (anyDownloadInFlight) {
    els.status.textContent = t("updater.downloadInProgress");
    els.progressWrap.hidden = false;
    renderProgress();
  } else if (lastUpdateError) {
    els.progressWrap.hidden = true;
    els.status.textContent = downloadedUpdate
      ? tWithVars("updater.installFailedWithMessage", { message: lastUpdateError })
      : tWithVars("updater.downloadFailedWithMessage", { message: lastUpdateError });
  } else if (downloadedUpdate || (backendDownloadedVersion != null && backendDownloadedVersion === info?.version)) {
    els.status.textContent = `${t("updater.downloadComplete")} ${t("updater.restartToInstall")}`;
    els.progressWrap.hidden = true;
  } else {
    els.status.textContent = "";
    els.progressWrap.hidden = true;
  }

  if (lastUpdateError) {
    els.viewVersionsBtn.textContent = t("updater.downloadManually");
    els.viewVersionsBtn.classList.add("updater-dialog__view-versions--primary");
  } else {
    els.viewVersionsBtn.textContent = t("updater.openReleasePage");
    els.viewVersionsBtn.classList.remove("updater-dialog__view-versions--primary");
  }
  els.restartBtn.textContent = t("updater.restartNow");
}

function renderProgress(): void {
  const els = ensureUpdateDialog();
  const total = progressTotal;
  const percent = progressPercent;

  if (typeof percent === "number") {
    const pct = clampPercent(percent);
    els.progressBar.max = 100;
    els.progressBar.value = pct;
    els.progressText.textContent = `${Math.round(pct)}%`;
    return;
  }

  if (typeof total === "number" && total > 0) {
    const pct = clampPercent((progressDownloaded / total) * 100);
    els.progressBar.max = 100;
    els.progressBar.value = pct;
    els.progressText.textContent = `${Math.round(pct)}%`;
    return;
  }

  // Unknown total size; show an indeterminate progress bar.
  els.progressBar.removeAttribute("value");
  els.progressText.textContent = t("updater.downloading");
}

function updateProgress(progress: UpdaterDownloadProgress): void {
  const percent = extractNumber(progress?.percent);
  if (typeof percent === "number") {
    progressPercent = clampPercent(percent);
    progressTotal = 100;
    progressDownloaded = progressPercent;
    renderProgress();
    return;
  }

  const total =
    extractNumber(progress?.total) ??
    extractNumber((progress as any)?.contentLength) ??
    extractNumber((progress as any)?.content_length) ??
    extractNumber((progress as any)?.data?.contentLength) ??
    extractNumber((progress as any)?.data?.content_length) ??
    null;
  const downloaded =
    extractNumber(progress?.downloaded) ??
    extractNumber(progress?.current) ??
    extractNumber((progress as any)?.downloaded_bytes) ??
    extractNumber((progress as any)?.downloadedBytes) ??
    null;
  const chunk =
    extractNumber((progress as any)?.chunkLength) ??
    extractNumber((progress as any)?.chunk_length) ??
    extractNumber((progress as any)?.data?.chunkLength) ??
    extractNumber((progress as any)?.data?.chunk_length) ??
    null;

  if (typeof total === "number") progressTotal = total;
  if (typeof downloaded === "number") progressDownloaded = downloaded;
  else if (typeof chunk === "number") progressDownloaded += chunk;

  renderProgress();
}

async function startUpdateDownload(): Promise<void> {
  if (downloadInFlight) return;
  if (!updateInfo) return;

  // The user initiated an update download; clear any persisted "Later" suppression.
  clearDismissalOnUpdateInitiated();

  const updater = getUpdaterApiOrNull();
  if (!updater) {
    console.warn("Updater API not available; cannot download update.");
    lastUpdateError = t("updater.unavailable");
    showToast(lastUpdateError, "error");
    renderUpdateDialog();
    try {
      ensureUpdateDialog().viewVersionsBtn.focus();
    } catch {
      // Best-effort focus.
    }
    return;
  }

  lastUpdateError = null;
  downloadInFlight = true;
  downloadedUpdate = null;
  progressDownloaded = 0;
  progressTotal = null;
  progressPercent = null;
  renderUpdateDialog();
  // Give the UI at least one tick to reflect the "Downloading…" state before we call
  // into the updater (which may complete synchronously in tests or on cached updates).
  await new Promise<void>((resolve) => setTimeout(resolve, 0));

  try {
    const update = await updater.check();
    if (!update) {
      throw new Error(t("updater.updateNoLongerAvailable"));
    }
    await update.download((progress) => {
      try {
        updateProgress((progress as any) ?? {});
      } catch (err) {
        console.warn("Failed to handle updater progress event:", err);
      }
    });
    downloadedUpdate = update;
  } catch (err) {
    console.error("Update download failed:", err);
    lastUpdateError = errorMessage(err);
    showToast(tWithVars("updater.errorWithMessage", { message: lastUpdateError }), "error");
    try {
      ensureUpdateDialog().viewVersionsBtn.focus();
    } catch {
      // Best-effort focus.
    }
  } finally {
    downloadInFlight = false;
    renderUpdateDialog();
  }
}

function openUpdateAvailableDialog(payload: UpdaterEventPayload): void {
  const version = typeof payload?.version === "string" && payload.version.trim() !== "" ? payload.version.trim() : "";
  if (!version) return;

  // If a new update arrives, clear any download state from a previous version.
  if (updateInfo?.version && updateInfo.version !== version) {
    downloadedUpdate = null;
    downloadInFlight = false;
    if (backendDownloadedVersion !== version) {
      backendDownloadedVersion = null;
    }
    if (backendDownloadInFlightVersion !== version) {
      backendDownloadInFlightVersion = null;
    }
  }

  const body = payload?.body == null ? null : String(payload.body);
  updateInfo = { version, body, manualDownloadUrl: resolveUpdateReleaseUrl(payload) };
  lastUpdateError = null;
  ensureUpdateDialog();
  renderUpdateDialog();
  safeShowModal(updateDialog!.dialog);
}

function showUpdateReadyToast(update: { version: string }): void {
  if (typeof document === "undefined") return;

  const root = document.getElementById("toast-root");
  if (!root) {
    // Fall back to a vanilla toast when running in environments without the shared toast root.
    try {
      const versionMessage = tWithVars("updater.updateAvailableMessage", { version: update.version });
      showToast(`${versionMessage}. ${t("updater.downloadComplete")} ${t("updater.restartToInstall")}`, "info");
    } catch {
      // ignore
    }
    return;
  }

  if (updateReadyToastShownForVersion === update.version) return;
  updateReadyToastShownForVersion = update.version;

  const toast = document.createElement("div");
  toast.className = "toast";
  toast.dataset.type = "info";
  toast.dataset.testid = "update-ready-toast";
  toast.style.display = "flex";
  toast.style.alignItems = "center";
  toast.style.gap = "calc(var(--space-4) + var(--space-1))";

  const message = document.createElement("div");
  const versionMessage = tWithVars("updater.updateAvailableMessage", { version: update.version });
  message.textContent = `${versionMessage}. ${t("updater.downloadComplete")} ${t("updater.restartToInstall")}`;
  message.style.flex = "1";
  toast.appendChild(message);

  const controls = document.createElement("div");
  // Reuse dialog button styles so the toast CTA matches the rest of the desktop shell.
  controls.className = "dialog__controls";
  controls.style.marginTop = "0";
  controls.style.pointerEvents = "auto";
  controls.style.flex = "0 0 auto";
  controls.style.flexWrap = "wrap";

  const laterBtn = document.createElement("button");
  laterBtn.type = "button";
  laterBtn.textContent = t("updater.later");
  laterBtn.dataset.testid = "update-ready-later";

  const viewVersionsBtn = document.createElement("button");
  viewVersionsBtn.type = "button";
  viewVersionsBtn.textContent = t("updater.openReleasePage");
  viewVersionsBtn.dataset.testid = "update-ready-view-versions";

  const restartBtn = document.createElement("button");
  restartBtn.type = "button";
  restartBtn.textContent = t("updater.restartNow");
  restartBtn.dataset.testid = "update-ready-restart";

  controls.appendChild(laterBtn);
  controls.appendChild(viewVersionsBtn);
  controls.appendChild(restartBtn);
  toast.appendChild(controls);
  root.appendChild(toast);

  const cleanup = () => {
    toast.remove();
  };

  viewVersionsBtn.addEventListener("click", () => {
    void openExternalUrl(FORMULA_RELEASES_URL);
    cleanup();
  });

  laterBtn.addEventListener("click", () => {
    const version = update.version.trim();
    if (version) {
      const storage = getLocalStorageOrNull();
      if (storage) setUpdaterDismissal(storage, version);
    }
    cleanup();
  });

  restartBtn.addEventListener("click", () => {
    void (async () => {
      // Prevent double-click restart attempts.
      clearDismissalOnUpdateInitiated();
      lastUpdateError = null;
      restartBtn.disabled = true;
      viewVersionsBtn.disabled = true;
      laterBtn.disabled = true;
      try {
        const didRestart = await restartToInstallUpdate();
        if (didRestart) {
          cleanup();
        }
      } finally {
        // If the restart was cancelled (unsaved changes) or failed, keep the toast visible.
        if (toast.isConnected) {
          restartBtn.disabled = false;
          viewVersionsBtn.disabled = false;
          laterBtn.disabled = false;
        }
      }
    })().catch(() => {
      // Best-effort: avoid unhandled rejections from the async restart handler.
    });
  });
}

export async function handleUpdaterEvent(name: UpdaterEventName, payload: UpdaterEventPayload): Promise<void> {
  const rawSource = payload?.source;
  const followUpEligible =
    name === "update-not-available" || name === "update-check-error" || name === "update-available";
  const followUpManual = manualUpdateCheckFollowUp && rawSource === "startup" && followUpEligible;
  const source = followUpManual ? "manual" : rawSource;

  if (name === "update-available" && rawSource === "startup" && !manualUpdateCheckFollowUp) {
    const version = typeof payload?.version === "string" ? payload.version.trim() : "";
    const body = typeof payload?.body === "string" ? payload.body.trim() : "";
    // If the user recently clicked "Later" for this version, suppress repeat system notifications too.
    const suppressed = version ? shouldSuppressStartupUpdatePrompt(version) : false;
    if (!suppressed) {
      const appName = t("app.title");
      const message =
        version && body
          ? tWithVars("updater.systemNotificationBodyWithNotes", { appName, version, notes: body })
          : version
            ? tWithVars("updater.systemNotificationBodyWithVersion", { appName, version })
            : body
              ? tWithVars("updater.systemNotificationBodyWithNotesUnknownVersion", { appName, notes: body })
              : tWithVars("updater.systemNotificationBodyGeneric", { appName });

      void notify({ title: t("updater.updateAvailableTitle"), body: message }).catch(() => {});
    }
  }

  // Tray-triggered manual checks can happen while the app is hidden to tray. Ensure the
  // window is visible before rendering any toast/dialog feedback.
  //
  // If we're treating a startup-sourced completion event as a manual follow-up (because the user
  // clicked "Check for Updates" while a startup check was already running), avoid re-showing /
  // refocusing the window: it was already surfaced for the initial manual request.
  if (source === "manual" && !followUpManual) {
    await showMainWindowBestEffort();
  }

  const shouldSurfaceCompletion =
    manualUpdateCheckFollowUp && (name === "update-not-available" || name === "update-check-error");
  const shouldSurfaceToast = source === "manual" || shouldSurfaceCompletion;

  // `update-downloaded` can happen from a startup background download. Surface lightweight UI so
  // the user can approve a restart/apply step when they're ready.
  if (name === "update-downloaded") {
    const version = typeof payload?.version === "string" && payload.version.trim() !== "" ? payload.version.trim() : "unknown";
    backendDownloadInFlightVersion = null;
    backendDownloadedVersion = version;
    lastUpdateError = null;

    // If the update dialog is already open for this version (e.g. user triggered a manual check
    // while a startup download was in-flight), update it to show the restart CTA immediately.
    if (updateDialog && updateInfo?.version === version) {
      renderUpdateDialog();
    }

    // If the user recently dismissed startup update prompts for this version, keep background
    // download completion quiet; they can still manually check and see the restart CTA.
    if (rawSource === "startup" && version && shouldSuppressStartupUpdatePrompt(version)) {
      return;
    }

    const dialogOpen = updateDialog?.dialog
      ? (updateDialog.dialog as any).open === true || updateDialog.dialog.hasAttribute("open")
      : false;
    if (!dialogOpen) {
      showUpdateReadyToast({ version });
    }
    return;
  }

  switch (name) {
    case "update-check-already-running": {
      if (source === "manual") {
        setManualUpdateCheckFollowUp(true);
        showToast(t("updater.alreadyChecking"), "info");
      }
      break;
    }
    case "update-check-started": {
      if (source === "manual") {
        setManualUpdateCheckFollowUp(false);
        showToast(t("updater.checking"), "info");
      }
      break;
    }
    case "update-not-available": {
      if (shouldSurfaceToast) {
        setManualUpdateCheckFollowUp(false);
        showToast(t("updater.upToDate"), "info");
      }
      break;
    }
    case "update-check-error": {
      if (!shouldSurfaceToast) break;
      setManualUpdateCheckFollowUp(false);
      const message =
        typeof payload?.error === "string" && payload.error.trim() !== ""
          ? payload.error
          : typeof payload?.message === "string" && payload.message.trim() !== ""
            ? payload.message
            : t("updater.unknownError");
      showToast(tWithVars("updater.errorWithMessage", { message }), "error");
      break;
    }
    case "update-available": {
      // Only show the in-app update dialog for manual checks. Startup checks should be silent
      // (the shell shows a system notification instead).
      //
      // Note: if the user triggers "Check for Updates" while a startup check is in-flight, the
      // backend will later emit the completion event with `source: "startup"`. We rewrite that to
      // `source: "manual"` above (see `followUpManual`), so this check still opens the dialog for
      // the user-requested result.
      if (source !== "manual") break;
      setManualUpdateCheckFollowUp(false);

      const version =
        typeof payload?.version === "string" && payload.version.trim() !== ""
          ? payload.version.trim()
          : "unknown";
      const body = typeof payload?.body === "string" && payload.body.trim() !== "" ? payload.body : null;

      // If a new version is available, drop any persisted "Later" suppression from prior versions.
      const storage = getLocalStorageOrNull();
      if (storage) {
        const record = readUpdaterDismissal(storage);
        if (record && record.version !== version) {
          clearUpdaterDismissal(storage);
        }
      }

      openUpdateAvailableDialog({ ...payload, version, body });
      break;
    }
    case "update-download-started": {
      const version = typeof payload?.version === "string" ? payload.version.trim() : "";
      if (!version) break;
      backendDownloadInFlightVersion = version;
      progressDownloaded = 0;
      progressTotal = null;
      progressPercent = null;

      // Only re-render if a dialog is already open for this version; avoid creating UI on startup.
      if (updateDialog && updateInfo?.version === version) {
        renderUpdateDialog();
      }
      break;
    }
    case "update-download-progress": {
      const version = typeof payload?.version === "string" ? payload.version.trim() : "";
      if (!version) break;
      if (backendDownloadInFlightVersion !== version) break;

      const percent = extractNumber(payload?.percent);
      if (typeof percent === "number") {
        progressPercent = clampPercent(percent);
        progressTotal = 100;
        progressDownloaded = progressPercent;
      } else {
        const total = extractNumber(payload?.total);
        const downloaded = extractNumber(payload?.downloaded);
        if (typeof total === "number") progressTotal = total;
        if (typeof downloaded === "number") progressDownloaded = downloaded;
      }

      if (updateDialog && updateInfo?.version === version) {
        renderUpdateDialog();
      }
      break;
    }
    case "update-download-error": {
      const version = typeof payload?.version === "string" ? payload.version.trim() : "";
      if (!version) break;
      if (backendDownloadInFlightVersion === version) {
        backendDownloadInFlightVersion = null;
      }

      // Only surface the error inside an already-open dialog for this version; otherwise keep
      // startup background download failures quiet.
      if (updateDialog && updateInfo?.version === version) {
        lastUpdateError =
          typeof payload?.message === "string" && payload.message.trim() !== "" ? payload.message : t("updater.unknownError");
        renderUpdateDialog();
      }
      break;
    }
    case "update-downloaded": {
      break;
    }
  }
}

export async function installUpdaterUi(listenArg?: TauriListen): Promise<void> {
  const listen = listenArg ?? getTauriListen();
  if (!listen) return;

  const events: UpdaterEventName[] = [
    "update-check-already-running",
    "update-check-started",
    "update-not-available",
    "update-check-error",
    "update-available",
    "update-download-started",
    "update-download-progress",
    "update-downloaded",
    "update-download-error",
  ];

  const installs = events.map(async (eventName) => {
    try {
      await listen(eventName, (event) => {
        const payload = (event as any)?.payload as UpdaterEventPayload;
        void handleUpdaterEvent(eventName, payload).catch((err) => {
          console.error(`[formula][updater-ui] Failed to handle ${eventName}:`, err);
        });
      });
    } catch (err) {
      // Best-effort: if a single event name is blocked by capabilities or missing in a given
      // runtime, keep the rest of the updater UI functional (and still allow the backend startup
      // check to proceed via the `updater-ui-ready` handshake).
      console.error(`[formula][updater-ui] Failed to listen for ${eventName}:`, err);
    }
  });

  await Promise.all(installs);
}

/**
 * Called by the updater UI when the user confirms "Restart now".
 *
 * This routes through the normal quit flow (Workbook_BeforeClose macros, backend-sync drain,
 * and the unsaved-changes confirm prompt) before triggering the updater install step.
 */
export async function restartToInstallUpdate(): Promise<boolean> {
  return await requestAppRestart({
    beforeQuit: async () => {
      try {
        const update = downloadedUpdate;
        if (update) {
          await update.install();
        } else {
          await installUpdateAndRestart();
        }
      } catch (err) {
        lastUpdateError = errorMessage(err);
        throw err;
      }
    },
    beforeQuitErrorToast: t("updater.restartFailed"),
  });
}

/**
 * Internal: reset module-scoped state for unit tests.
 *
 * The updater UI module maintains dialog + in-flight state in module-level variables so the
 * desktop shell can treat it as a singleton. Vitest runs multiple suites in the same worker,
 * so tests must be able to start from a clean state.
 */
export function __resetUpdaterUiStateForTests(): void {
  try {
    updateDialog?.dialog.close();
  } catch {
    // ignore
  }
  try {
    updateDialog?.dialog.remove();
  } catch {
    // ignore
  }

  updateDialog = null;
  updateInfo = null;
  downloadedUpdate = null;
  downloadInFlight = false;
  lastUpdateError = null;

  progressDownloaded = 0;
  progressTotal = null;
  progressPercent = null;

  updateReadyToastShownForVersion = null;
  backendDownloadedVersion = null;
  backendDownloadInFlightVersion = null;
  try {
    document.querySelectorAll?.('[data-testid="update-ready-toast"]').forEach((el) => el.remove());
  } catch {
    // ignore
  }

  manualUpdateCheckFollowUp = false;
  if (manualUpdateCheckFollowUpTimeout) {
    clearTimeout(manualUpdateCheckFollowUpTimeout);
    manualUpdateCheckFollowUpTimeout = null;
  }
}
