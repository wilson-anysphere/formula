import { showToast } from "../extensions/ui.js";
import { t, tWithVars } from "../i18n/index.js";

import { requestAppRestart } from "./appQuit";
import { notify } from "./notifications";
import { shellOpen } from "./shellOpen";
import { installUpdateAndRestart } from "./updater";

export const FORMULA_RELEASES_URL = "https://github.com/wilson-anysphere/formula/releases";

export const UPDATER_DISMISSED_VERSION_KEY = "formula.updater.dismissedVersion";
export const UPDATER_DISMISSED_AT_KEY = "formula.updater.dismissedAt";
const UPDATER_DISMISSAL_TTL_MS = 7 * 24 * 60 * 60 * 1000;

type UpdaterEventName =
  | "update-check-already-running"
  | "update-check-started"
  | "update-not-available"
  | "update-check-error"
  | "update-available";

type UpdaterEventPayload = {
  source?: string;
  version?: string;
  body?: string | null;
  message?: string;
  error?: string;
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

type TauriListen = (event: string, handler: (event: any) => void) => Promise<() => void>;

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

let updateDialogShownForVersion: string | null = null;

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
  const listen = (globalThis as any).__TAURI__?.event?.listen as TauriListen | undefined;
  if (typeof listen !== "function") return null;
  return listen;
}

function getTauriWindowHandle(): any | null {
  const winApi = (globalThis as any).__TAURI__?.window;
  if (!winApi) return null;

  // Mirrors the flexible handle lookup used in `main.ts`. We intentionally avoid
  // a hard dependency on `@tauri-apps/api`.
  const handle =
    (typeof winApi.getCurrentWebviewWindow === "function" ? winApi.getCurrentWebviewWindow() : null) ??
    (typeof winApi.getCurrentWindow === "function" ? winApi.getCurrentWindow() : null) ??
    (typeof winApi.getCurrent === "function" ? winApi.getCurrent() : null) ??
    winApi.appWindow ??
    null;

  return handle ?? null;
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
  const tauri = (globalThis as any).__TAURI__;
  const updater = tauri?.updater ?? tauri?.plugin?.updater ?? tauri?.plugins?.updater ?? null;
  if (!updater) return null;

  const check = updater?.check as (() => Promise<unknown>) | undefined;
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
      document.body.appendChild(updateDialog.dialog);
    }
    return updateDialog;
  }

  const dialog = document.createElement("dialog");
  dialog.className = "dialog";
  dialog.dataset.testid = "updater-dialog";
  dialog.style.maxWidth = "min(640px, calc(100vw - 32px))";
  dialog.style.width = "520px";

  const title = document.createElement("div");
  title.className = "dialog__title";
  title.textContent = t("updater.updateAvailableTitle");

  const version = document.createElement("div");
  version.dataset.testid = "updater-version";
  version.style.fontWeight = "600";

  const releaseNotesTitle = document.createElement("div");
  releaseNotesTitle.dataset.testid = "updater-release-notes-title";
  releaseNotesTitle.style.marginTop = "10px";
  releaseNotesTitle.style.fontWeight = "600";
  releaseNotesTitle.textContent = t("updater.releaseNotes");

  const body = document.createElement("pre");
  body.dataset.testid = "updater-body";
  body.style.whiteSpace = "pre-wrap";
  body.style.margin = "6px 0 0";
  body.style.maxHeight = "240px";
  body.style.overflow = "auto";

  const status = document.createElement("div");
  status.dataset.testid = "updater-status";
  status.style.marginTop = "10px";
  status.style.whiteSpace = "pre-wrap";

  const progressWrap = document.createElement("div");
  progressWrap.dataset.testid = "updater-progress-wrap";
  progressWrap.style.display = "none";
  progressWrap.style.marginTop = "10px";

  const progressBar = document.createElement("progress");
  progressBar.dataset.testid = "updater-progress";
  progressBar.max = 100;
  progressBar.value = 0;
  progressBar.style.width = "100%";

  const progressText = document.createElement("div");
  progressText.dataset.testid = "updater-progress-text";
  progressText.style.marginTop = "6px";
  progressText.style.fontSize = "12px";
  progressText.style.color = "var(--text-secondary)";

  progressWrap.appendChild(progressBar);
  progressWrap.appendChild(progressText);

  const controls = document.createElement("div");
  controls.style.display = "flex";
  controls.style.justifyContent = "flex-end";
  controls.style.gap = "8px";
  controls.style.marginTop = "12px";

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
  restartBtn.style.display = "none";

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
    })();
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

  els.title.textContent = t("updater.updateAvailableTitle");
  els.version.textContent = info ? tWithVars("updater.updateAvailableMessage", { version: info.version }) : "";
  els.body.textContent = info?.body ? info.body : "";
  els.releaseNotesTitle.textContent = t("updater.releaseNotes");
  els.releaseNotesTitle.style.display = info?.body ? "" : "none";

  els.laterBtn.disabled = downloadInFlight;
  els.viewVersionsBtn.disabled = downloadInFlight;
  els.downloadBtn.disabled = downloadInFlight || !!downloadedUpdate;
  els.laterBtn.textContent = t("updater.later");
  els.viewVersionsBtn.textContent = t("updater.openReleasePage");
  els.downloadBtn.textContent = downloadInFlight ? t("updater.downloading") : t("updater.download");

  const readyToInstall = !!downloadedUpdate && !downloadInFlight;
  els.restartBtn.style.display = readyToInstall ? "" : "none";
  els.restartBtn.disabled = false;

  if (downloadInFlight) {
    els.status.textContent = t("updater.downloadInProgress");
    els.progressWrap.style.display = "";
    renderProgress();
  } else if (lastUpdateError) {
    els.progressWrap.style.display = "none";
    els.status.textContent = downloadedUpdate
      ? tWithVars("updater.installFailedWithMessage", { message: lastUpdateError })
      : tWithVars("updater.downloadFailedWithMessage", { message: lastUpdateError });
  } else if (downloadedUpdate) {
    els.status.textContent = `${t("updater.downloadComplete")} ${t("updater.restartToInstall")}`;
    els.progressWrap.style.display = "none";
  } else {
    els.status.textContent = "";
    els.progressWrap.style.display = "none";
  }

  if (lastUpdateError) {
    els.viewVersionsBtn.textContent = t("updater.downloadManually");
    els.viewVersionsBtn.style.background = "var(--accent)";
    els.viewVersionsBtn.style.borderColor = "var(--accent-border)";
    els.viewVersionsBtn.style.color = "var(--text-on-accent)";
    els.viewVersionsBtn.style.fontWeight = "700";
  } else {
    els.viewVersionsBtn.textContent = t("updater.openReleasePage");
    els.viewVersionsBtn.style.background = "";
    els.viewVersionsBtn.style.borderColor = "";
    els.viewVersionsBtn.style.color = "";
    els.viewVersionsBtn.style.fontWeight = "";
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
    null;
  const downloaded =
    extractNumber(progress?.downloaded) ??
    extractNumber(progress?.current) ??
    extractNumber((progress as any)?.downloaded_bytes) ??
    extractNumber((progress as any)?.downloadedBytes) ??
    null;
  const chunk = extractNumber((progress as any)?.chunkLength) ?? extractNumber((progress as any)?.chunk_length) ?? null;

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
  }

  const body = payload?.body == null ? null : String(payload.body);
  updateInfo = { version, body, manualDownloadUrl: resolveUpdateReleaseUrl(payload) };
  lastUpdateError = null;
  ensureUpdateDialog();
  renderUpdateDialog();
  safeShowModal(updateDialog!.dialog);
}

export async function handleUpdaterEvent(name: UpdaterEventName, payload: UpdaterEventPayload): Promise<void> {
  const source = payload?.source;

  if (name === "update-available" && source === "startup" && !manualUpdateCheckFollowUp) {
    const version = typeof payload?.version === "string" ? payload.version.trim() : "";
    const body = typeof payload?.body === "string" ? payload.body.trim() : "";
    const appName = t("app.title");
    const message =
      version && body
        ? tWithVars("updater.systemNotificationBodyWithNotes", { appName, version, notes: body })
        : version
          ? tWithVars("updater.systemNotificationBodyWithVersion", { appName, version })
          : body
            ? tWithVars("updater.systemNotificationBodyWithNotesUnknownVersion", { appName, notes: body })
            : tWithVars("updater.systemNotificationBodyGeneric", { appName });

    void notify({ title: t("updater.updateAvailableTitle"), body: message });
  }

  // Tray-triggered manual checks can happen while the app is hidden to tray. Ensure the
  // window is visible before rendering any toast/dialog feedback.
  if (source === "manual") {
    await showMainWindowBestEffort();
  }

  const shouldSurfaceCompletion =
    manualUpdateCheckFollowUp && (name === "update-not-available" || name === "update-check-error");
  const shouldSurfaceToast = source === "manual" || shouldSurfaceCompletion;

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
      const manualFollowUp = manualUpdateCheckFollowUp;
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

      // Suppress repeat startup prompts for the same version if the user recently
      // clicked "Later" (persisted across launches). Manual checks always surface the update.
      //
      // Note: a manual check can occur while a background startup check is in-flight. In that
      // case the backend emits `update-check-already-running` (source: manual) and then later
      // delivers the update result from the startup check (source: startup). Treat that as a
      // manual request (no suppression).
      if (source === "startup" && !manualFollowUp && shouldSuppressStartupUpdatePrompt(version)) {
        break;
      }

      // Avoid repeatedly showing the startup prompt for the same version.
      if (source !== "manual" && updateDialogShownForVersion === version) break;
      updateDialogShownForVersion = version;

      openUpdateAvailableDialog({ ...payload, version, body });
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
  ];

  const installs = events.map((eventName) => {
    try {
      return listen(eventName, (event) => {
        const payload = (event as any)?.payload as UpdaterEventPayload;
        void handleUpdaterEvent(eventName, payload);
      }).catch((err) => {
        console.error(`[formula][updater-ui] Failed to listen for ${eventName}:`, err);
        return () => {};
      });
    } catch (err) {
      console.error(`[formula][updater-ui] Failed to listen for ${eventName}:`, err);
      return Promise.resolve(() => {});
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
