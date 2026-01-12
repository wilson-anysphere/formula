import { showToast } from "../extensions/ui.js";

import { requestAppRestart } from "./appQuit";
import { shellOpen } from "./shellOpen";
import { installUpdateAndRestart } from "./updater";

export const FORMULA_RELEASES_URL = "https://github.com/wilson-anysphere/formula/releases";

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
  version: HTMLElement;
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

function getRelaunchOrNull(): (() => Promise<void>) | null {
  const tauri = (globalThis as any).__TAURI__;
  const relaunch =
    (tauri?.process?.relaunch as (() => Promise<void> | void) | undefined) ??
    (tauri?.process?.restart as (() => Promise<void> | void) | undefined) ??
    (tauri?.app?.relaunch as (() => Promise<void> | void) | undefined) ??
    (tauri?.app?.restart as (() => Promise<void> | void) | undefined) ??
    null;
  if (typeof relaunch !== "function") return null;
  return async () => {
    await relaunch.call(null);
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

function ensureUpdateDialog(): DialogElements {
  if (updateDialog) return updateDialog;

  const dialog = document.createElement("dialog");
  dialog.className = "dialog";
  dialog.dataset.testid = "updater-dialog";
  dialog.style.maxWidth = "min(640px, calc(100vw - 32px))";
  dialog.style.width = "520px";

  const title = document.createElement("div");
  title.className = "dialog__title";
  title.textContent = "Update available";

  const version = document.createElement("div");
  version.dataset.testid = "updater-version";
  version.style.fontWeight = "600";

  const body = document.createElement("pre");
  body.dataset.testid = "updater-body";
  body.style.whiteSpace = "pre-wrap";
  body.style.margin = "10px 0 0";
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
  laterBtn.textContent = "Later";
  laterBtn.dataset.testid = "updater-later";

  const viewVersionsBtn = document.createElement("button");
  viewVersionsBtn.type = "button";
  viewVersionsBtn.textContent = "Open release page";
  viewVersionsBtn.dataset.testid = "updater-view-versions";

  const downloadBtn = document.createElement("button");
  downloadBtn.type = "button";
  downloadBtn.textContent = "Download";
  downloadBtn.dataset.testid = "updater-download";

  const restartBtn = document.createElement("button");
  restartBtn.type = "button";
  restartBtn.textContent = "Restart now";
  restartBtn.dataset.testid = "updater-restart";
  restartBtn.style.display = "none";

  controls.appendChild(laterBtn);
  controls.appendChild(viewVersionsBtn);
  controls.appendChild(downloadBtn);
  controls.appendChild(restartBtn);

  dialog.appendChild(title);
  dialog.appendChild(version);
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
    safeClose(dialog, "later");
  });

  laterBtn.addEventListener("click", () => {
    if (downloadInFlight) return;
    safeClose(dialog, "later");
  });

  viewVersionsBtn.addEventListener("click", () => {
    const url = updateInfo?.manualDownloadUrl ?? FORMULA_RELEASES_URL;
    void openExternalUrl(url);
    if (!downloadInFlight) safeClose(dialog, "versions");
  });

  downloadBtn.addEventListener("click", () => {
    void startUpdateDownload();
  });

  restartBtn.addEventListener("click", () => {
    void (async () => {
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
    version,
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

  els.version.textContent = info ? `Version ${info.version}` : "";
  els.body.textContent = info?.body ? info.body : "";

  els.laterBtn.disabled = downloadInFlight;
  els.viewVersionsBtn.disabled = downloadInFlight;
  els.downloadBtn.disabled = downloadInFlight || !!downloadedUpdate;
  els.downloadBtn.textContent = downloadInFlight ? "Downloading…" : "Download";

  const readyToInstall = !!downloadedUpdate && !downloadInFlight;
  els.restartBtn.style.display = readyToInstall ? "" : "none";
  els.restartBtn.disabled = false;

  if (downloadInFlight) {
    els.status.textContent = "Downloading update…";
    els.progressWrap.style.display = "";
    renderProgress();
  } else if (lastUpdateError) {
    els.progressWrap.style.display = "none";
    els.status.textContent =
      downloadedUpdate
        ? `Update failed to install automatically.\n\nDownload manually from the release page, or try restarting again.\n\n${lastUpdateError}`
        : `Update failed to download.\n\nDownload manually from the release page.\n\n${lastUpdateError}`;
  } else if (downloadedUpdate) {
    els.status.textContent = "Update ready to install. Restart now?";
    els.progressWrap.style.display = "none";
  } else {
    els.status.textContent = "";
    els.progressWrap.style.display = "none";
  }

  if (lastUpdateError) {
    els.viewVersionsBtn.textContent = "Download manually";
    els.viewVersionsBtn.style.background = "var(--accent)";
    els.viewVersionsBtn.style.borderColor = "var(--accent-border)";
    els.viewVersionsBtn.style.color = "var(--text-on-accent)";
    els.viewVersionsBtn.style.fontWeight = "700";
  } else {
    els.viewVersionsBtn.textContent = "Open release page";
    els.viewVersionsBtn.style.background = "";
    els.viewVersionsBtn.style.borderColor = "";
    els.viewVersionsBtn.style.color = "";
    els.viewVersionsBtn.style.fontWeight = "";
  }

  if (readyToInstall && !getRelaunchOrNull()) {
    els.restartBtn.textContent = "Quit now";
  } else {
    els.restartBtn.textContent = "Restart now";
  }
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
  els.progressText.textContent = "Downloading…";
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

  const updater = getUpdaterApiOrNull();
  if (!updater) {
    console.warn("Updater API not available; cannot download update.");
    lastUpdateError = "Auto-updater is unavailable in this build.";
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
      throw new Error("Update no longer available");
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
    lastUpdateError = String(err);
    showToast(lastUpdateError, "error");
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

  // Tray-triggered manual checks can happen while the app is hidden to tray. Ensure the
  // window is visible before rendering any toast/dialog feedback.
  if (source === "manual") {
    await showMainWindowBestEffort();
  }

  switch (name) {
    case "update-check-already-running": {
      if (source === "manual") showToast("Already checking for updates…", "info");
      break;
    }
    case "update-check-started": {
      if (source === "manual") showToast("Checking for updates…", "info");
      break;
    }
    case "update-not-available": {
      if (source === "manual") showToast("You're up to date.", "info");
      break;
    }
    case "update-check-error": {
      if (source !== "manual") break;
      const message =
        typeof payload?.error === "string" && payload.error.trim() !== ""
          ? payload.error
          : typeof payload?.message === "string" && payload.message.trim() !== ""
            ? payload.message
            : "Unknown error";
      showToast(message, "error");
      break;
    }
    case "update-available": {
      const version = typeof payload?.version === "string" && payload.version.trim() !== "" ? payload.version.trim() : "unknown";
      const body = typeof payload?.body === "string" && payload.body.trim() !== "" ? payload.body : null;

      // Avoid repeatedly showing the startup prompt for the same version.
      if (source !== "manual" && updateDialogShownForVersion === version) break;
      updateDialogShownForVersion = version;

      openUpdateAvailableDialog({ ...payload, version, body });
      break;
    }
  }
}

export function installUpdaterUi(listenArg?: TauriListen): void {
  const listen = listenArg ?? getTauriListen();
  if (!listen) return;

  const events: UpdaterEventName[] = [
    "update-check-already-running",
    "update-check-started",
    "update-not-available",
    "update-check-error",
    "update-available",
  ];

  for (const eventName of events) {
    void listen(eventName, (event) => {
      const payload = (event as any)?.payload as UpdaterEventPayload;
      void handleUpdaterEvent(eventName, payload);
    });
  }
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
        lastUpdateError = String(err);
        throw err;
      }

      const relaunch = getRelaunchOrNull();
      if (relaunch) {
        try {
          await relaunch();
        } catch {
          // Fall back to the `quit_app` hard-exit in `requestAppRestart`.
        }
      }
    },
    beforeQuitErrorToast: "Failed to restart to install the update.",
  });
}
