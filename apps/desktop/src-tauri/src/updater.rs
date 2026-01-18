use serde::{Deserialize, Serialize};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::OnceLock;
use tauri::{AppHandle, Emitter, Manager};
use tauri_plugin_updater::UpdaterExt;
use tokio::sync::{Mutex, Notify};
use tokio::time::{timeout, Duration, Instant};

use crate::ipc_origin;
use crate::updater_download_cache::UpdaterDownloadCache;

static UPDATE_CHECK_IN_FLIGHT: AtomicBool = AtomicBool::new(false);
static UPDATE_DOWNLOAD_STATE: OnceLock<Mutex<UpdateDownloadState>> = OnceLock::new();

struct DownloadedUpdate {
    version: String,
    update: tauri_plugin_updater::Update,
    download_path: std::path::PathBuf,
}

struct UpdateDownloadState {
    in_flight: bool,
    downloading_version: Option<String>,
    downloaded: Option<DownloadedUpdate>,
    last_error: Option<String>,
    notify: std::sync::Arc<Notify>,
}

fn update_download_state() -> &'static Mutex<UpdateDownloadState> {
    UPDATE_DOWNLOAD_STATE.get_or_init(|| {
        Mutex::new(UpdateDownloadState {
            in_flight: false,
            downloading_version: None,
            downloaded: None,
            last_error: None,
            notify: std::sync::Arc::new(Notify::new()),
        })
    })
}

fn updater_download_cache_base_dir(app: &AppHandle) -> std::path::PathBuf {
    match app.path().app_cache_dir() {
        Ok(dir) => dir,
        Err(err) => {
            // Best-effort fallback: `app_cache_dir` should be available in normal desktop builds.
            crate::stdio::stderrln(format_args!(
                "failed to resolve app cache dir for updater downloads: {err}"
            ));
            std::env::temp_dir().join("formula").join("updater-cache")
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum UpdateCheckSource {
    Startup,
    Manual,
}

fn env_flag_truthy(name: &str) -> bool {
    match std::env::var(name) {
        Ok(raw) => {
            let v = raw.trim().to_ascii_lowercase();
            !(v.is_empty() || v == "0" || v == "false")
        }
        Err(_) => false,
    }
}

pub fn spawn_update_check(app: &AppHandle, source: UpdateCheckSource) {
    // Benchmarks (startup + idle memory) need to be stable/independent of updater/network behavior.
    // The primary gate lives in `main.rs` (skipping the startup listener entirely), but keep this
    // as defense-in-depth so any future `UpdateCheckSource::Startup` callsites respect it too.
    if matches!(source, UpdateCheckSource::Startup)
        && env_flag_truthy("FORMULA_DISABLE_STARTUP_UPDATE_CHECK")
    {
        return;
    }

    if UPDATE_CHECK_IN_FLIGHT.swap(true, Ordering::SeqCst) {
        // Startup checks should not interrupt the user; manual checks should provide feedback so
        // the frontend can display "Already checking..." without triggering a second network call.
        if matches!(source, UpdateCheckSource::Manual) {
            let _ = app.emit(
                "update-check-already-running",
                serde_json::json!({ "source": source }),
            );
        }
        return;
    }

    let handle = app.clone();
    let _ = handle.emit("update-check-started", serde_json::json!({ "source": source }));
    tauri::async_runtime::spawn(async move {
        struct UpdateCheckInFlightGuard;
        impl Drop for UpdateCheckInFlightGuard {
            fn drop(&mut self) {
                UPDATE_CHECK_IN_FLIGHT.store(false, Ordering::SeqCst);
            }
        }
        let _guard = UpdateCheckInFlightGuard;

        match handle.updater() {
            Ok(updater) => match updater.check().await {
                Ok(Some(update)) => {
                    let version = update.version.clone();
                    let body = update.body.clone();
                    let payload = serde_json::json!({
                        "source": source,
                        "version": version,
                        "body": body,
                    });
                    let _ = handle.emit("update-available", payload);

                    // Start a best-effort background download so the user can apply the update
                    // immediately once they approve a restart.
                    if matches!(source, UpdateCheckSource::Startup) {
                        spawn_update_download(&handle, source, update).await;
                    }
                }
                Ok(None) => {
                    let _ = handle
                        .emit("update-not-available", serde_json::json!({ "source": source }));
                }
                Err(err) => {
                    let msg = err.to_string();
                    let _ = handle.emit(
                        "update-check-error",
                        serde_json::json!({ "source": source, "message": msg }),
                    );
                    crate::stdio::stderrln(format_args!("updater check failed: {err}"));
                }
            },
            Err(err) => {
                let msg = err.to_string();
                let _ = handle.emit(
                    "update-check-error",
                    serde_json::json!({ "source": source, "message": msg }),
                );
                crate::stdio::stderrln(format_args!("updater check failed: {err}"));
            }
        }
    });
}

async fn spawn_update_download(
    app: &AppHandle,
    source: UpdateCheckSource,
    update: tauri_plugin_updater::Update,
) {
    let version = update.version.clone();

    {
        let mut state = update_download_state().lock().await;

        if state
            .downloaded
            .as_ref()
            .is_some_and(|downloaded| downloaded.version == version)
        {
            return;
        }

        if state.in_flight {
            // Another download is already in flight (avoid double-downloads). If it's for this
            // same version, we also skip.
            if state.downloading_version.as_deref() == Some(&version) {
                return;
            }
            return;
        }

        state.in_flight = true;
        state.downloading_version = Some(version.clone());
        state.last_error = None;
    }

    let handle = app.clone();
    let version_for_events = version.clone();
    let _ = handle.emit(
        "update-download-started",
        serde_json::json!({ "source": source, "version": version_for_events }),
    );

    tauri::async_runtime::spawn(async move {
        let version_for_progress = version.clone();
        let mut downloaded_bytes: u64 = 0;
        let download_result = update
            .download(
                |chunk_length, content_length| {
                    downloaded_bytes = downloaded_bytes.saturating_add(chunk_length as u64);
                    let percent =
                        content_length.and_then(|total| (total > 0).then(|| (downloaded_bytes as f64 / total as f64) * 100.0));
                    let _ = handle.emit(
                        "update-download-progress",
                        serde_json::json!({
                            "source": source,
                            "version": version_for_progress.as_str(),
                            "chunkLength": chunk_length,
                            "downloaded": downloaded_bytes,
                            "total": content_length,
                            "percent": percent,
                        }),
                    );
                },
                || {},
            )
            .await;

        match download_result {
            Ok(bytes) => {
                let cache = UpdaterDownloadCache::new(updater_download_cache_base_dir(&handle));
                let download_path = match cache.write(&version, &bytes) {
                    Ok(path) => path,
                    Err(err) => {
                        // Treat cache write failures as a download failure: without a persisted
                        // payload we can't install later without keeping bytes resident in memory.
                        let msg = err.to_string();
                        let (notify, stale) = {
                            let mut state = update_download_state().lock().await;
                            state.in_flight = false;
                            if state.downloading_version.as_deref() == Some(version.as_str()) {
                                state.downloading_version = None;
                            }
                            let stale = state.downloaded.take();
                            state.downloaded = None;
                            state.last_error = Some(msg.clone());
                            (state.notify.clone(), stale)
                        };
                        // If we had a previously cached file, it is now stale (the state no longer
                        // references it). Best-effort cleanup.
                        if let Some(stale) = stale {
                            let _ = UpdaterDownloadCache::delete(&stale.download_path);
                        }
                        notify.notify_waiters();

                        let _ = handle.emit(
                            "update-download-error",
                            serde_json::json!({ "source": source, "version": version, "message": msg }),
                        );
                        crate::stdio::stderrln(format_args!(
                            "updater download cache write failed: {err}"
                        ));
                        return;
                    }
                };

                // Release the downloaded bytes from memory ASAP. We'll reload from disk when the
                // user approves install.
                drop(bytes);

                let (notify, stale) = {
                    let mut state = update_download_state().lock().await;
                    state.in_flight = false;
                    if state.downloading_version.as_deref() == Some(version.as_str()) {
                        state.downloading_version = None;
                    }

                    let stale = state.downloaded.take();
                    state.downloaded = Some(DownloadedUpdate {
                        version: version.clone(),
                        update,
                        download_path: download_path.clone(),
                    });
                    state.last_error = None;
                    (state.notify.clone(), stale)
                };
                notify.notify_waiters();

                // Best-effort cleanup for the previously downloaded payload, if any.
                if let Some(stale) = stale {
                    if stale.download_path != download_path {
                        let _ = UpdaterDownloadCache::delete(&stale.download_path);
                    }
                }

                let _ = handle.emit(
                    "update-downloaded",
                    serde_json::json!({ "source": source, "version": version }),
                );
            }
            Err(err) => {
                let msg = err.to_string();
                let (notify, stale) = {
                    let mut state = update_download_state().lock().await;
                    state.in_flight = false;
                    if state.downloading_version.as_deref() == Some(version.as_str()) {
                        state.downloading_version = None;
                    }
                    let stale = state.downloaded.take();
                    state.downloaded = None;
                    state.last_error = Some(msg.clone());
                    (state.notify.clone(), stale)
                };
                notify.notify_waiters();

                if let Some(stale) = stale {
                    let _ = UpdaterDownloadCache::delete(&stale.download_path);
                }

                let _ = handle.emit(
                    "update-download-error",
                    serde_json::json!({ "source": source, "version": version, "message": msg }),
                );
                crate::stdio::stderrln(format_args!("updater download failed: {err}"));
            }
        }
    });
}

/// Installs the currently downloaded update (if any).
///
/// Intended to be called after the user approves a restart via `restartToInstallUpdate()`.
#[tauri::command]
pub async fn install_downloaded_update(window: tauri::WebviewWindow) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "update installation", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "update installation", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "update installation", ipc_origin::Verb::Is)?;

    // Avoid hanging the IPC handler indefinitely if the background download never resolves.
    const WAIT_TIMEOUT: Duration = Duration::from_secs(60);
    let start = Instant::now();

    loop {
        // Create the wait handle *before* checking state so we can't miss a `notify_waiters()`
        // that happens between observing `in_flight` and calling `.notified().await`.
        let notify = {
            let state = update_download_state().lock().await;
            state.notify.clone()
        };
        let notified = notify.notified();

        // Try to grab the downloaded update bytes without holding the mutex while we run the
        // (potentially slow / IO-heavy) install step.
        let downloaded = {
            let mut state = update_download_state().lock().await;

            if let Some(downloaded) = state.downloaded.take() {
                Some(downloaded)
            } else if !state.in_flight {
                if let Some(err) = state.last_error.clone() {
                    return Err(err);
                }
                return Err("No downloaded update is available".to_string());
            } else {
                None
            }
        };

        if let Some(downloaded) = downloaded {
            let bytes = UpdaterDownloadCache::read(&downloaded.download_path)
                .map_err(|err| err.to_string())?;

            let result = downloaded
                .update
                .install(&bytes)
                .map_err(|err| err.to_string());

            // Ensure we don't keep the update payload resident after the install attempt.
            drop(bytes);

            if result.is_err() {
                // Restore the downloaded update so the user can retry without forcing a re-download.
                let mut state = update_download_state().lock().await;
                if state.downloaded.is_none() {
                    state.downloaded = Some(downloaded);
                } else {
                    // Another update finished downloading while we attempted to install; our cached
                    // file is now stale and unreachable. Clean it up best-effort.
                    let _ = UpdaterDownloadCache::delete(&downloaded.download_path);
                }
            } else {
                // Install succeeded: remove the cached payload.
                let _ = UpdaterDownloadCache::delete(&downloaded.download_path);
            }

            return result;
        }

        let remaining = WAIT_TIMEOUT
            .checked_sub(start.elapsed())
            .unwrap_or(Duration::ZERO);
        if remaining.is_zero() {
            return Err("Timed out waiting for a downloaded update".to_string());
        }
        timeout(remaining, notified)
            .await
            .map_err(|_| "Timed out waiting for a downloaded update".to_string())?;
    }
}
