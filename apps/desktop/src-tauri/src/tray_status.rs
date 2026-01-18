use std::sync::Mutex;

use tauri::tray::TrayIcon;

use crate::ipc_limits::{LimitedString, MAX_IPC_TRAY_STATUS_BYTES};
use crate::ipc_origin;

#[derive(Default)]
pub struct TrayStatusState {
    tray: Mutex<Option<TrayIcon>>,
    status: Mutex<Option<String>>,
}

impl TrayStatusState {
    pub fn set_tray(&self, tray: TrayIcon) {
        {
            let mut guard = self
                .tray
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner());
            *guard = Some(tray);
        }

        // If the frontend set a status before the tray icon finished initializing,
        // apply it now.
        let status = self
            .status
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .clone();
        if let Some(status) = status {
            let _ = self.apply_status(&status);
        }
    }

    fn normalize_status(status: &str) -> Result<&str, String> {
        match status.trim() {
            "idle" | "syncing" | "error" => Ok(status.trim()),
            other => Err(format!(
                "invalid tray status '{other}' (expected 'idle', 'syncing', or 'error')"
            )),
        }
    }

    fn apply_status(&self, status: &str) -> Result<(), String> {
        let status = Self::normalize_status(status)?;

        let tooltip = match status {
            "idle" => "Formula — Idle",
            "syncing" => "Formula — Syncing…",
            "error" => "Formula — Error",
            other => return Err(format!("invalid tray status '{other}'")),
        };

        let icon = match status {
            "idle" => tauri::include_image!("icons/tray.png"),
            "syncing" => tauri::include_image!("icons/tray-syncing.png"),
            "error" => tauri::include_image!("icons/tray-error.png"),
            other => return Err(format!("invalid tray status '{other}'")),
        };

        let guard = self
            .tray
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        let Some(tray) = guard.as_ref() else {
            // Not ready yet; the status is still recorded and will be applied when the tray icon
            // is registered.
            return Ok(());
        };

        // Prefer updating both the tooltip and title (on macOS, title can be shown next to the
        // status bar icon). Ignore title failures so we still update tooltip + icon everywhere.
        tray.set_tooltip(Some(tooltip.to_string()))
            .map_err(|e| e.to_string())?;
        let _ = tray.set_title(Some(tooltip.to_string()));

        tray.set_icon(Some(icon)).map_err(|e| e.to_string())?;

        Ok(())
    }

    pub fn update_status(&self, status: &str) -> Result<(), String> {
        let status = Self::normalize_status(status)?;

        let should_apply = {
            let mut guard = self
                .status
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner());
            if guard.as_deref() == Some(status) {
                false
            } else {
                *guard = Some(status.to_string());
                true
            }
        };

        if should_apply {
            self.apply_status(status)?;
        }

        Ok(())
    }
}

/// Update the system tray icon + tooltip to match a simple background status.
///
/// Supported statuses:
/// - `idle`
/// - `syncing`
/// - `error`
#[tauri::command]
pub fn set_tray_status(
    window: tauri::WebviewWindow,
    state: tauri::State<'_, TrayStatusState>,
    status: LimitedString<MAX_IPC_TRAY_STATUS_BYTES>,
) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    ipc_origin::ensure_main_window(window.label(), "tray status", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_trusted_origin(&url, "tray status", ipc_origin::Verb::Is)?;
    ipc_origin::ensure_stable_origin(&window, "tray status", ipc_origin::Verb::Is)?;
    state.inner().update_status(status.as_ref())
}
