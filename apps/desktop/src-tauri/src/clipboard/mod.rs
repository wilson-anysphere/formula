use serde::{Deserialize, Serialize};

pub mod platform;

#[cfg(target_os = "linux")]
mod linux;

#[cfg(target_os = "macos")]
mod macos;

/// Clipboard contents read from the OS.
///
/// This intentionally carries multiple representations of the same copied content
/// (e.g. plain text + HTML + RTF + PNG) so downstream paste handlers can pick the
/// richest available format.
#[derive(Clone, Debug, Default, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ClipboardContent {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtf: Option<String>,
    /// PNG bytes encoded as base64.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub png_base64: Option<String>,
}

/// Payload written to the OS clipboard.
#[derive(Clone, Debug, Default, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct ClipboardWritePayload {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtf: Option<String>,
    /// PNG bytes encoded as base64.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub png_base64: Option<String>,
}

impl ClipboardWritePayload {
    pub fn validate(&self) -> Result<(), ClipboardError> {
        if self.text.is_none() && self.html.is_none() && self.rtf.is_none() && self.png_base64.is_none() {
            return Err(ClipboardError::InvalidPayload(
                "must include at least one of text, html, rtf, pngBase64".to_string(),
            ));
        }
        Ok(())
    }
}

#[derive(Debug, thiserror::Error)]
pub enum ClipboardError {
    #[error("clipboard is not supported on this platform")]
    UnsupportedPlatform,
    #[error("clipboard backend is unavailable: {0}")]
    Unavailable(String),
    #[error("clipboard payload is invalid: {0}")]
    InvalidPayload(String),
    #[error("clipboard operation failed: {0}")]
    OperationFailed(String),
}

pub fn read() -> Result<ClipboardContent, ClipboardError> {
    platform::read()
}

pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    payload.validate()?;
    platform::write(payload)
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn clipboard_read(app: tauri::AppHandle) -> Result<ClipboardContent, String> {
    // Clipboard APIs on macOS call into AppKit, which is not thread-safe.
    // Dispatch to the main thread before touching NSPasteboard.
    #[cfg(target_os = "macos")]
    {
        use tauri::Manager as _;
        return app
            .run_on_main_thread(read)
            .map_err(|e| e.to_string())?
            .map_err(|e| e.to_string());
    }

    #[cfg(not(target_os = "macos"))]
    {
        let _ = app;
        read().map_err(|e| e.to_string())
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn clipboard_write(app: tauri::AppHandle, payload: ClipboardWritePayload) -> Result<(), String> {
    // See `clipboard_read` for why we dispatch to the main thread on macOS.
    #[cfg(target_os = "macos")]
    {
        use tauri::Manager as _;
        return app
            .run_on_main_thread(move || write(&payload))
            .map_err(|e| e.to_string())?
            .map_err(|e| e.to_string());
    }

    #[cfg(not(target_os = "macos"))]
    {
        let _ = app;
        write(&payload).map_err(|e| e.to_string())
    }
}
