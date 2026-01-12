use serde::{Deserialize, Serialize};

pub mod platform;

mod cf_html;

// Clipboard items can contain extremely large rich payloads (especially images).
// Guard against unbounded memory usage / IPC payload sizes by skipping oversized formats.
//
// These match the frontend clipboard provider limits.
const MAX_IMAGE_BYTES: usize = 10 * 1024 * 1024; // 10 MiB
const MAX_RICH_TEXT_BYTES: usize = 2 * 1024 * 1024; // 2 MiB (HTML / RTF)

fn estimate_base64_decoded_len(base64: &str) -> Option<usize> {
    let s = base64.trim();
    if s.is_empty() {
        return Some(0);
    }

    let len = s.len();
    let padding = if s.ends_with("==") {
        2
    } else if s.ends_with('=') {
        1
    } else {
        0
    };

    // If the encoded length is well-formed, we can compute the exact decoded length. Otherwise,
    // compute a conservative upper bound to avoid allocating huge buffers during base64 decode.
    if len % 4 == 0 {
        let groups = len / 4;
        groups.checked_mul(3)?.checked_sub(padding)
    } else {
        let groups = (len + 3) / 4;
        groups.checked_mul(3)
    }
}

#[cfg(any(target_os = "windows", test))]
mod windows_dib;
#[cfg(target_os = "windows")]
mod windows;

// Keep the GTK-backed Linux clipboard implementation behind the `desktop` feature for production
// builds, but still compile the module in unit tests so we can validate pure helper logic (e.g.
// target selection) without requiring a system GTK/WebKit toolchain.
#[cfg(all(target_os = "linux", any(feature = "desktop", test)))]
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

        if let Some(html) = self.html.as_deref() {
            if html.as_bytes().len() > MAX_RICH_TEXT_BYTES {
                return Err(ClipboardError::InvalidPayload(format!(
                    "html exceeds maximum size ({MAX_RICH_TEXT_BYTES} bytes)"
                )));
            }
        }

        if let Some(rtf) = self.rtf.as_deref() {
            if rtf.as_bytes().len() > MAX_RICH_TEXT_BYTES {
                return Err(ClipboardError::InvalidPayload(format!(
                    "rtf exceeds maximum size ({MAX_RICH_TEXT_BYTES} bytes)"
                )));
            }
        }

        if let Some(png_base64) = self.png_base64.as_deref() {
            if !png_base64.trim().is_empty() {
                let decoded_len = estimate_base64_decoded_len(png_base64).unwrap_or(usize::MAX);
                if decoded_len > MAX_IMAGE_BYTES {
                    return Err(ClipboardError::InvalidPayload(format!(
                        "pngBase64 exceeds maximum size ({MAX_IMAGE_BYTES} bytes)"
                    )));
                }
            }
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

fn sanitize_clipboard_content_with_limits(
    mut content: ClipboardContent,
    max_rich_text_bytes: usize,
    max_image_bytes: usize,
) -> ClipboardContent {
    if matches!(content.html, Some(ref s) if s.as_bytes().len() > max_rich_text_bytes) {
        content.html = None;
    }
    if matches!(content.rtf, Some(ref s) if s.as_bytes().len() > max_rich_text_bytes) {
        content.rtf = None;
    }
    if let Some(ref s) = content.png_base64 {
        let decoded_len = estimate_base64_decoded_len(s).unwrap_or(usize::MAX);
        if decoded_len > max_image_bytes {
            content.png_base64 = None;
        }
    }

    content
}

pub fn read() -> Result<ClipboardContent, ClipboardError> {
    let content = platform::read()?;
    Ok(sanitize_clipboard_content_with_limits(
        content,
        MAX_RICH_TEXT_BYTES,
        MAX_IMAGE_BYTES,
    ))
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

#[cfg(test)]
mod tests {
    use super::{
        estimate_base64_decoded_len, sanitize_clipboard_content_with_limits, ClipboardContent,
        ClipboardError, ClipboardWritePayload, MAX_RICH_TEXT_BYTES,
    };

    #[test]
    fn estimate_base64_decoded_len_handles_padding() {
        assert_eq!(estimate_base64_decoded_len(""), Some(0));
        assert_eq!(estimate_base64_decoded_len("Zg=="), Some(1)); // "f"
        assert_eq!(estimate_base64_decoded_len("Zm8="), Some(2)); // "fo"
        assert_eq!(estimate_base64_decoded_len("Zm9v"), Some(3)); // "foo"
        assert_eq!(estimate_base64_decoded_len("AAAA"), Some(3));
    }

    #[test]
    fn estimate_base64_decoded_len_is_conservative_for_malformed_lengths() {
        // Not a multiple of 4: use an upper bound (and let base64 decode validation reject it later).
        assert_eq!(estimate_base64_decoded_len("Zg"), Some(3));
    }

    #[test]
    fn sanitize_clipboard_content_drops_oversized_fields() {
        let content = ClipboardContent {
            text: Some("hello".to_string()),
            html: Some("123456".to_string()),      // 6 bytes
            rtf: Some("123456".to_string()),       // 6 bytes
            png_base64: Some("AAAA".to_string()),  // 3 decoded bytes
        };

        let sanitized = sanitize_clipboard_content_with_limits(content, 5, 2);
        assert_eq!(sanitized.text, Some("hello".to_string()));
        assert_eq!(sanitized.html, None);
        assert_eq!(sanitized.rtf, None);
        assert_eq!(sanitized.png_base64, None);
    }

    #[test]
    fn sanitize_clipboard_content_keeps_fields_within_limits() {
        let content = ClipboardContent {
            text: Some("hello".to_string()),
            html: Some("12345".to_string()),     // 5 bytes
            rtf: Some("12345".to_string()),      // 5 bytes
            png_base64: Some("AA==".to_string()), // 1 decoded byte
        };

        let sanitized = sanitize_clipboard_content_with_limits(content.clone(), 5, 2);
        assert_eq!(sanitized, content);
    }

    #[test]
    fn validate_rejects_oversized_html() {
        let payload = ClipboardWritePayload {
            text: Some("hello".to_string()),
            html: Some("x".repeat(MAX_RICH_TEXT_BYTES + 1)),
            rtf: None,
            png_base64: None,
        };

        let err = payload.validate().expect_err("expected size check to fail");
        match err {
            ClipboardError::InvalidPayload(msg) => assert!(msg.contains("html exceeds maximum size")),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn validate_rejects_oversized_rtf() {
        let payload = ClipboardWritePayload {
            text: Some("hello".to_string()),
            html: None,
            rtf: Some("x".repeat(MAX_RICH_TEXT_BYTES + 1)),
            png_base64: None,
        };

        let err = payload.validate().expect_err("expected size check to fail");
        match err {
            ClipboardError::InvalidPayload(msg) => assert!(msg.contains("rtf exceeds maximum size")),
            other => panic!("unexpected error: {other:?}"),
        }
    }
}
