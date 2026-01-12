use base64::{engine::general_purpose::STANDARD, Engine as _};
use serde::{Deserialize, Serialize};

pub mod platform;

mod cf_html;

// Clipboard items can contain extremely large rich payloads (especially images).
// Guard against unbounded memory usage / IPC payload sizes by skipping oversized formats.
//
// Oversized clipboard items are omitted (not treated as an error) so paste operations remain
// responsive.

/// Maximum number of raw PNG bytes we will read from the OS clipboard.
///
/// Large images (e.g. screenshots) can easily reach tens of MB. Returning them requires base64
/// encoding for IPC transport, which can cause expensive transfers and OOM the app. Oversized
/// clipboard items are omitted (not treated as an error).
pub const MAX_PNG_BYTES: usize = 5 * 1024 * 1024;

/// Maximum number of UTF-8 bytes we will read for string clipboard formats (text/plain, text/html,
/// text/rtf).
pub const MAX_TEXT_BYTES: usize = 2 * 1024 * 1024;

// Internal aliases used by the backend for "image" (PNG) and rich text limits.
const MAX_IMAGE_BYTES: usize = MAX_PNG_BYTES;
const MAX_RICH_TEXT_BYTES: usize = MAX_TEXT_BYTES;

fn normalize_base64_str(mut base64: &str) -> &str {
    base64 = base64.trim();
    if base64
        .get(0..5)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("data:"))
    {
        // Scan only a small prefix for the comma separator so malformed inputs like
        // `data:AAAAA...` don't force an O(n) search over huge payloads.
        let comma = base64
            .as_bytes()
            .iter()
            .take(1024)
            .position(|&b| b == b',');
        if let Some(comma) = comma {
            base64 = &base64[comma + 1..];
        } else {
            // Malformed data URL (missing comma separator). Treat as empty so callers don't
            // accidentally decode `data:...` as base64.
            return "";
        }
    }
    base64.trim()
}

fn estimate_base64_decoded_len(base64: &str) -> Option<usize> {
    let s = normalize_base64_str(base64);
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
#[cfg(all(target_os = "windows", feature = "desktop"))]
mod windows;

// Keep the GTK-backed Linux clipboard implementation behind the `desktop` feature for production
// builds, but still compile the module in unit tests so we can validate pure helper logic (e.g.
// target selection) without requiring a system GTK/WebKit toolchain.
#[cfg(all(target_os = "linux", any(feature = "desktop", test)))]
mod linux;

#[cfg(all(target_os = "macos", feature = "desktop"))]
mod macos;

/// Returns `true` if a payload of `len` bytes is within the provided limit.
#[inline]
pub fn within_limit(len: usize, max_bytes: usize) -> bool {
    len <= max_bytes
}

/// Returns `Some(value)` if the UTF-8 byte length is within `max_bytes`; otherwise `None`.
#[inline]
pub fn string_within_limit(value: String, max_bytes: usize) -> Option<String> {
    within_limit(value.as_bytes().len(), max_bytes).then_some(value)
}

/// Returns `Some(base64)` if `bytes.len()` is within `max_bytes`; otherwise `None`.
#[inline]
pub fn bytes_to_base64_within_limit(bytes: &[u8], max_bytes: usize) -> Option<String> {
    if !within_limit(bytes.len(), max_bytes) {
        return None;
    }
    Some(STANDARD.encode(bytes))
}

/// Clipboard contents read from the OS.
///
/// This intentionally carries multiple representations of the same copied content
/// (e.g. plain text + HTML + RTF + PNG) so downstream paste handlers can pick the
/// richest available format.
#[derive(Clone, Debug, Default, Serialize, Deserialize, PartialEq)]
pub struct ClipboardContent {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtf: Option<String>,
    /// PNG bytes encoded as base64.
    #[serde(skip_serializing_if = "Option::is_none", alias = "png_base64", alias = "pngBase64")]
    pub image_png_base64: Option<String>,
}

/// Payload written to the OS clipboard.
#[derive(Clone, Debug, Default, Serialize, Deserialize, PartialEq)]
pub struct ClipboardWritePayload {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtf: Option<String>,
    /// PNG bytes encoded as base64.
    #[serde(skip_serializing_if = "Option::is_none", alias = "png_base64", alias = "pngBase64")]
    pub image_png_base64: Option<String>,
}

impl ClipboardWritePayload {
    fn validate_with_limits(
        &self,
        max_rich_text_bytes: usize,
        max_image_bytes: usize,
    ) -> Result<(), ClipboardError> {
        let has_png = self
            .image_png_base64
            .as_deref()
            .map(normalize_base64_str)
            .is_some_and(|s| !s.is_empty());

        if self.text.is_none() && self.html.is_none() && self.rtf.is_none() && !has_png {
            return Err(ClipboardError::InvalidPayload(
                "must include at least one of text, html, rtf, image_png_base64".to_string(),
            ));
        }

        if let Some(text) = self.text.as_deref() {
            let len = text.as_bytes().len();
            if len > max_rich_text_bytes {
                return Err(ClipboardError::InvalidPayload(format!(
                    "text exceeds maximum size ({len} > {max_rich_text_bytes} bytes)"
                )));
            }
        }

        if let Some(html) = self.html.as_deref() {
            let len = html.as_bytes().len();
            if len > max_rich_text_bytes {
                return Err(ClipboardError::InvalidPayload(format!(
                    "html exceeds maximum size ({len} > {max_rich_text_bytes} bytes)"
                )));
            }
        }

        if let Some(rtf) = self.rtf.as_deref() {
            let len = rtf.as_bytes().len();
            if len > max_rich_text_bytes {
                return Err(ClipboardError::InvalidPayload(format!(
                    "rtf exceeds maximum size ({len} > {max_rich_text_bytes} bytes)"
                )));
            }
        }

        if has_png {
            let decoded_len =
                estimate_base64_decoded_len(self.image_png_base64.as_deref().unwrap_or(""))
                    .unwrap_or(usize::MAX);
            if decoded_len > max_image_bytes {
                return Err(ClipboardError::InvalidPayload(format!(
                    "pngBase64 exceeds maximum size ({decoded_len} > {max_image_bytes} bytes)"
                )));
            }
        }
        Ok(())
    }

    pub fn validate(&self) -> Result<(), ClipboardError> {
        self.validate_with_limits(MAX_RICH_TEXT_BYTES, MAX_IMAGE_BYTES)
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
    if matches!(content.text, Some(ref s) if s.as_bytes().len() > max_rich_text_bytes) {
        content.text = None;
    }
    if matches!(content.html, Some(ref s) if s.as_bytes().len() > max_rich_text_bytes) {
        content.html = None;
    }
    if matches!(content.rtf, Some(ref s) if s.as_bytes().len() > max_rich_text_bytes) {
        content.rtf = None;
    }
    if let Some(ref s) = content.image_png_base64 {
        let decoded_len = estimate_base64_decoded_len(s).unwrap_or(usize::MAX);
        // Normalize legacy `data:*;base64,...` prefixes so callers see consistent wire format.
        let normalized = normalize_base64_str(s);
        if decoded_len > max_image_bytes {
            content.image_png_base64 = None;
        } else if normalized.is_empty() {
            content.image_png_base64 = None;
        } else if normalized.len() != s.len() {
            content.image_png_base64 = Some(normalized.to_string());
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
pub async fn clipboard_read(window: tauri::WebviewWindow) -> Result<ClipboardContent, String> {
    crate::ipc_origin::ensure_main_window(
        window.label(),
        "clipboard access",
        crate::ipc_origin::Verb::Is,
    )?;
    let url = window.url().map_err(|err| err.to_string())?;
    crate::ipc_origin::ensure_trusted_origin(&url, "clipboard access", crate::ipc_origin::Verb::Is)?;

    // Clipboard APIs on macOS call into AppKit, which is not thread-safe.
    // Dispatch to the main thread before touching NSPasteboard.
    #[cfg(target_os = "macos")]
    {
        use tauri::Manager as _;
        return window
            .app_handle()
            .run_on_main_thread(read)
            .map_err(|e| e.to_string())?
            .map_err(|e| e.to_string());
    }

    #[cfg(not(target_os = "macos"))]
    {
        tauri::async_runtime::spawn_blocking(|| read().map_err(|e| e.to_string()))
            .await
            .map_err(|e| e.to_string())?
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn clipboard_write(
    window: tauri::WebviewWindow,
    payload: ClipboardWritePayload,
) -> Result<(), String> {
    crate::ipc_origin::ensure_main_window(
        window.label(),
        "clipboard access",
        crate::ipc_origin::Verb::Is,
    )?;
    let url = window.url().map_err(|err| err.to_string())?;
    crate::ipc_origin::ensure_trusted_origin(&url, "clipboard access", crate::ipc_origin::Verb::Is)?;

    // See `clipboard_read` for why we dispatch to the main thread on macOS.
    #[cfg(target_os = "macos")]
    {
        use tauri::Manager as _;
        return window
            .app_handle()
            .run_on_main_thread(move || write(&payload))
            .map_err(|e| e.to_string())?
            .map_err(|e| e.to_string());
    }

    #[cfg(not(target_os = "macos"))]
    {
        tauri::async_runtime::spawn_blocking(move || write(&payload).map_err(|e| e.to_string()))
            .await
            .map_err(|e| e.to_string())?
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn estimate_base64_decoded_len_handles_padding() {
        assert_eq!(estimate_base64_decoded_len(""), Some(0));
        assert_eq!(estimate_base64_decoded_len("Zg=="), Some(1)); // "f"
        assert_eq!(estimate_base64_decoded_len("Zm8="), Some(2)); // "fo"
        assert_eq!(estimate_base64_decoded_len("Zm9v"), Some(3)); // "foo"
        assert_eq!(estimate_base64_decoded_len("AAAA"), Some(3));
        assert_eq!(
            estimate_base64_decoded_len("data:image/png;base64"),
            Some(0),
            "malformed data URL without a comma should be treated as empty"
        );
        assert_eq!(
            estimate_base64_decoded_len("data:image/png;base64,Zm9v"),
            Some(3)
        );
        assert_eq!(
            estimate_base64_decoded_len("DATA:image/png;base64,Zm9v"),
            Some(3)
        );
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
            image_png_base64: Some("AAAA".to_string()), // 3 decoded bytes
        };

        let sanitized = sanitize_clipboard_content_with_limits(content, 5, 2);
        assert_eq!(sanitized.text, Some("hello".to_string()));
        assert_eq!(sanitized.html, None);
        assert_eq!(sanitized.rtf, None);
        assert_eq!(sanitized.image_png_base64, None);
    }

    #[test]
    fn sanitize_clipboard_content_keeps_fields_within_limits() {
        let content = ClipboardContent {
            text: Some("hello".to_string()),
            html: Some("12345".to_string()),     // 5 bytes
            rtf: Some("12345".to_string()),      // 5 bytes
            image_png_base64: Some("AA==".to_string()), // 1 decoded byte
        };

        let sanitized = sanitize_clipboard_content_with_limits(content.clone(), 5, 2);
        assert_eq!(sanitized, content);
    }

    #[test]
    fn sanitize_clipboard_content_normalizes_data_url_png_base64() {
        let content = ClipboardContent {
            text: None,
            html: None,
            rtf: None,
            image_png_base64: Some("data:image/png;base64,AAAA".to_string()),
        };

        let sanitized = sanitize_clipboard_content_with_limits(content, 5, 10);
        assert_eq!(sanitized.image_png_base64, Some("AAAA".to_string()));
    }

    #[test]
    fn validate_rejects_oversized_html() {
        let payload = ClipboardWritePayload {
            text: Some("hello".to_string()),
            html: Some("x".repeat(MAX_TEXT_BYTES + 1)),
            rtf: None,
            image_png_base64: None,
        };

        let err = payload.validate().expect_err("expected size check to fail");
        match err {
            ClipboardError::InvalidPayload(msg) => assert!(msg.contains("html exceeds maximum size")),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn validate_rejects_oversized_text() {
        let payload = ClipboardWritePayload {
            text: Some("x".repeat(MAX_TEXT_BYTES + 1)),
            html: None,
            rtf: None,
            image_png_base64: None,
        };

        let err = payload.validate().expect_err("expected size check to fail");
        match err {
            ClipboardError::InvalidPayload(msg) => assert!(msg.contains("text exceeds maximum size")),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn validate_rejects_oversized_rtf() {
        let payload = ClipboardWritePayload {
            text: Some("hello".to_string()),
            html: None,
            rtf: Some("x".repeat(MAX_TEXT_BYTES + 1)),
            image_png_base64: None,
        };

        let err = payload.validate().expect_err("expected size check to fail");
        match err {
            ClipboardError::InvalidPayload(msg) => assert!(msg.contains("rtf exceeds maximum size")),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn validate_rejects_oversized_png_base64() {
        let payload = ClipboardWritePayload {
            text: Some("hello".to_string()),
            html: None,
            rtf: None,
            // "AAAA" decodes to 3 bytes.
            image_png_base64: Some("AAAA".to_string()),
        };

        let err = payload
            .validate_with_limits(MAX_RICH_TEXT_BYTES, 2)
            .expect_err("expected size check to fail");
        match err {
            ClipboardError::InvalidPayload(msg) => assert!(msg.contains("pngBase64 exceeds maximum size")),
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn validate_rejects_empty_png_base64_after_normalization() {
        let payload = ClipboardWritePayload {
            text: None,
            html: None,
            rtf: None,
            image_png_base64: Some("data:image/png;base64,".to_string()),
        };

        let err = payload.validate().expect_err("expected validation to fail");
        match err {
            ClipboardError::InvalidPayload(msg) => {
                assert!(msg.contains("must include at least one of text"));
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn validate_rejects_malformed_data_url_png_base64_missing_comma() {
        let payload = ClipboardWritePayload {
            text: None,
            html: None,
            rtf: None,
            image_png_base64: Some("data:image/png;base64".to_string()),
        };

        let err = payload.validate().expect_err("expected validation to fail");
        match err {
            ClipboardError::InvalidPayload(msg) => {
                assert!(msg.contains("must include at least one of text"));
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn within_limit_allows_equal() {
        assert!(within_limit(10, 10));
    }

    #[test]
    fn within_limit_rejects_over() {
        assert!(!within_limit(11, 10));
    }

    #[test]
    fn string_within_limit_keeps_small_strings() {
        assert_eq!(
            string_within_limit("abc".to_string(), 3),
            Some("abc".to_string())
        );
    }

    #[test]
    fn string_within_limit_drops_large_strings() {
        assert_eq!(string_within_limit("abcd".to_string(), 3), None);
    }

    #[test]
    fn bytes_to_base64_within_limit_keeps_small_payloads() {
        let bytes = b"hello";
        let b64 = bytes_to_base64_within_limit(bytes, 5).expect("expected base64");
        assert_eq!(b64, "aGVsbG8=");
    }

    #[test]
    fn bytes_to_base64_within_limit_drops_large_payloads() {
        let bytes = [0u8; 6];
        assert_eq!(bytes_to_base64_within_limit(&bytes, 5), None);
    }

    #[test]
    fn serde_contract_uses_image_png_base64_and_accepts_legacy_aliases() {
        let content = ClipboardContent {
            text: None,
            html: None,
            rtf: None,
            image_png_base64: Some("CQgH".to_string()),
        };
        let json = serde_json::to_value(&content).expect("serialize ClipboardContent");

        // The canonical wire key is `image_png_base64` (snake_case) to match the existing JS
        // clipboard provider contract.
        assert_eq!(json.get("image_png_base64").and_then(|v| v.as_str()), Some("CQgH"));
        assert!(json.get("pngBase64").is_none());
        assert!(json.get("png_base64").is_none());

        // But we must remain backwards-compatible with older/alternate bridges that used
        // `pngBase64` or `png_base64`.
        let from_png_base64: ClipboardContent =
            serde_json::from_value(serde_json::json!({ "png_base64": "CQgH" }))
                .expect("deserialize from png_base64");
        assert_eq!(from_png_base64.image_png_base64.as_deref(), Some("CQgH"));

        let from_png_base64_camel: ClipboardContent =
            serde_json::from_value(serde_json::json!({ "pngBase64": "CQgH" }))
                .expect("deserialize from pngBase64");
        assert_eq!(from_png_base64_camel.image_png_base64.as_deref(), Some("CQgH"));

        let payload_from_png_base64: ClipboardWritePayload =
            serde_json::from_value(serde_json::json!({ "pngBase64": "CQgH" }))
                .expect("deserialize write payload from pngBase64");
        assert_eq!(payload_from_png_base64.image_png_base64.as_deref(), Some("CQgH"));

        let payload = ClipboardWritePayload {
            text: None,
            html: None,
            rtf: None,
            image_png_base64: Some("CQgH".to_string()),
        };
        let payload_json = serde_json::to_value(&payload).expect("serialize ClipboardWritePayload");
        assert_eq!(
            payload_json.get("image_png_base64").and_then(|v| v.as_str()),
            Some("CQgH")
        );
    }
}
