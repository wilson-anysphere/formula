use base64::{engine::general_purpose::STANDARD, Engine as _};
use serde::{Deserialize, Serialize};
use std::sync::atomic::{AtomicU8, Ordering};

use crate::resource_limits::LimitedString;

pub mod platform;

mod cf_html;
#[cfg(any(target_os = "windows", test))]
mod retry;

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

/// Maximum number of raw TIFF bytes we will read/write on the macOS clipboard.
///
/// Some native macOS apps prefer `public.tiff` on the pasteboard even when PNG is present.
/// Unfortunately, TIFF encodings can be significantly larger than their PNG equivalents for the
/// same image. We therefore allow a larger (but still bounded) TIFF payload size so we can attach
/// TIFF representations for interoperability without weakening IPC guardrails (PNG is still capped
/// at [`MAX_PNG_BYTES`]).
pub const MAX_TIFF_BYTES: usize = 4 * MAX_PNG_BYTES;

/// Maximum number of UTF-8 bytes we will read for string clipboard formats (text/plain, text/html,
/// text/rtf).
pub const MAX_TEXT_BYTES: usize = 2 * 1024 * 1024;

// Internal aliases used by the backend for "image" (PNG) and rich text limits.
const MAX_IMAGE_BYTES: usize = MAX_PNG_BYTES;
pub const MAX_RICH_TEXT_BYTES: usize = MAX_TEXT_BYTES;

/// Maximum number of UTF-8 bytes accepted for plain-text-only clipboard writes over IPC.
///
/// The rich clipboard IPC path (`clipboard_write`) is capped at [`MAX_TEXT_BYTES`] (2 MiB) to keep
/// multi-format payloads small. Some operations can legitimately exceed that size (e.g. copying a
/// very large worksheet range as plain text). This separate cap allows best-effort large
/// `text/plain` writes while still bounding allocations during deserialization.
pub const MAX_PLAINTEXT_WRITE_BYTES: usize = 10 * 1024 * 1024; // 10 MiB

// We support optional `data:*;base64,` prefixes for backwards compatibility, but we intentionally
// scan only a small prefix for the comma separator so malformed inputs like `data:AAAA...` don't
// force an O(n) search over huge payloads.
const DATA_URL_COMMA_SCAN_LIMIT: usize = 1024;

const fn base64_encoded_len(decoded_len: usize) -> usize {
    // Base64 encodes 3 bytes as 4 chars, rounded up.
    ((decoded_len + 2) / 3) * 4
}

/// Maximum number of UTF-8 bytes accepted for `image_png_base64` over IPC.
///
/// This is a conservative cap on the encoded size, enforced during deserialization to prevent a
/// compromised WebView from forcing large allocations. The decoded byte limit is still enforced by
/// [`ClipboardWritePayload::validate`] as defense-in-depth.
pub const MAX_IMAGE_PNG_BASE64_BYTES: usize =
    base64_encoded_len(MAX_IMAGE_BYTES) + DATA_URL_COMMA_SCAN_LIMIT;

// ---------------------------------------------------------------------------
// Debug logging
//
// Clipboard interop issues are notoriously hard to diagnose in the field (formats vary per app,
// platform, and clipboard manager). Provide an opt-in, lightweight logging mechanism that can be
// enabled without attaching a native debugger.
//
// IMPORTANT: Do not log clipboard contents (privacy). Only log format names and sizes.

const DEBUG_CLIPBOARD_ENV_VAR: &str = "FORMULA_DEBUG_CLIPBOARD";

// 0 = unknown, 1 = disabled, 2 = enabled
static DEBUG_CLIPBOARD_ENABLED: AtomicU8 = AtomicU8::new(0);

#[inline]
fn env_truthy(raw: &str) -> bool {
    let v = raw.trim().to_ascii_lowercase();
    !(v.is_empty() || v == "0" || v == "false")
}

#[inline]
fn debug_clipboard_enabled() -> bool {
    match DEBUG_CLIPBOARD_ENABLED.load(Ordering::Relaxed) {
        1 => false,
        2 => true,
        _ => {
            let enabled = std::env::var(DEBUG_CLIPBOARD_ENV_VAR)
                .ok()
                .is_some_and(|raw| env_truthy(&raw));
            DEBUG_CLIPBOARD_ENABLED.store(if enabled { 2 } else { 1 }, Ordering::Relaxed);
            enabled
        }
    }
}

#[inline]
fn debug_clipboard_log(args: std::fmt::Arguments<'_>) {
    if !debug_clipboard_enabled() {
        return;
    }
    crate::stdio::stderrln(format_args!("[clipboard] {}", args));
}

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
            .take(DATA_URL_COMMA_SCAN_LIMIT)
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

// ---------------------------------------------------------------------------
// Image dimension helpers (macOS TIFF/PNG conversion guards)
//
// Some platforms require converting images to/from other formats for clipboard interoperability
// (e.g. macOS apps often put `public.tiff` on the pasteboard). Even when the encoded payload is
// small, highly-compressible images can expand to extremely large decoded pixel buffers.
//
// These helpers provide cheap (header-only) dimension parsing so conversion code can skip
// pathological images before asking platform APIs to decode them.

/// Maximum number of bytes we're willing to allocate for a decoded RGBA buffer when converting
/// between image formats.
///
/// This is intentionally larger than [`MAX_PNG_BYTES`] because decoded pixel buffers are
/// uncompressed, but still bounded to avoid exhausting memory.
#[cfg(any(target_os = "macos", test))]
const MAX_DECODED_IMAGE_BYTES: usize = MAX_TIFF_BYTES;

#[cfg(any(target_os = "macos", test))]
fn decoded_rgba_len(width: u32, height: u32) -> Option<usize> {
    if width == 0 || height == 0 {
        return None;
    }
    let w = usize::try_from(width).ok()?;
    let h = usize::try_from(height).ok()?;
    w.checked_mul(h)?.checked_mul(4)
}

#[cfg(any(target_os = "macos", test))]
fn png_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    // Parse width/height from the IHDR chunk without decoding pixel data.
    const SIG: &[u8; 8] = b"\x89PNG\r\n\x1a\n";
    if bytes.len() < 24 {
        return None;
    }
    if bytes.get(0..8)? != SIG {
        return None;
    }
    if bytes.get(12..16)? != b"IHDR" {
        return None;
    }
    let w = u32::from_be_bytes(bytes.get(16..20)?.try_into().ok()?);
    let h = u32::from_be_bytes(bytes.get(20..24)?.try_into().ok()?);
    Some((w, h))
}

#[cfg(any(target_os = "macos", test))]
#[derive(Clone, Copy)]
enum TiffEndian {
    Little,
    Big,
}

#[cfg(any(target_os = "macos", test))]
fn read_u16_tiff(endian: TiffEndian, bytes: &[u8], offset: usize) -> Option<u16> {
    let b = bytes.get(offset..offset + 2)?;
    Some(match endian {
        TiffEndian::Little => u16::from_le_bytes([b[0], b[1]]),
        TiffEndian::Big => u16::from_be_bytes([b[0], b[1]]),
    })
}

#[cfg(any(target_os = "macos", test))]
fn read_u32_tiff(endian: TiffEndian, bytes: &[u8], offset: usize) -> Option<u32> {
    let b = bytes.get(offset..offset + 4)?;
    Some(match endian {
        TiffEndian::Little => u32::from_le_bytes([b[0], b[1], b[2], b[3]]),
        TiffEndian::Big => u32::from_be_bytes([b[0], b[1], b[2], b[3]]),
    })
}

#[cfg(any(target_os = "macos", test))]
fn read_u64_tiff(endian: TiffEndian, bytes: &[u8], offset: usize) -> Option<u64> {
    let b = bytes.get(offset..offset + 8)?;
    Some(match endian {
        TiffEndian::Little => u64::from_le_bytes([b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7]]),
        TiffEndian::Big => u64::from_be_bytes([b[0], b[1], b[2], b[3], b[4], b[5], b[6], b[7]]),
    })
}

#[cfg(any(target_os = "macos", test))]
fn type_size_tiff(ty: u16) -> Option<usize> {
    // TIFF field types (subset).
    match ty {
        1 => Some(1),  // BYTE
        2 => Some(1),  // ASCII
        3 => Some(2),  // SHORT
        4 => Some(4),  // LONG
        16 => Some(8), // LONG8 (BigTIFF)
        _ => None,
    }
}

#[cfg(any(target_os = "macos", test))]
fn read_tiff_value_as_u32(
    bytes: &[u8],
    endian: TiffEndian,
    ty: u16,
    count: u64,
    value_field_offset: usize,
    value_or_offset: u64,
    max_inline_bytes: usize,
) -> Option<u32> {
    if count == 0 {
        return None;
    }
    let type_size = type_size_tiff(ty)?;
    let total_size = type_size.checked_mul(usize::try_from(count).ok()?)?;

    let read_at = if total_size <= max_inline_bytes {
        // Inline value stored directly in the value field.
        value_field_offset
    } else {
        usize::try_from(value_or_offset).ok()?
    };

    match ty {
        3 => read_u16_tiff(endian, bytes, read_at).map(u32::from), // SHORT
        4 => read_u32_tiff(endian, bytes, read_at),                // LONG
        16 => read_u64_tiff(endian, bytes, read_at).and_then(|v| u32::try_from(v).ok()), // LONG8
        _ => None,
    }
}

#[cfg(any(target_os = "macos", test))]
fn parse_tiff_ifd_standard(bytes: &[u8], endian: TiffEndian, offset: usize) -> Option<(u32, u32)> {
    let entry_count = usize::from(read_u16_tiff(endian, bytes, offset)?);
    let mut width = None;
    let mut height = None;

    let entries_base = offset.checked_add(2)?;
    for i in 0..entry_count {
        let entry_off = entries_base.checked_add(i.checked_mul(12)?)?;
        let tag = read_u16_tiff(endian, bytes, entry_off)?;
        let ty = read_u16_tiff(endian, bytes, entry_off + 2)?;
        let count = u64::from(read_u32_tiff(endian, bytes, entry_off + 4)?);
        let value_field_offset = entry_off + 8;
        let value_or_offset = u64::from(read_u32_tiff(endian, bytes, value_field_offset)?);

        if tag == 256 || tag == 257 {
            let v = read_tiff_value_as_u32(
                bytes,
                endian,
                ty,
                count,
                value_field_offset,
                value_or_offset,
                4,
            )?;
            if tag == 256 {
                width = Some(v);
            } else {
                height = Some(v);
            }
            if width.is_some() && height.is_some() {
                break;
            }
        }
    }

    match (width, height) {
        (Some(w), Some(h)) if w > 0 && h > 0 => Some((w, h)),
        _ => None,
    }
}

#[cfg(any(target_os = "macos", test))]
fn parse_tiff_ifd_bigtiff(bytes: &[u8], endian: TiffEndian, offset: usize) -> Option<(u32, u32)> {
    let entry_count = usize::try_from(read_u64_tiff(endian, bytes, offset)?).ok()?;
    let mut width = None;
    let mut height = None;

    let entries_base = offset.checked_add(8)?;
    for i in 0..entry_count {
        let entry_off = entries_base.checked_add(i.checked_mul(20)?)?;
        let tag = read_u16_tiff(endian, bytes, entry_off)?;
        let ty = read_u16_tiff(endian, bytes, entry_off + 2)?;
        let count = read_u64_tiff(endian, bytes, entry_off + 4)?;
        let value_field_offset = entry_off + 12;
        let value_or_offset = read_u64_tiff(endian, bytes, value_field_offset)?;

        if tag == 256 || tag == 257 {
            let v = read_tiff_value_as_u32(
                bytes,
                endian,
                ty,
                count,
                value_field_offset,
                value_or_offset,
                8,
            )?;
            if tag == 256 {
                width = Some(v);
            } else {
                height = Some(v);
            }
            if width.is_some() && height.is_some() {
                break;
            }
        }
    }

    match (width, height) {
        (Some(w), Some(h)) if w > 0 && h > 0 => Some((w, h)),
        _ => None,
    }
}

/// Parse TIFF width/height from the first IFD (supports classic TIFF + BigTIFF).
#[cfg(any(target_os = "macos", test))]
fn tiff_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    if bytes.len() < 8 {
        return None;
    }

    let endian = match bytes.get(0..2)? {
        b"II" => TiffEndian::Little,
        b"MM" => TiffEndian::Big,
        _ => return None,
    };

    let magic = read_u16_tiff(endian, bytes, 2)?;
    match magic {
        42 => {
            // Classic TIFF: offset to first IFD is a u32 at byte 4.
            let ifd_offset = usize::try_from(read_u32_tiff(endian, bytes, 4)?).ok()?;
            parse_tiff_ifd_standard(bytes, endian, ifd_offset)
        }
        43 => {
            // BigTIFF: offset size (u16) at byte 4, then u64 first IFD offset at byte 8.
            if bytes.len() < 16 {
                return None;
            }
            let offset_size = read_u16_tiff(endian, bytes, 4)?;
            if offset_size != 8 {
                return None;
            }
            let ifd_offset = usize::try_from(read_u64_tiff(endian, bytes, 8)?).ok()?;
            parse_tiff_ifd_bigtiff(bytes, endian, ifd_offset)
        }
        _ => None,
    }
}

#[cfg(any(target_os = "windows", test))]
mod windows_dib;
#[cfg(any(target_os = "windows", test))]
mod windows_format_cache;
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

/// Convert UTF-8-ish clipboard bytes into a `String`, trimming trailing NUL terminators.
///
/// Some clipboard producers (notably on Windows and occasionally via macOS NSPasteboard `dataForType`)
/// include an extra `\0` terminator even when the payload is length-delimited. We strip any trailing
/// NUL bytes/characters for consistency across platforms.
///
/// Returns `None` for empty inputs or when the trimmed output is empty.
#[cfg(any(test, all(target_os = "macos", feature = "desktop")))]
fn bytes_to_string_trim_nuls(bytes: &[u8]) -> Option<String> {
    if bytes.is_empty() {
        return None;
    }
    let s = String::from_utf8_lossy(bytes);
    let s = s.trim_end_matches('\0');
    if s.is_empty() {
        None
    } else {
        Some(s.to_string())
    }
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

/// IPC-only clipboard write payload.
///
/// This mirrors [`ClipboardWritePayload`] but uses bounded string types so oversized inputs fail
/// fast during deserialization (before allocating multi-megabyte `String`s).
#[derive(Clone, Debug, Default, Serialize, Deserialize, PartialEq)]
pub struct ClipboardWritePayloadIpc {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<LimitedString<MAX_RICH_TEXT_BYTES>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub html: Option<LimitedString<MAX_RICH_TEXT_BYTES>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rtf: Option<LimitedString<MAX_RICH_TEXT_BYTES>>,
    /// PNG bytes encoded as base64.
    #[serde(skip_serializing_if = "Option::is_none", alias = "png_base64", alias = "pngBase64")]
    pub image_png_base64: Option<LimitedString<MAX_IMAGE_PNG_BASE64_BYTES>>,
}

impl From<ClipboardWritePayloadIpc> for ClipboardWritePayload {
    fn from(value: ClipboardWritePayloadIpc) -> Self {
        ClipboardWritePayload {
            text: value.text.map(LimitedString::into_inner),
            html: value.html.map(LimitedString::into_inner),
            rtf: value.rtf.map(LimitedString::into_inner),
            image_png_base64: value.image_png_base64.map(LimitedString::into_inner),
        }
    }
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
    let sanitized = sanitize_clipboard_content_with_limits(
        content,
        MAX_RICH_TEXT_BYTES,
        MAX_IMAGE_BYTES,
    );

    // Cross-platform summary (format names + byte counts only).
    if debug_clipboard_enabled() {
        let text_bytes = sanitized.text.as_ref().map(|s| s.as_bytes().len());
        let html_bytes = sanitized.html.as_ref().map(|s| s.as_bytes().len());
        let rtf_bytes = sanitized.rtf.as_ref().map(|s| s.as_bytes().len());
        let png_bytes = sanitized
            .image_png_base64
            .as_deref()
            .and_then(estimate_base64_decoded_len);
        debug_clipboard_log(format_args!(
            "read summary: text_bytes={text_bytes:?} html_bytes={html_bytes:?} rtf_bytes={rtf_bytes:?} png_bytes={png_bytes:?} caps(text={MAX_TEXT_BYTES}, png={MAX_PNG_BYTES})"
        ));
    }

    Ok(sanitized)
}

pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    payload.validate()?;
    if debug_clipboard_enabled() {
        let text_bytes = payload.text.as_ref().map(|s| s.as_bytes().len());
        let html_bytes = payload.html.as_ref().map(|s| s.as_bytes().len());
        let rtf_bytes = payload.rtf.as_ref().map(|s| s.as_bytes().len());
        let png_bytes = payload
            .image_png_base64
            .as_deref()
            .and_then(estimate_base64_decoded_len);
        debug_clipboard_log(format_args!(
            "write summary: text_bytes={text_bytes:?} html_bytes={html_bytes:?} rtf_bytes={rtf_bytes:?} png_bytes={png_bytes:?} caps(text={MAX_TEXT_BYTES}, png={MAX_PNG_BYTES})"
        ));
    }
    platform::write(payload)
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn clipboard_read(window: tauri::WebviewWindow) -> Result<ClipboardContent, String> {
    let url = window.url().map_err(|err| err.to_string())?;
    crate::ipc_origin::ensure_main_window(window.label(), "clipboard access", crate::ipc_origin::Verb::Is)?;
    crate::ipc_origin::ensure_trusted_origin(&url, "clipboard access", crate::ipc_origin::Verb::Is)?;
    crate::ipc_origin::ensure_stable_origin(&window, "clipboard access", crate::ipc_origin::Verb::Is)?;

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
    payload: ClipboardWritePayloadIpc,
) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    crate::ipc_origin::ensure_main_window(window.label(), "clipboard access", crate::ipc_origin::Verb::Is)?;
    crate::ipc_origin::ensure_trusted_origin(&url, "clipboard access", crate::ipc_origin::Verb::Is)?;
    crate::ipc_origin::ensure_stable_origin(&window, "clipboard access", crate::ipc_origin::Verb::Is)?;

    let payload: ClipboardWritePayload = payload.into();

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

/// Best-effort plain-text clipboard write for payloads larger than [`MAX_TEXT_BYTES`].
///
/// This is used by the frontend clipboard provider as a fallback when the text is too large to be
/// sent over the rich multi-format clipboard IPC path. It is still bounded to avoid unbounded
/// allocations from a compromised WebView.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn clipboard_write_text(
    window: tauri::WebviewWindow,
    text: LimitedString<MAX_PLAINTEXT_WRITE_BYTES>,
) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    crate::ipc_origin::ensure_main_window(window.label(), "clipboard access", crate::ipc_origin::Verb::Is)?;
    crate::ipc_origin::ensure_trusted_origin(&url, "clipboard access", crate::ipc_origin::Verb::Is)?;
    crate::ipc_origin::ensure_stable_origin(&window, "clipboard access", crate::ipc_origin::Verb::Is)?;

    let payload = ClipboardWritePayload {
        text: Some(text.into_inner()),
        ..Default::default()
    };

    // Dispatch to main thread on macOS (AppKit is not thread-safe).
    #[cfg(target_os = "macos")]
    {
        use tauri::Manager as _;
        return window
            .app_handle()
            .run_on_main_thread(move || platform::write(&payload))
            .map_err(|e| e.to_string())?
            .map_err(|e| e.to_string());
    }

    #[cfg(not(target_os = "macos"))]
    {
        tauri::async_runtime::spawn_blocking(move || platform::write(&payload).map_err(|e| e.to_string()))
            .await
            .map_err(|e| e.to_string())?
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn env_truthy_parses_common_values() {
        assert!(!env_truthy(""));
        assert!(!env_truthy("  "));
        assert!(!env_truthy("0"));
        assert!(!env_truthy("false"));
        assert!(!env_truthy(" FALSE "));
        assert!(env_truthy("1"));
        assert!(env_truthy("true"));
        assert!(env_truthy(" TRUE "));
    }

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
    fn decoded_rgba_len_computes_pixel_buffer_size() {
        let len = decoded_rgba_len(100, 200).expect("expected valid dimensions");
        assert_eq!(len, 100 * 200 * 4);
        assert!(len < MAX_DECODED_IMAGE_BYTES);
        assert_eq!(decoded_rgba_len(0, 1), None);
        assert_eq!(decoded_rgba_len(1, 0), None);
    }

    #[test]
    fn png_dimensions_parses_ihdr_chunk() {
        // 1x1 transparent PNG.
        let png = base64::engine::general_purpose::STANDARD
            .decode(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO9C9VwAAAAASUVORK5CYII=",
            )
            .unwrap();
        assert_eq!(png_dimensions(&png), Some((1, 1)));
    }

    #[test]
    fn tiff_dimensions_parses_classic_little_endian_ifd() {
        // Minimal classic TIFF with ImageWidth=100, ImageLength=200.
        let tiff: Vec<u8> = vec![
            0x49, 0x49, // II
            0x2A, 0x00, // 42
            0x08, 0x00, 0x00, 0x00, // IFD offset = 8
            0x02, 0x00, // entry count = 2
            // Tag 256 (ImageWidth), type LONG, count 1, value 100
            0x00, 0x01, // tag
            0x04, 0x00, // type LONG
            0x01, 0x00, 0x00, 0x00, // count
            0x64, 0x00, 0x00, 0x00, // value
            // Tag 257 (ImageLength), type LONG, count 1, value 200
            0x01, 0x01, // tag
            0x04, 0x00, // type LONG
            0x01, 0x00, 0x00, 0x00, // count
            0xC8, 0x00, 0x00, 0x00, // value
            // next IFD offset = 0
            0x00, 0x00, 0x00, 0x00,
        ];

        assert_eq!(tiff_dimensions(&tiff), Some((100, 200)));
    }

    #[test]
    fn tiff_dimensions_parses_classic_big_endian_ifd() {
        // Minimal big-endian classic TIFF with ImageWidth=10 (SHORT), ImageLength=20 (SHORT).
        let tiff: Vec<u8> = vec![
            0x4D, 0x4D, // MM
            0x00, 0x2A, // 42
            0x00, 0x00, 0x00, 0x08, // IFD offset = 8
            0x00, 0x02, // entry count = 2
            // Tag 256 (ImageWidth), type SHORT, count 1, value 10 (padded)
            0x01, 0x00, // tag
            0x00, 0x03, // type SHORT
            0x00, 0x00, 0x00, 0x01, // count
            0x00, 0x0A, 0x00, 0x00, // value (big-endian SHORT + padding)
            // Tag 257 (ImageLength), type SHORT, count 1, value 20 (padded)
            0x01, 0x01, // tag
            0x00, 0x03, // type SHORT
            0x00, 0x00, 0x00, 0x01, // count
            0x00, 0x14, 0x00, 0x00, // value
            // next IFD offset = 0
            0x00, 0x00, 0x00, 0x00,
        ];

        assert_eq!(tiff_dimensions(&tiff), Some((10, 20)));
    }

    #[test]
    fn tiff_dimensions_parses_bigtiff_little_endian_ifd() {
        // Minimal BigTIFF (little-endian) with ImageWidth=1234, ImageLength=5678 (LONG).
        let mut tiff: Vec<u8> = vec![
            0x49, 0x49, // II
            0x2B, 0x00, // 43 (BigTIFF)
            0x08, 0x00, // offset size = 8
            0x00, 0x00, // reserved
            0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // first IFD offset = 16
            0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // entry count = 2 (u64)
        ];

        // Entry 1: tag 256, type LONG, count 1, value 1234 (inline, padded to 8 bytes).
        tiff.extend_from_slice(&[
            0x00, 0x01, // tag
            0x04, 0x00, // type LONG
            0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // count (u64)
            0xD2, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // value (u64)
        ]);

        // Entry 2: tag 257, type LONG, count 1, value 5678.
        tiff.extend_from_slice(&[
            0x01, 0x01, // tag
            0x04, 0x00, // type LONG
            0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // count (u64)
            0x2E, 0x16, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, // value (u64)
        ]);

        // next IFD offset = 0 (u64)
        tiff.extend_from_slice(&[0x00; 8]);

        assert_eq!(tiff_dimensions(&tiff), Some((1234, 5678)));
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
    fn bytes_to_string_trim_nuls_returns_none_for_empty_input() {
        assert_eq!(bytes_to_string_trim_nuls(b""), None);
    }

    #[test]
    fn bytes_to_string_trim_nuls_trims_trailing_nuls() {
        assert_eq!(
            bytes_to_string_trim_nuls(b"hello\0\0"),
            Some("hello".to_string())
        );
    }

    #[test]
    fn bytes_to_string_trim_nuls_preserves_interior_nuls() {
        assert_eq!(
            bytes_to_string_trim_nuls(b"he\0llo\0"),
            Some("he\0llo".to_string())
        );
    }

    #[test]
    fn bytes_to_string_trim_nuls_returns_none_when_only_nuls() {
        assert_eq!(bytes_to_string_trim_nuls(b"\0\0"), None);
    }

    #[test]
    fn bytes_to_string_trim_nuls_is_best_effort_for_invalid_utf8() {
        // 0xFF is invalid UTF-8 and will be replaced with U+FFFD.
        assert_eq!(
            bytes_to_string_trim_nuls(&[0xFF, 0x00]),
            Some("\u{FFFD}".to_string())
        );
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

        let ipc_payload_from_png_base64: ClipboardWritePayloadIpc =
            serde_json::from_value(serde_json::json!({ "pngBase64": "CQgH" }))
                .expect("deserialize IPC write payload from pngBase64");
        assert_eq!(
            ipc_payload_from_png_base64
                .image_png_base64
                .as_ref()
                .map(LimitedString::as_str),
            Some("CQgH")
        );

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

    #[test]
    fn limited_string_rejects_oversized_strings_during_deserialization() {
        #[derive(Debug, Deserialize)]
        struct TestPayload {
            #[allow(dead_code)]
            text: Option<LimitedString<5>>,
        }

        let err = serde_json::from_value::<TestPayload>(serde_json::json!({
            "text": "123456"
        }))
        .expect_err("expected deserialization to fail");
        assert!(
            err.to_string().contains("exceeds maximum size"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn limited_string_rejects_oversized_base64_during_deserialization() {
        #[derive(Debug, Deserialize)]
        struct TestPayload {
            #[allow(dead_code)]
            #[serde(alias = "pngBase64")]
            image_png_base64: Option<LimitedString<8>>,
        }

        let err = serde_json::from_value::<TestPayload>(serde_json::json!({
            "pngBase64": "AAAAAAAAA"
        }))
        .expect_err("expected deserialization to fail");
        assert!(
            err.to_string().contains("exceeds maximum size"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn limited_string_allows_payloads_within_limits() {
        #[derive(Debug, Deserialize)]
        struct TestPayload {
            text: Option<LimitedString<5>>,
            #[serde(alias = "pngBase64")]
            image_png_base64: Option<LimitedString<8>>,
        }

        let payload = serde_json::from_value::<TestPayload>(serde_json::json!({
            "text": "12345",
            "pngBase64": "AAAA"
        }))
        .expect("expected deserialization to succeed");

        assert_eq!(
            payload.text.map(LimitedString::into_inner),
            Some("12345".to_string())
        );
        assert_eq!(
            payload.image_png_base64.map(LimitedString::into_inner),
            Some("AAAA".to_string())
        );
    }
}
