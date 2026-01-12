use std::time::Duration;

use base64::{engine::general_purpose::STANDARD, Engine as _};
use windows::core::PCWSTR;
use windows::Win32::Foundation::{GlobalFree, HANDLE, HGLOBAL};
use windows::Win32::System::DataExchange::{
    CloseClipboard, EmptyClipboard, GetClipboardData, OpenClipboard, RegisterClipboardFormatW,
    SetClipboardData,
};
use windows::Win32::System::Memory::{
    GlobalAlloc, GlobalLock, GlobalSize, GlobalUnlock, GMEM_MOVEABLE,
};

use super::cf_html::{build_cf_html_payload, extract_cf_html_fragment_best_effort};
use super::windows_dib::{dibv5_to_png, png_to_dib_and_dibv5};
use super::{
    normalize_base64_str, string_within_limit, ClipboardContent, ClipboardError, ClipboardWritePayload,
    MAX_PNG_BYTES, MAX_TEXT_BYTES,
};

// Built-in clipboard formats that we use directly. Keeping these as numeric constants avoids
// needing Win32 System Ole bindings just for format IDs.
const CF_UNICODETEXT: u32 = 13;
const CF_DIB: u32 = 8;
const CF_DIBV5: u32 = 17;

// DIB payloads are uncompressed, so they can be significantly larger than the corresponding PNG.
// Allow a larger cap for DIB formats so we can still convert many clipboard images to PNG.
const MAX_DIB_BYTES: usize = 4 * MAX_PNG_BYTES;

struct ClipboardGuard;

impl Drop for ClipboardGuard {
    fn drop(&mut self) {
        unsafe {
            let _ = CloseClipboard();
        }
    }
}

fn win_err(context: &str, err: windows::core::Error) -> ClipboardError {
    ClipboardError::OperationFailed(format!("{context}: {err}"))
}

fn open_clipboard_with_retry() -> Result<ClipboardGuard, ClipboardError> {
    // The clipboard is often temporarily locked by another process. Retry briefly.
    const ATTEMPTS: usize = 10;
    for attempt in 0..ATTEMPTS {
        match unsafe { OpenClipboard(None) } {
            Ok(()) => return Ok(ClipboardGuard),
            Err(_) if attempt + 1 < ATTEMPTS => {
                // Small backoff.
                std::thread::sleep(Duration::from_millis(10));
                continue;
            }
            Err(err) => return Err(win_err("OpenClipboard failed", err)),
        }
    }
    unreachable!()
}

fn register_format(name: &str) -> Result<u32, ClipboardError> {
    let mut wide: Vec<u16> = name.encode_utf16().collect();
    wide.push(0);
    let id = unsafe { RegisterClipboardFormatW(PCWSTR(wide.as_ptr())) };
    if id == 0 {
        return Err(ClipboardError::OperationFailed(format!(
            "RegisterClipboardFormatW(\"{name}\") failed"
        )));
    }
    Ok(id)
}

fn try_get_clipboard_bytes(
    format: u32,
    max_bytes: usize,
) -> Result<Option<Vec<u8>>, ClipboardError> {
    let handle = match unsafe { GetClipboardData(format) } {
        Ok(handle) => handle,
        Err(_) => return Ok(None),
    };

    // Clipboard data handles for our formats are HGLOBALs.
    let hglobal = HGLOBAL(handle.0);
    let size = unsafe { GlobalSize(hglobal) };
    if size == 0 {
        return Ok(None);
    }
    if size > max_bytes {
        return Ok(None);
    }

    let ptr = unsafe { GlobalLock(hglobal) };
    if ptr.is_null() {
        return Err(ClipboardError::OperationFailed(
            "GlobalLock returned null".to_string(),
        ));
    }

    let slice = unsafe { std::slice::from_raw_parts(ptr as *const u8, size) };
    let out = slice.to_vec();

    // `GlobalUnlock` returns FALSE both when the lock count reaches zero and when it fails.
    // The Windows crate maps FALSE to an error, so ignore the return value and rely on OS
    // correctness.
    unsafe {
        let _ = GlobalUnlock(hglobal);
    }

    Ok(Some(out))
}

fn try_get_unicode_text() -> Result<Option<String>, ClipboardError> {
    let handle = match unsafe { GetClipboardData(CF_UNICODETEXT) } {
        Ok(handle) => handle,
        Err(_) => return Ok(None),
    };

    let hglobal = HGLOBAL(handle.0);
    let size_bytes = unsafe { GlobalSize(hglobal) };
    if size_bytes == 0 {
        return Ok(None);
    }
    // UTF-16 is always at least 2 bytes/code unit, and UTF-8 is always at least 1 byte/code unit.
    // If the UTF-16 backing store already exceeds 2x the maximum UTF-8 bytes we allow, we know the
    // decoded string would exceed the limit. Skip without allocating a giant String.
    if size_bytes > MAX_TEXT_BYTES.saturating_mul(2) {
        return Ok(None);
    }

    let ptr = unsafe { GlobalLock(hglobal) };
    if ptr.is_null() {
        return Err(ClipboardError::OperationFailed(
            "GlobalLock returned null".to_string(),
        ));
    }

    let len_u16 = size_bytes / 2;
    let slice = unsafe { std::slice::from_raw_parts(ptr as *const u16, len_u16) };
    let nul = slice.iter().position(|&c| c == 0).unwrap_or(slice.len());
    let text = String::from_utf16_lossy(&slice[..nul]);

    unsafe {
        let _ = GlobalUnlock(hglobal);
    }

    if text.is_empty() {
        Ok(None)
    } else {
        Ok(string_within_limit(text, MAX_TEXT_BYTES))
    }
}

fn set_clipboard_bytes(format: u32, bytes: &[u8]) -> Result<(), ClipboardError> {
    if bytes.is_empty() {
        return Ok(());
    }

    let hglobal = unsafe { GlobalAlloc(GMEM_MOVEABLE, bytes.len()) }.map_err(|e| {
        ClipboardError::OperationFailed(format!("GlobalAlloc failed ({} bytes): {e}", bytes.len()))
    })?;

    let ptr = unsafe { GlobalLock(hglobal) };
    if ptr.is_null() {
        unsafe {
            let _ = GlobalFree(Some(hglobal));
        }
        return Err(ClipboardError::OperationFailed(
            "GlobalLock returned null".to_string(),
        ));
    }

    unsafe {
        std::ptr::copy_nonoverlapping(bytes.as_ptr(), ptr as *mut u8, bytes.len());
        let _ = GlobalUnlock(hglobal);
    }

    let handle = HANDLE(hglobal.0);
    match unsafe { SetClipboardData(format, Some(handle)) } {
        Ok(_) => Ok(()),
        Err(err) => {
            // Ownership is only transferred to the system on success.
            unsafe {
                let _ = GlobalFree(Some(hglobal));
            }
            Err(win_err(
                &format!("SetClipboardData(format={format}) failed"),
                err,
            ))
        }
    }
}

fn set_unicode_text(text: &str) -> Result<(), ClipboardError> {
    let mut wide: Vec<u16> = text.encode_utf16().collect();
    wide.push(0);
    let bytes = unsafe { std::slice::from_raw_parts(wide.as_ptr() as *const u8, wide.len() * 2) };
    set_clipboard_bytes(CF_UNICODETEXT, bytes)
}

fn parse_png_dimensions(png_bytes: &[u8]) -> Option<(u32, u32)> {
    // Parse the IHDR chunk so we can estimate decoded sizes without fully decoding.
    const PNG_SIGNATURE: [u8; 8] = [0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A];

    if png_bytes.len() < 24 {
        return None;
    }
    if !png_bytes.starts_with(&PNG_SIGNATURE) {
        return None;
    }

    let ihdr_len = u32::from_be_bytes([png_bytes[8], png_bytes[9], png_bytes[10], png_bytes[11]])
        as usize;
    if ihdr_len < 8 {
        return None;
    }
    if png_bytes.get(12..16)? != b"IHDR" {
        return None;
    }

    let data_start = 16;
    let data = png_bytes.get(data_start..data_start + ihdr_len)?;
    let width = u32::from_be_bytes([data[0], data[1], data[2], data[3]]);
    let height = u32::from_be_bytes([data[4], data[5], data[6], data[7]]);
    if width == 0 || height == 0 {
        return None;
    }
    Some((width, height))
}

pub fn read() -> Result<ClipboardContent, ClipboardError> {
    // Rich clipboard formats are best-effort: if registration fails (rare), still return whatever
    // formats we can read (e.g. plain text).
    let format_html = register_format("HTML Format").ok();
    let format_rtf = register_format("Rich Text Format").ok();
    let format_png = register_format("PNG").ok();
    // Some producers use a MIME-like name for PNG. Treat this as best-effort fallback.
    let format_image_png = register_format("image/png").ok();
    // MIME-like aliases used by some cross-platform apps.
    let format_text_html = register_format("text/html").ok();
    let format_text_rtf = register_format("text/rtf").ok();

    let _guard = open_clipboard_with_retry()?;

    // Best-effort reads: don't fail the entire operation if a single format can't be decoded.
    let text = try_get_unicode_text().ok().flatten();

    let mut html = format_html
        .and_then(|format| try_get_clipboard_bytes(format, MAX_TEXT_BYTES).ok().flatten())
        .map(|bytes| extract_cf_html_fragment_best_effort(&bytes))
        .filter(|s| !s.is_empty());
    if html.is_none() {
        if let Some(format) = format_text_html {
            html = try_get_clipboard_bytes(format, MAX_TEXT_BYTES)
                .ok()
                .flatten()
                .map(|bytes| extract_cf_html_fragment_best_effort(&bytes))
                .filter(|s| !s.is_empty());
        }
    }
    let html = html.and_then(|s| string_within_limit(s, MAX_TEXT_BYTES));

    let mut rtf = format_rtf
        .and_then(|format| try_get_clipboard_bytes(format, MAX_TEXT_BYTES).ok().flatten())
        .map(|bytes| {
            String::from_utf8_lossy(&bytes)
                .trim_end_matches('\0')
                .to_string()
        })
        .filter(|s| !s.is_empty());
    if rtf.is_none() {
        if let Some(format) = format_text_rtf {
            rtf = try_get_clipboard_bytes(format, MAX_TEXT_BYTES)
                .ok()
                .flatten()
                .map(|bytes| {
                    String::from_utf8_lossy(&bytes)
                        .trim_end_matches('\0')
                        .to_string()
                })
                .filter(|s| !s.is_empty());
        }
    }
    let rtf = rtf.and_then(|s| string_within_limit(s, MAX_TEXT_BYTES));

    let mut png_base64 = format_png
        .and_then(|format| try_get_clipboard_bytes(format, MAX_PNG_BYTES).ok().flatten())
        .map(|png_bytes| STANDARD.encode(png_bytes));

    if png_base64.is_none() {
        if let Some(format) = format_image_png {
            png_base64 = try_get_clipboard_bytes(format, MAX_PNG_BYTES)
                .ok()
                .flatten()
                .map(|bytes| STANDARD.encode(bytes));
        }
    }

    if png_base64.is_none() {
        png_base64 = try_get_clipboard_bytes(CF_DIBV5, MAX_DIB_BYTES)
            .ok()
            .flatten()
            .and_then(|dib_bytes| dibv5_to_png(&dib_bytes).ok())
            .filter(|png| png.len() <= MAX_PNG_BYTES)
            .map(|png| STANDARD.encode(png));
    }

    if png_base64.is_none() {
        if let Some(dib_bytes) = try_get_clipboard_bytes(CF_DIB, MAX_DIB_BYTES).ok().flatten() {
            png_base64 = dibv5_to_png(&dib_bytes)
                .ok()
                .filter(|png| png.len() <= MAX_PNG_BYTES)
                .map(|png| STANDARD.encode(png));
        }
    }

    Ok(ClipboardContent {
        text,
        html,
        rtf,
        png_base64,
    })
}

pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    let html_bytes = payload
        .html
        .as_deref()
        .filter(|s| !s.is_empty())
        .map(|html| {
            build_cf_html_payload(html)
                .map_err(|e| ClipboardError::InvalidPayload(format!("invalid html: {e}")))
        })
        .transpose()?;
    let html_plain_bytes = payload
        .html
        .as_deref()
        .filter(|s| !s.is_empty())
        .map(|s| s.as_bytes().to_vec());
    let rtf_bytes = payload
        .rtf
        .as_deref()
        .filter(|s| !s.is_empty())
        .map(|s| s.as_bytes().to_vec());

    let png_bytes = payload
        .png_base64
        .as_deref()
        .map(normalize_base64_str)
        .filter(|s| !s.is_empty())
        .map(|s| {
            STANDARD
                .decode(s)
                .map_err(|e| ClipboardError::InvalidPayload(format!("invalid pngBase64: {e}")))
        })
        .transpose()?;

    // Interop: provide DIB representations when possible, but treat conversion as best-effort.
    //
    // DIBs are uncompressed; if the PNG has huge dimensions (e.g. a compression bomb), decoding it
    // into a full BGRA buffer can be extremely memory-intensive. Skip DIB generation when the
    // decoded pixel buffer would exceed our DIB size cap.
    let dibs = png_bytes.as_deref().and_then(|png| {
        if let Some((w, h)) = parse_png_dimensions(png) {
            let decoded_len = (w as usize)
                .checked_mul(h as usize)
                .and_then(|px| px.checked_mul(4))?;
            if decoded_len > MAX_DIB_BYTES {
                return None;
            }
        }

        png_to_dib_and_dibv5(png).ok()
    });

    let (dib_bytes, dibv5_bytes) = match dibs {
        Some((dib, dibv5)) => (Some(dib), Some(dibv5)),
        None => (None, None),
    };

    let format_html = html_bytes
        .as_ref()
        .map(|_| register_format("HTML Format"))
        .transpose()?;
    let format_text_html = html_plain_bytes
        .as_ref()
        .map(|_| register_format("text/html").ok())
        .flatten();
    let format_rtf = rtf_bytes
        .as_ref()
        .map(|_| register_format("Rich Text Format"))
        .transpose()?;
    let format_text_rtf = rtf_bytes
        .as_ref()
        .map(|_| register_format("text/rtf").ok())
        .flatten();
    let format_png = png_bytes
        .as_ref()
        .map(|_| register_format("PNG"))
        .transpose()?;
    // Best-effort PNG variant used by some apps.
    let format_image_png = png_bytes.as_ref().and_then(|_| register_format("image/png").ok());

    let _guard = open_clipboard_with_retry()?;

    unsafe { EmptyClipboard() }.map_err(|e| win_err("EmptyClipboard failed", e))?;

    if let Some(text) = payload.text.as_deref() {
        set_unicode_text(text)?;
    }

    if let (Some(format), Some(bytes)) = (format_html, html_bytes.as_deref()) {
        let mut nul_terminated = bytes.to_vec();
        nul_terminated.push(0);
        set_clipboard_bytes(format, &nul_terminated)?;
    }
    if let (Some(format), Some(bytes)) = (format_text_html, html_plain_bytes.as_deref()) {
        let mut nul_terminated = bytes.to_vec();
        nul_terminated.push(0);
        let _ = set_clipboard_bytes(format, &nul_terminated);
    }

    if let (Some(format), Some(bytes)) = (format_rtf, rtf_bytes.as_deref()) {
        let mut nul_terminated = bytes.to_vec();
        nul_terminated.push(0);
        set_clipboard_bytes(format, &nul_terminated)?;
    }
    if let (Some(format), Some(bytes)) = (format_text_rtf, rtf_bytes.as_deref()) {
        let mut nul_terminated = bytes.to_vec();
        nul_terminated.push(0);
        let _ = set_clipboard_bytes(format, &nul_terminated);
    }

    if let (Some(format), Some(bytes)) = (format_png, png_bytes.as_deref()) {
        set_clipboard_bytes(format, bytes)?;
    }
    if let (Some(format), Some(bytes)) = (format_image_png, png_bytes.as_deref()) {
        // Best-effort; don't fail the entire clipboard write if this alias can't be set.
        let _ = set_clipboard_bytes(format, bytes);
    }

    // Interop: also provide DIB representations.
    //
    // - CF_DIBV5 supports alpha and modern headers.
    // - CF_DIB is more widely supported but may ignore alpha; we provide an opaque buffer to avoid
    //   consumers accidentally treating the 4th byte as transparent alpha.
    if let Some(bytes) = dibv5_bytes.as_deref() {
        set_clipboard_bytes(CF_DIBV5, bytes)?;
    }
    if let Some(bytes) = dib_bytes.as_deref() {
        // Best-effort: don't fail the entire clipboard write if the legacy format can't be set.
        let _ = set_clipboard_bytes(CF_DIB, bytes);
    }

    Ok(())
}
