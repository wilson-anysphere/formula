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
use super::windows_dib::{dibv5_to_png, png_to_dibv5};
use super::{
    normalize_base64_str, ClipboardContent, ClipboardError, ClipboardWritePayload, MAX_IMAGE_BYTES,
    MAX_RICH_TEXT_BYTES,
};

// Built-in clipboard formats that we use directly. Keeping these as numeric constants avoids
// needing Win32 System Ole bindings just for format IDs.
const CF_UNICODETEXT: u32 = 13;
const CF_DIB: u32 = 8;
const CF_DIBV5: u32 = 17;

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
            Err(err) if attempt + 1 < ATTEMPTS => {
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
    let mut out = slice.to_vec();

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
        Ok(Some(text))
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

pub fn read() -> Result<ClipboardContent, ClipboardError> {
    let format_html = register_format("HTML Format")?;
    let format_rtf = register_format("Rich Text Format")?;
    let format_png = register_format("PNG")?;
    // Some producers use a MIME-like name for PNG. Treat this as best-effort fallback.
    let format_image_png = register_format("image/png").ok();
    // MIME-like aliases used by some cross-platform apps.
    let format_text_html = register_format("text/html").ok();
    let format_text_rtf = register_format("text/rtf").ok();

    let _guard = open_clipboard_with_retry()?;

    // Best-effort reads: don't fail the entire operation if a single format can't be decoded.
    let text = try_get_unicode_text().ok().flatten();

    let mut html = try_get_clipboard_bytes(format_html, MAX_RICH_TEXT_BYTES)
        .ok()
        .flatten()
        .map(|bytes| extract_cf_html_fragment_best_effort(&bytes))
        .filter(|s| !s.is_empty());
    if html.is_none() {
        if let Some(format) = format_text_html {
            html = try_get_clipboard_bytes(format, MAX_RICH_TEXT_BYTES)
                .ok()
                .flatten()
                .map(|bytes| extract_cf_html_fragment_best_effort(&bytes))
                .filter(|s| !s.is_empty());
        }
    }

    let mut rtf = try_get_clipboard_bytes(format_rtf, MAX_RICH_TEXT_BYTES)
        .ok()
        .flatten()
        .map(|bytes| {
            String::from_utf8_lossy(&bytes)
                .trim_end_matches('\0')
                .to_string()
        })
        .filter(|s| !s.is_empty());
    if rtf.is_none() {
        if let Some(format) = format_text_rtf {
            rtf = try_get_clipboard_bytes(format, MAX_RICH_TEXT_BYTES)
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

    const MAX_DIB_BYTES: usize = 4 * MAX_IMAGE_BYTES; // allow larger uncompressed DIBs before converting to PNG

    let mut png_base64 = if let Some(png_bytes) =
        try_get_clipboard_bytes(format_png, MAX_IMAGE_BYTES).ok().flatten()
    {
        Some(STANDARD.encode(png_bytes))
    } else if let Some(format) = format_image_png {
        try_get_clipboard_bytes(format, MAX_IMAGE_BYTES)
            .ok()
            .flatten()
            .map(|bytes| STANDARD.encode(bytes))
    } else if let Some(dib_bytes) = try_get_clipboard_bytes(CF_DIBV5, MAX_DIB_BYTES).ok().flatten() {
        dibv5_to_png(&dib_bytes)
            .ok()
            .filter(|png| png.len() <= MAX_IMAGE_BYTES)
            .map(|png| STANDARD.encode(png))
    } else {
        None
    };

    if png_base64.is_none() {
        if let Some(dib_bytes) = try_get_clipboard_bytes(CF_DIB, MAX_DIB_BYTES).ok().flatten() {
            png_base64 = dibv5_to_png(&dib_bytes)
                .ok()
                .filter(|png| png.len() <= MAX_IMAGE_BYTES)
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

    let dib_bytes = png_bytes
        .as_deref()
        .map(|png| png_to_dibv5(png).map_err(|e| ClipboardError::InvalidPayload(e)))
        .transpose()?;

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

    // Interop: also provide a DIBV5 representation.
    if let Some(bytes) = dib_bytes.as_deref() {
        set_clipboard_bytes(CF_DIBV5, bytes)?;
        // Additional interop: provide the older CF_DIB flavor too. Many consumers accept a
        // BITMAPV5HEADER payload even when requesting CF_DIB.
        let _ = set_clipboard_bytes(CF_DIB, bytes);
    }

    Ok(())
}
