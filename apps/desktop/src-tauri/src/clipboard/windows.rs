use std::time::Duration;

use base64::{engine::general_purpose::STANDARD, Engine as _};
use windows::core::PCWSTR;
use windows::Win32::Foundation::{GlobalFree, HANDLE, HGLOBAL, HWND};
use windows::Win32::Globalization::{
    MultiByteToWideChar, CP_ACP, CP_OEMCP, MB_ERR_INVALID_CHARS, MULTI_BYTE_TO_WIDE_CHAR_FLAGS,
};
use windows::Win32::System::DataExchange::{
    CloseClipboard, EmptyClipboard, GetClipboardData, OpenClipboard, RegisterClipboardFormatW,
    SetClipboardData,
};
use windows::Win32::System::Memory::{
    GlobalAlloc, GlobalLock, GlobalSize, GlobalUnlock, GMEM_MOVEABLE,
};
use windows::Win32::UI::WindowsAndMessaging::{GetOpenClipboardWindow, GetWindowThreadProcessId};

use super::cf_html::{build_cf_html_payload, extract_cf_html_fragment_best_effort};
use super::retry::{retry_with_delays_if, total_delay, OPEN_CLIPBOARD_RETRY_DELAYS};
use super::windows_dib::{dibv5_to_png, png_to_dib_and_dibv5};
use super::windows_format_cache::CachedClipboardFormat;
use super::{
    debug_clipboard_log,
    normalize_base64_str, string_within_limit, ClipboardContent, ClipboardError, ClipboardWritePayload,
    MAX_PNG_BYTES, MAX_TEXT_BYTES,
};

// Built-in clipboard formats that we use directly. Keeping these as numeric constants avoids
// needing Win32 System Ole bindings just for format IDs.
const CF_TEXT: u32 = 1;
const CF_UNICODETEXT: u32 = 13;
const CF_OEMTEXT: u32 = 7;
const CF_DIB: u32 = 8;
const CF_DIBV5: u32 = 17;

// DIB payloads are uncompressed, so they can be significantly larger than the corresponding PNG.
// Allow a larger cap for DIB formats so we can still convert many clipboard images to PNG.
const MAX_DIB_BYTES: usize = 4 * MAX_PNG_BYTES;

// Custom clipboard formats we support. These are registered lazily and cached to keep clipboard
// reads/writes on hot paths allocation-free.
static FORMAT_HTML: CachedClipboardFormat = CachedClipboardFormat::new("HTML Format");
static FORMAT_RTF: CachedClipboardFormat = CachedClipboardFormat::new("Rich Text Format");
static FORMAT_PNG: CachedClipboardFormat = CachedClipboardFormat::new("PNG");
static FORMAT_IMAGE_PNG: CachedClipboardFormat = CachedClipboardFormat::new("image/png");
static FORMAT_IMAGE_X_PNG: CachedClipboardFormat = CachedClipboardFormat::new("image/x-png");
static FORMAT_TEXT_HTML: CachedClipboardFormat = CachedClipboardFormat::new("text/html");
static FORMAT_TEXT_HTML_UTF8: CachedClipboardFormat =
    CachedClipboardFormat::new("text/html;charset=utf-8");
static FORMAT_TEXT_RTF: CachedClipboardFormat = CachedClipboardFormat::new("text/rtf");
static FORMAT_TEXT_RTF_UTF8: CachedClipboardFormat =
    CachedClipboardFormat::new("text/rtf;charset=utf-8");
static FORMAT_APPLICATION_RTF: CachedClipboardFormat =
    CachedClipboardFormat::new("application/rtf");
static FORMAT_APPLICATION_X_RTF: CachedClipboardFormat =
    CachedClipboardFormat::new("application/x-rtf");

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

fn is_retryable_open_clipboard_error(err: &windows::core::Error) -> bool {
    // Per Win32 docs, `OpenClipboard` commonly fails when another process is holding the clipboard
    // lock. The windows crate reports these as HRESULTs (typically `HRESULT_FROM_WIN32`).
    //
    // In practice we most commonly see:
    // - `ERROR_ACCESS_DENIED` (5) -> 0x80070005
    // - `ERROR_BUSY` (170)        -> 0x800700AA
    //
    // Some producers/layers may surface OLE clipboard HRESULTs; treat `CLIPBRD_E_CANT_OPEN` as
    // retryable as well since it often indicates the clipboard is temporarily unavailable.
    const E_ACCESSDENIED: windows::core::HRESULT = windows::core::HRESULT(0x80070005u32 as i32);
    const E_BUSY: windows::core::HRESULT = windows::core::HRESULT(0x800700AAu32 as i32);
    const CLIPBRD_E_CANT_OPEN: windows::core::HRESULT =
        windows::core::HRESULT(0x800401D0u32 as i32);

    let code = err.code();
    code == E_ACCESSDENIED || code == E_BUSY || code == CLIPBRD_E_CANT_OPEN
}

fn get_open_clipboard_lock_holder() -> Option<(HWND, u32, u32)> {
    unsafe {
        let hwnd = GetOpenClipboardWindow();
        if hwnd.0 == 0 {
            return None;
        }
        let mut pid = 0u32;
        let tid = GetWindowThreadProcessId(hwnd, Some(&mut pid));
        Some((hwnd, pid, tid))
    }
}

fn open_clipboard_with_retry() -> Result<ClipboardGuard, ClipboardError> {
    // The clipboard is a shared global resource. `OpenClipboard` can fail temporarily when another
    // process is holding the clipboard lock. Use a deterministic exponential backoff (bounded total
    // sleep budget) to reduce flakiness under contention.
    let max_attempts = OPEN_CLIPBOARD_RETRY_DELAYS.len() + 1;
    let max_total_sleep = total_delay(OPEN_CLIPBOARD_RETRY_DELAYS);
    let mut attempts = 0usize;
    let mut slept = Duration::from_millis(0);
    let mut lock_holder = None::<(HWND, u32, u32)>;

    retry_with_delays_if(
        || {
            attempts += 1;
            unsafe { OpenClipboard(None) }
        },
        OPEN_CLIPBOARD_RETRY_DELAYS,
        is_retryable_open_clipboard_error,
        |d| {
            // Best-effort: try to capture the window + PID that currently has the clipboard open so
            // we can include it in the final error for debugging. This runs only on retryable
            // failures, not on the success path.
            if let Some(holder) = get_open_clipboard_lock_holder() {
                lock_holder = Some(holder);
            }
            slept += d;
            std::thread::sleep(d);
        },
    )
    .map(|()| ClipboardGuard)
    .map_err(|err| {
        let last_hresult = err.code().0 as u32;
        let retriable = is_retryable_open_clipboard_error(&err);
        let lock_holder_ctx = lock_holder
            .map(|(hwnd, pid, tid)| {
                format!(
                    "open_clipboard_window=0x{:X}, open_clipboard_pid={pid}, open_clipboard_tid={tid}",
                    hwnd.0 as usize
                )
            })
            .unwrap_or_else(|| "open_clipboard_window=None".to_string());
        win_err(
            &format!(
                "OpenClipboard failed after {attempts}/{max_attempts} attempts over {}ms sleep (budget {}ms, retriable={retriable}, last_hresult=0x{last_hresult:08X}, {lock_holder_ctx})",
                slept.as_millis(),
                max_total_sleep.as_millis(),
            ),
            err,
        )
    })
}

fn register_format(name: &'static str) -> Option<u32> {
    let mut wide: Vec<u16> = name.encode_utf16().collect();
    wide.push(0);
    let id = unsafe { RegisterClipboardFormatW(PCWSTR(wide.as_ptr())) };
    (id != 0).then_some(id)
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

fn try_get_ansi_text(format: u32, code_page: u32) -> Result<Option<String>, ClipboardError> {
    let bytes = match try_get_clipboard_bytes(format, MAX_TEXT_BYTES)? {
        Some(bytes) => bytes,
        None => return Ok(None),
    };

    let nul = bytes.iter().position(|&b| b == 0).unwrap_or(bytes.len());
    let bytes = &bytes[..nul];
    if bytes.is_empty() {
        return Ok(None);
    }

    unsafe {
        // `MultiByteToWideChar` returns 0 on error.
        let mut flags = MB_ERR_INVALID_CHARS;
        let mut wide_len = MultiByteToWideChar(code_page, flags, bytes, None);
        if wide_len == 0 {
            // Be permissive: fall back to lossy conversion when the input contains invalid bytes.
            flags = MULTI_BYTE_TO_WIDE_CHAR_FLAGS(0);
            wide_len = MultiByteToWideChar(code_page, flags, bytes, None);
        }
        if wide_len <= 0 {
            return Ok(None);
        }

        let wide_len = wide_len as usize;
        let mut wide = vec![0u16; wide_len];
        let written = MultiByteToWideChar(code_page, flags, bytes, Some(&mut wide));
        if written <= 0 {
            return Ok(None);
        }
        wide.truncate(written as usize);

        let text = String::from_utf16_lossy(&wide);
        if text.is_empty() {
            Ok(None)
        } else {
            Ok(string_within_limit(text, MAX_TEXT_BYTES))
        }
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

    let ihdr_len =
        u32::from_be_bytes([png_bytes[8], png_bytes[9], png_bytes[10], png_bytes[11]]) as usize;
    if ihdr_len < 8 {
        return None;
    }
    if png_bytes.get(12..16)? != b"IHDR" {
        return None;
    }

    let data_start = 16;
    let data_end = data_start.checked_add(ihdr_len)?;
    let data = png_bytes.get(data_start..data_end)?;
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
    let format_html = FORMAT_HTML.get_with(register_format);
    let format_rtf = FORMAT_RTF.get_with(register_format);
    let format_png = FORMAT_PNG.get_with(register_format);
    // Some producers use a MIME-like name for PNG. Treat this as best-effort fallback.
    let format_image_png = FORMAT_IMAGE_PNG.get_with(register_format);
    let format_image_x_png = FORMAT_IMAGE_X_PNG.get_with(register_format);
    // MIME-like aliases used by some cross-platform apps.
    let format_text_html = FORMAT_TEXT_HTML.get_with(register_format);
    let format_text_html_utf8 = FORMAT_TEXT_HTML_UTF8.get_with(register_format);
    let format_text_rtf = FORMAT_TEXT_RTF.get_with(register_format);
    let format_text_rtf_utf8 = FORMAT_TEXT_RTF_UTF8.get_with(register_format);
    let format_application_rtf = FORMAT_APPLICATION_RTF.get_with(register_format);
    let format_application_x_rtf = FORMAT_APPLICATION_X_RTF.get_with(register_format);

    let _guard = open_clipboard_with_retry()?;

    // Best-effort reads: don't fail the entire operation if a single format can't be decoded.
    let mut text_source: Option<&'static str> = None;
    let mut text = match try_get_unicode_text() {
        Ok(Some(s)) => {
            text_source = Some("CF_UNICODETEXT");
            Some(s)
        }
        _ => None,
    };
    if text.is_none() {
        if let Ok(Some(s)) = try_get_ansi_text(CF_TEXT, CP_ACP) {
            text_source = Some("CF_TEXT");
            text = Some(s);
        }
    }
    if text.is_none() {
        if let Ok(Some(s)) = try_get_ansi_text(CF_OEMTEXT, CP_OEMCP) {
            text_source = Some("CF_OEMTEXT");
            text = Some(s);
        }
    }

    let mut html_source: Option<&'static str> = None;
    let mut html = None;
    if let Some(format) = format_html {
        if let Ok(Some(bytes)) = try_get_clipboard_bytes(format, MAX_TEXT_BYTES) {
            let fragment = extract_cf_html_fragment_best_effort(&bytes);
            if !fragment.is_empty() {
                html_source = Some("HTML Format");
                html = Some(fragment);
            }
        }
    }
    if html.is_none() {
        if let Some(format) = format_text_html {
            if let Ok(Some(bytes)) = try_get_clipboard_bytes(format, MAX_TEXT_BYTES) {
                let fragment = extract_cf_html_fragment_best_effort(&bytes);
                if !fragment.is_empty() {
                    html_source = Some("text/html");
                    html = Some(fragment);
                }
            }
        }
    }
    if html.is_none() {
        if let Some(format) = format_text_html_utf8 {
            if let Ok(Some(bytes)) = try_get_clipboard_bytes(format, MAX_TEXT_BYTES) {
                let fragment = extract_cf_html_fragment_best_effort(&bytes);
                if !fragment.is_empty() {
                    html_source = Some("text/html;charset=utf-8");
                    html = Some(fragment);
                }
            }
        }
    }
    let html = html.and_then(|s| string_within_limit(s, MAX_TEXT_BYTES));
    if html.is_none() {
        html_source = None;
    }

    let mut rtf_source: Option<&'static str> = None;
    let mut rtf = None;
    if let Some(format) = format_rtf {
        if let Ok(Some(bytes)) = try_get_clipboard_bytes(format, MAX_TEXT_BYTES) {
            let s = String::from_utf8_lossy(&bytes)
                .trim_end_matches('\0')
                .to_string();
            if !s.is_empty() {
                rtf_source = Some("Rich Text Format");
                rtf = Some(s);
            }
        }
    }
    if rtf.is_none() {
        if let Some(format) = format_text_rtf {
            if let Ok(Some(bytes)) = try_get_clipboard_bytes(format, MAX_TEXT_BYTES) {
                let s = String::from_utf8_lossy(&bytes)
                    .trim_end_matches('\0')
                    .to_string();
                if !s.is_empty() {
                    rtf_source = Some("text/rtf");
                    rtf = Some(s);
                }
            }
        }
    }
    if rtf.is_none() {
        if let Some(format) = format_text_rtf_utf8 {
            if let Ok(Some(bytes)) = try_get_clipboard_bytes(format, MAX_TEXT_BYTES) {
                let s = String::from_utf8_lossy(&bytes)
                    .trim_end_matches('\0')
                    .to_string();
                if !s.is_empty() {
                    rtf_source = Some("text/rtf;charset=utf-8");
                    rtf = Some(s);
                }
            }
        }
    }
    if rtf.is_none() {
        if let Some(format) = format_application_rtf {
            if let Ok(Some(bytes)) = try_get_clipboard_bytes(format, MAX_TEXT_BYTES) {
                let s = String::from_utf8_lossy(&bytes)
                    .trim_end_matches('\0')
                    .to_string();
                if !s.is_empty() {
                    rtf_source = Some("application/rtf");
                    rtf = Some(s);
                }
            }
        }
    }
    if rtf.is_none() {
        if let Some(format) = format_application_x_rtf {
            if let Ok(Some(bytes)) = try_get_clipboard_bytes(format, MAX_TEXT_BYTES) {
                let s = String::from_utf8_lossy(&bytes)
                    .trim_end_matches('\0')
                    .to_string();
                if !s.is_empty() {
                    rtf_source = Some("application/x-rtf");
                    rtf = Some(s);
                }
            }
        }
    }
    let rtf = rtf.and_then(|s| string_within_limit(s, MAX_TEXT_BYTES));
    if rtf.is_none() {
        rtf_source = None;
    }

    // Images can be large; avoid doing expensive conversions/encoding while holding the clipboard
    // lock. First copy the raw bytes out of the clipboard, then close it before decoding or
    // base64-encoding.
    enum RawImage {
        Png { source: &'static str, bytes: Vec<u8> },
        Dib { source: &'static str, bytes: Vec<u8> },
    }

    let mut raw_image: Option<RawImage> = None;
    if let Some(format) = format_png {
        if let Ok(Some(bytes)) = try_get_clipboard_bytes(format, MAX_PNG_BYTES) {
            raw_image = Some(RawImage::Png {
                source: "PNG",
                bytes,
            });
        }
    }
    if raw_image.is_none() {
        if let Some(format) = format_image_png {
            if let Ok(Some(bytes)) = try_get_clipboard_bytes(format, MAX_PNG_BYTES) {
                raw_image = Some(RawImage::Png {
                    source: "image/png",
                    bytes,
                });
            }
        }
    }
    if raw_image.is_none() {
        if let Some(format) = format_image_x_png {
            if let Ok(Some(bytes)) = try_get_clipboard_bytes(format, MAX_PNG_BYTES) {
                raw_image = Some(RawImage::Png {
                    source: "image/x-png",
                    bytes,
                });
            }
        }
    }
    if raw_image.is_none() {
        if let Ok(Some(bytes)) = try_get_clipboard_bytes(CF_DIBV5, MAX_DIB_BYTES) {
            raw_image = Some(RawImage::Dib {
                source: "CF_DIBV5",
                bytes,
            });
        }
    }
    if raw_image.is_none() {
        if let Ok(Some(bytes)) = try_get_clipboard_bytes(CF_DIB, MAX_DIB_BYTES) {
            raw_image = Some(RawImage::Dib {
                source: "CF_DIB",
                bytes,
            });
        }
    }

    drop(_guard);

    let mut image_source: Option<&'static str> = None;
    let mut image_bytes: Option<usize> = None;
    let mut image_png_base64 = None;
    match raw_image {
        Some(RawImage::Png { source, bytes }) => {
            image_source = Some(source);
            image_bytes = Some(bytes.len());
            image_png_base64 = Some(STANDARD.encode(&bytes));
        }
        Some(RawImage::Dib { source, bytes }) => {
            if let Ok(png) = dibv5_to_png(&bytes) {
                if png.len() <= MAX_PNG_BYTES {
                    image_source = Some(match source {
                        "CF_DIBV5" => "CF_DIBV5->PNG",
                        "CF_DIB" => "CF_DIB->PNG",
                        _ => "CF_DIB->PNG",
                    });
                    image_bytes = Some(png.len());
                    image_png_base64 = Some(STANDARD.encode(&png));
                }
            }
        }
        None => {}
    }

    let text_bytes = text.as_ref().map(|s| s.as_bytes().len());
    let html_bytes = html.as_ref().map(|s| s.as_bytes().len());
    let rtf_bytes = rtf.as_ref().map(|s| s.as_bytes().len());
    debug_clipboard_log(format_args!(
        "windows read: text_source={text_source:?} text_bytes={text_bytes:?} html_source={html_source:?} html_bytes={html_bytes:?} rtf_source={rtf_source:?} rtf_bytes={rtf_bytes:?} image_source={image_source:?} image_bytes={image_bytes:?} caps(text={MAX_TEXT_BYTES}, png={MAX_PNG_BYTES})"
    ));

    Ok(ClipboardContent {
        text,
        html,
        rtf,
        image_png_base64,
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
        .image_png_base64
        .as_deref()
        .map(normalize_base64_str)
        .filter(|s| !s.is_empty())
        .map(|s| {
            STANDARD
                .decode(s)
                .map_err(|e| ClipboardError::InvalidPayload(format!("invalid png base64: {e}")))
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

    let text_len = payload.text.as_ref().map(|s| s.as_bytes().len());
    let cf_html_len = html_bytes.as_ref().map(|b| b.len());
    let text_html_len = html_plain_bytes.as_ref().map(|b| b.len());
    let rtf_len = rtf_bytes.as_ref().map(|b| b.len());
    let png_len = png_bytes.as_ref().map(|b| b.len());
    let dib_len = dib_bytes.as_ref().map(|b| b.len());
    let dibv5_len = dibv5_bytes.as_ref().map(|b| b.len());
    debug_clipboard_log(format_args!(
        "windows write: text_bytes={text_len:?} cf_html_bytes={cf_html_len:?} text_html_bytes={text_html_len:?} rtf_bytes={rtf_len:?} png_bytes={png_len:?} dibv5_bytes={dibv5_len:?} dib_bytes={dib_len:?} caps(text={MAX_TEXT_BYTES}, png={MAX_PNG_BYTES})"
    ));

    let format_html = html_bytes
        .as_ref()
        .and_then(|_| FORMAT_HTML.get_with(register_format));
    let format_text_html = html_plain_bytes
        .as_ref()
        .and_then(|_| FORMAT_TEXT_HTML.get_with(register_format));
    let format_text_html_utf8 = html_plain_bytes
        .as_ref()
        .and_then(|_| FORMAT_TEXT_HTML_UTF8.get_with(register_format));
    let format_rtf = rtf_bytes
        .as_ref()
        .and_then(|_| FORMAT_RTF.get_with(register_format));
    let format_text_rtf = rtf_bytes
        .as_ref()
        .and_then(|_| FORMAT_TEXT_RTF.get_with(register_format));
    let format_text_rtf_utf8 = rtf_bytes
        .as_ref()
        .and_then(|_| FORMAT_TEXT_RTF_UTF8.get_with(register_format));
    let format_application_rtf = rtf_bytes
        .as_ref()
        .and_then(|_| FORMAT_APPLICATION_RTF.get_with(register_format));
    let format_application_x_rtf = rtf_bytes
        .as_ref()
        .and_then(|_| FORMAT_APPLICATION_X_RTF.get_with(register_format));
    let format_png = png_bytes
        .as_ref()
        .and_then(|_| FORMAT_PNG.get_with(register_format));
    // Best-effort PNG variant used by some apps.
    let format_image_png = png_bytes
        .as_ref()
        .and_then(|_| FORMAT_IMAGE_PNG.get_with(register_format));
    let format_image_x_png = png_bytes
        .as_ref()
        .and_then(|_| FORMAT_IMAGE_X_PNG.get_with(register_format));

    // If we can't set any representation (e.g. clipboard format registration failed), don't clear
    // the existing clipboard contents.
    let requested_any = payload.text.is_some()
        || html_bytes.is_some()
        || rtf_bytes.is_some()
        || png_bytes.is_some();
    let can_set_any = payload.text.is_some()
        || format_html.is_some()
        || format_text_html.is_some()
        || format_text_html_utf8.is_some()
        || format_rtf.is_some()
        || format_text_rtf.is_some()
        || format_text_rtf_utf8.is_some()
        || format_application_rtf.is_some()
        || format_application_x_rtf.is_some()
        || format_png.is_some()
        || format_image_png.is_some()
        || format_image_x_png.is_some()
        || dibv5_bytes.is_some()
        || dib_bytes.is_some();
    if requested_any && !can_set_any {
        return Ok(());
    }

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
    if let (Some(format), Some(bytes)) = (format_text_html_utf8, html_plain_bytes.as_deref()) {
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
    if let (Some(format), Some(bytes)) = (format_text_rtf_utf8, rtf_bytes.as_deref()) {
        let mut nul_terminated = bytes.to_vec();
        nul_terminated.push(0);
        let _ = set_clipboard_bytes(format, &nul_terminated);
    }
    if let (Some(format), Some(bytes)) = (format_application_rtf, rtf_bytes.as_deref()) {
        let mut nul_terminated = bytes.to_vec();
        nul_terminated.push(0);
        let _ = set_clipboard_bytes(format, &nul_terminated);
    }
    if let (Some(format), Some(bytes)) = (format_application_x_rtf, rtf_bytes.as_deref()) {
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
    if let (Some(format), Some(bytes)) = (format_image_x_png, png_bytes.as_deref()) {
        // Best-effort; don't fail the entire clipboard write if this alias can't be set.
        let _ = set_clipboard_bytes(format, bytes);
    }

    // Interop: also provide DIB representations.
    //
    // - CF_DIBV5 supports alpha and modern headers.
    // - CF_DIB is more widely supported but may ignore alpha; we provide an opaque buffer to avoid
    //   consumers accidentally treating the 4th byte as transparent alpha.
    if let Some(bytes) = dibv5_bytes.as_deref() {
        // Best-effort: clipboard image data is already available via the registered PNG format, so
        // failure to set the DIBV5 representation shouldn't fail the whole write.
        let _ = set_clipboard_bytes(CF_DIBV5, bytes);
    }
    if let Some(bytes) = dib_bytes.as_deref() {
        // Best-effort: don't fail the entire clipboard write if the legacy format can't be set.
        let _ = set_clipboard_bytes(CF_DIB, bytes);
    }

    Ok(())
}
