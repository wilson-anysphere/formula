use super::super::{ClipboardContent, ClipboardError, ClipboardWritePayload};

#[cfg(feature = "desktop")]
mod win32_backend {
    use base64::{engine::general_purpose::STANDARD, Engine as _};
    use std::ffi::OsStr;
    use std::os::windows::ffi::OsStrExt;
    use std::ptr;
    use std::thread;
    use std::time::Duration;

    use windows_sys::Win32::Foundation::GetLastError;
    use windows_sys::Win32::System::DataExchange::{
        CloseClipboard, EmptyClipboard, GetClipboardData, IsClipboardFormatAvailable, OpenClipboard,
        RegisterClipboardFormatW, SetClipboardData, CF_UNICODETEXT,
    };
    use windows_sys::Win32::System::Memory::{
        GlobalAlloc, GlobalFree, GlobalLock, GlobalSize, GlobalUnlock, GMEM_MOVEABLE,
    };

    use crate::clipboard::{cf_html, ClipboardContent, ClipboardError, ClipboardWritePayload};

    fn win32_error(context: &str) -> ClipboardError {
        let code = unsafe { GetLastError() };
        ClipboardError::OperationFailed(format!("{context} failed (win32 error {code})"))
    }

    struct ClipboardGuard;

    impl Drop for ClipboardGuard {
        fn drop(&mut self) {
            unsafe {
                CloseClipboard();
            }
        }
    }

    fn open_clipboard_with_retries() -> Result<ClipboardGuard, ClipboardError> {
        // The clipboard is a global resource and can be temporarily locked by other processes.
        // Retrying for a short period avoids flaky failures, especially during rapid copy/paste.
        const ATTEMPTS: usize = 10;
        for attempt in 0..ATTEMPTS {
            let ok = unsafe { OpenClipboard(0) };
            if ok != 0 {
                return Ok(ClipboardGuard);
            }
            if attempt + 1 < ATTEMPTS {
                thread::sleep(Duration::from_millis(10));
            }
        }
        Err(win32_error("OpenClipboard"))
    }

    fn wide_null_terminated(s: &str) -> Vec<u16> {
        OsStr::new(s).encode_wide().chain(Some(0)).collect()
    }

    fn register_format(name: &str) -> Result<u32, ClipboardError> {
        let wide = wide_null_terminated(name);
        let id = unsafe { RegisterClipboardFormatW(wide.as_ptr()) };
        if id == 0 {
            Err(win32_error("RegisterClipboardFormatW"))
        } else {
            Ok(id)
        }
    }

    fn read_unicode_text() -> Result<Option<String>, ClipboardError> {
        let available = unsafe { IsClipboardFormatAvailable(CF_UNICODETEXT) };
        if available == 0 {
            return Ok(None);
        }
        let handle = unsafe { GetClipboardData(CF_UNICODETEXT) };
        if handle == 0 {
            return Ok(None);
        }

        let locked = unsafe { GlobalLock(handle) } as *const u16;
        if locked.is_null() {
            return Err(win32_error("GlobalLock(CF_UNICODETEXT)"));
        }

        let size_bytes = unsafe { GlobalSize(handle) };
        let slice = unsafe { std::slice::from_raw_parts(locked, size_bytes / 2) };
        let len = slice.iter().position(|&c| c == 0).unwrap_or(slice.len());
        let text = String::from_utf16_lossy(&slice[..len]);

        unsafe {
            GlobalUnlock(handle);
        }

        if text.is_empty() {
            Ok(None)
        } else {
            Ok(Some(text))
        }
    }

    fn read_clipboard_bytes(format: u32, label: &str) -> Result<Option<Vec<u8>>, ClipboardError> {
        let available = unsafe { IsClipboardFormatAvailable(format) };
        if available == 0 {
            return Ok(None);
        }

        let handle = unsafe { GetClipboardData(format) };
        if handle == 0 {
            return Ok(None);
        }

        let locked = unsafe { GlobalLock(handle) } as *const u8;
        if locked.is_null() {
            return Err(win32_error(&format!("GlobalLock({label})")));
        }

        let size = unsafe { GlobalSize(handle) };
        let out = if size == 0 {
            Vec::new()
        } else {
            unsafe { std::slice::from_raw_parts(locked, size) }.to_vec()
        };

        unsafe {
            GlobalUnlock(handle);
        }

        if out.is_empty() {
            Ok(None)
        } else {
            Ok(Some(out))
        }
    }

    fn read_clipboard_string(format: u32, label: &str) -> Result<Option<String>, ClipboardError> {
        let Some(bytes) = read_clipboard_bytes(format, label)? else {
            return Ok(None);
        };
        let s = String::from_utf8_lossy(&bytes).trim_end_matches('\0').to_string();
        if s.is_empty() {
            Ok(None)
        } else {
            Ok(Some(s))
        }
    }

    fn set_clipboard_bytes(format: u32, bytes: &[u8], null_terminate: bool) -> Result<(), ClipboardError> {
        let mut owned;
        let data = if null_terminate {
            owned = Vec::with_capacity(bytes.len() + 1);
            owned.extend_from_slice(bytes);
            owned.push(0);
            owned.as_slice()
        } else {
            bytes
        };

        let handle = unsafe { GlobalAlloc(GMEM_MOVEABLE, data.len()) };
        if handle == 0 {
            return Err(win32_error("GlobalAlloc"));
        }

        let locked = unsafe { GlobalLock(handle) } as *mut u8;
        if locked.is_null() {
            unsafe {
                GlobalFree(handle);
            }
            return Err(win32_error("GlobalLock"));
        }

        unsafe {
            ptr::copy_nonoverlapping(data.as_ptr(), locked, data.len());
            GlobalUnlock(handle);
        }

        let res = unsafe { SetClipboardData(format, handle) };
        if res == 0 {
            unsafe {
                // The clipboard did not take ownership. Free our allocation.
                GlobalFree(handle);
            }
            return Err(win32_error("SetClipboardData"));
        }

        Ok(())
    }

    fn set_clipboard_unicode_text(text: &str) -> Result<(), ClipboardError> {
        // CF_UNICODETEXT requires UTF-16 with a terminating NUL.
        let mut wide: Vec<u16> = OsStr::new(text).encode_wide().collect();
        wide.push(0);

        let bytes_len = wide.len() * 2;
        let handle = unsafe { GlobalAlloc(GMEM_MOVEABLE, bytes_len) };
        if handle == 0 {
            return Err(win32_error("GlobalAlloc(CF_UNICODETEXT)"));
        }

        let locked = unsafe { GlobalLock(handle) } as *mut u16;
        if locked.is_null() {
            unsafe {
                GlobalFree(handle);
            }
            return Err(win32_error("GlobalLock(CF_UNICODETEXT)"));
        }

        unsafe {
            ptr::copy_nonoverlapping(wide.as_ptr(), locked, wide.len());
            GlobalUnlock(handle);
        }

        let res = unsafe { SetClipboardData(CF_UNICODETEXT, handle) };
        if res == 0 {
            unsafe {
                GlobalFree(handle);
            }
            return Err(win32_error("SetClipboardData(CF_UNICODETEXT)"));
        }

        Ok(())
    }

    pub(super) fn read() -> Result<ClipboardContent, ClipboardError> {
        let _guard = open_clipboard_with_retries()?;

        // Best-effort reads: if a particular format fails to decode, keep going.
        let text = read_unicode_text().ok().flatten();

        let html = register_format("HTML Format")
            .ok()
            .and_then(|fmt| read_clipboard_string(fmt, "HTML Format").ok().flatten())
            .and_then(|payload| cf_html::decode_cf_html(&payload));

        let rtf = register_format("Rich Text Format")
            .ok()
            .and_then(|fmt| read_clipboard_string(fmt, "Rich Text Format").ok().flatten());

        // PNG is a registered clipboard format on Windows. Some producers use `PNG`, others use
        // `image/png`. Try both.
        let png_bytes = register_format("PNG")
            .ok()
            .and_then(|fmt| read_clipboard_bytes(fmt, "PNG").ok().flatten())
            .or_else(|| {
                register_format("image/png")
                    .ok()
                    .and_then(|fmt| read_clipboard_bytes(fmt, "image/png").ok().flatten())
            });

        let png_base64 = png_bytes
            .filter(|b| !b.is_empty())
            .map(|bytes| STANDARD.encode(bytes));

        Ok(ClipboardContent {
            text,
            html,
            rtf,
            png_base64,
        })
    }

    pub(super) fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
        let _guard = open_clipboard_with_retries()?;

        let ok = unsafe { EmptyClipboard() };
        if ok == 0 {
            return Err(win32_error("EmptyClipboard"));
        }

        if let Some(text) = payload.text.as_deref() {
            set_clipboard_unicode_text(text)?;
        }

        if let Some(html) = payload.html.as_deref() {
            let cf_html_bytes = cf_html::build_cf_html_payload(html)
                .map_err(|e| ClipboardError::InvalidPayload(format!("invalid html: {e}")))?;
            let fmt = register_format("HTML Format")?;
            set_clipboard_bytes(fmt, &cf_html_bytes, true)?;
        }

        if let Some(rtf) = payload.rtf.as_deref() {
            let fmt = register_format("Rich Text Format")?;
            set_clipboard_bytes(fmt, rtf.as_bytes(), true)?;
        }

        if let Some(base64) = payload.png_base64.as_deref() {
            if !base64.is_empty() {
                let bytes = STANDARD
                    .decode(base64)
                    .map_err(|e| ClipboardError::InvalidPayload(format!("invalid pngBase64: {e}")))?;

                // Prefer the canonical Windows clipboard format name.
                let fmt = register_format("PNG")?;
                set_clipboard_bytes(fmt, &bytes, false)?;

                // Best-effort compatibility for producers/consumers that use the MIME-like name.
                if let Ok(fmt) = register_format("image/png") {
                    let _ = set_clipboard_bytes(fmt, &bytes, false);
                }
            }
        }

        Ok(())
    }
}

#[cfg(feature = "desktop")]
pub fn read() -> Result<ClipboardContent, ClipboardError> {
    win32_backend::read()
}

#[cfg(feature = "desktop")]
pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    win32_backend::write(payload)
}

#[cfg(not(feature = "desktop"))]
pub fn read() -> Result<ClipboardContent, ClipboardError> {
    Err(ClipboardError::Unavailable(
        "Windows clipboard backend requires the `desktop` feature".to_string(),
    ))
}

#[cfg(not(feature = "desktop"))]
pub fn write(_payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    Err(ClipboardError::Unavailable(
        "Windows clipboard backend requires the `desktop` feature".to_string(),
    ))
}

