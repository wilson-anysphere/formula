//! macOS clipboard backend.
//!
//! We use `NSPasteboard` to read/write rich clipboard formats.
//!
//! ## Threading
//! AppKit is not thread-safe; clipboard calls must occur on the main thread. Tauri
//! commands may execute on a background thread, so the command wrapper should dispatch
//! to the main thread (e.g. `AppHandle::run_on_main_thread`) before calling into this
//! module. We also enforce this at runtime to avoid hard-to-debug crashes.

use base64::{engine::general_purpose::STANDARD, Engine as _};
use objc2::rc::{autoreleasepool, Id, Owned};
use objc2::runtime::AnyObject;

use std::ffi::{c_void, CStr};

use super::{
    ClipboardContent, ClipboardError, ClipboardWritePayload, MAX_IMAGE_BYTES, MAX_RICH_TEXT_BYTES,
};

// Ensure the framework crates are linked (and silence `unused_crate_dependencies`).
use objc2_app_kit as _;
use objc2_foundation as _;

// NSPasteboard type constants (UTType identifiers).
const TYPE_STRING: &str = "public.utf8-plain-text"; // NSPasteboardTypeString
const TYPE_HTML: &str = "public.html"; // NSPasteboardTypeHTML
const TYPE_RTF: &str = "public.rtf"; // NSPasteboardTypeRTF
const TYPE_PNG: &str = "public.png"; // NSPasteboardTypePNG

fn ensure_main_thread() -> Result<(), ClipboardError> {
    // +[NSThread isMainThread]
    let is_main: bool = unsafe { objc2::msg_send![objc2::class!(NSThread), isMainThread] };
    if is_main {
        Ok(())
    } else {
        Err(ClipboardError::OperationFailed(
            "clipboard operations must be performed on the main thread".to_string(),
        ))
    }
}

unsafe fn nsstring_from_str(s: &str) -> Result<Id<AnyObject, Owned>, ClipboardError> {
    // [[NSString alloc] initWithBytes:length:encoding:]
    //
    // We use the bytes+length initializer to avoid issues with embedded NULs
    // (CString-based APIs would truncate).
    let cls = objc2::class!(NSString);
    let alloc: *mut AnyObject = objc2::msg_send![cls, alloc];
    let bytes = s.as_bytes();
    let obj: *mut AnyObject = objc2::msg_send![
        alloc,
        initWithBytes: bytes.as_ptr()
        length: bytes.len()
        encoding: 4usize /* NSUTF8StringEncoding */
    ];
    if obj.is_null() {
        return Err(ClipboardError::OperationFailed(
            "failed to allocate NSString".to_string(),
        ));
    }
    Ok(Id::from_retained_ptr(obj))
}

unsafe fn nsdata_from_bytes(bytes: &[u8]) -> Result<Id<AnyObject, Owned>, ClipboardError> {
    // [[NSData alloc] initWithBytes:length:]
    let cls = objc2::class!(NSData);
    let alloc: *mut AnyObject = objc2::msg_send![cls, alloc];
    let obj: *mut AnyObject =
        objc2::msg_send![alloc, initWithBytes: bytes.as_ptr() length: bytes.len()];
    if obj.is_null() {
        return Err(ClipboardError::OperationFailed(
            "failed to allocate NSData".to_string(),
        ));
    }
    Ok(Id::from_retained_ptr(obj))
}

unsafe fn nsstring_to_rust_string(ns_str: *mut AnyObject) -> Option<String> {
    if ns_str.is_null() {
        return None;
    }
    let c_str: *const i8 = objc2::msg_send![ns_str, UTF8String];
    if c_str.is_null() {
        return None;
    }
    Some(CStr::from_ptr(c_str).to_string_lossy().into_owned())
}

unsafe fn pasteboard_string_for_type(pasteboard: *mut AnyObject, ty: &AnyObject) -> Option<String> {
    // -[NSPasteboard stringForType:]
    let ns_str: *mut AnyObject = objc2::msg_send![pasteboard, stringForType: ty];
    nsstring_to_rust_string(ns_str)
}

unsafe fn nsdata_to_vec(data: *mut AnyObject, max_bytes: usize) -> Vec<u8> {
    // -[NSData bytes], -[NSData length]
    let bytes_ptr: *const c_void = objc2::msg_send![data, bytes];
    let len: usize = objc2::msg_send![data, length];
    if bytes_ptr.is_null() || len == 0 || len > max_bytes {
        return Vec::new();
    }
    std::slice::from_raw_parts(bytes_ptr as *const u8, len).to_vec()
}

unsafe fn pasteboard_data_for_type(
    pasteboard: *mut AnyObject,
    ty: &AnyObject,
    max_bytes: usize,
) -> Option<Vec<u8>> {
    // -[NSPasteboard dataForType:]
    let data: *mut AnyObject = objc2::msg_send![pasteboard, dataForType: ty];
    if data.is_null() {
        return None;
    }
    let bytes = nsdata_to_vec(data, max_bytes);
    if bytes.is_empty() {
        None
    } else {
        Some(bytes)
    }
}

pub fn read() -> Result<ClipboardContent, ClipboardError> {
    ensure_main_thread()?;

    autoreleasepool(|_| unsafe {
        // +[NSPasteboard generalPasteboard]
        let pasteboard: *mut AnyObject =
            objc2::msg_send![objc2::class!(NSPasteboard), generalPasteboard];
        if pasteboard.is_null() {
            return Err(ClipboardError::OperationFailed(
                "NSPasteboard.generalPasteboard returned nil".to_string(),
            ));
        }

        let ty_string = nsstring_from_str(TYPE_STRING)?;
        let ty_html = nsstring_from_str(TYPE_HTML)?;
        let ty_rtf = nsstring_from_str(TYPE_RTF)?;
        let ty_png = nsstring_from_str(TYPE_PNG)?;

        let text = pasteboard_string_for_type(pasteboard, &*ty_string).or_else(|| {
            pasteboard_data_for_type(pasteboard, &*ty_string, usize::MAX)
                .map(|bytes| String::from_utf8_lossy(&bytes).into_owned())
        });

        let html = pasteboard_string_for_type(pasteboard, &*ty_html).or_else(|| {
            pasteboard_data_for_type(pasteboard, &*ty_html, MAX_RICH_TEXT_BYTES)
                .map(|bytes| String::from_utf8_lossy(&bytes).into_owned())
        });

        let rtf = pasteboard_data_for_type(pasteboard, &*ty_rtf, MAX_RICH_TEXT_BYTES)
            .map(|bytes| String::from_utf8_lossy(&bytes).into_owned());

        let png_base64 = pasteboard_data_for_type(pasteboard, &*ty_png, MAX_IMAGE_BYTES)
            .filter(|bytes| !bytes.is_empty())
            .map(|bytes| STANDARD.encode(bytes));

        Ok(ClipboardContent {
            text,
            html,
            rtf,
            png_base64,
        })
    })
}

pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    ensure_main_thread()?;

    let png_bytes = payload
        .png_base64
        .as_deref()
        .filter(|s| !s.is_empty())
        .map(|s| {
            STANDARD
                .decode(s)
                .map_err(|e| ClipboardError::InvalidPayload(format!("invalid pngBase64: {e}")))
        })
        .transpose()?;

    autoreleasepool(|_| unsafe {
        // +[NSPasteboard generalPasteboard]
        let pasteboard: *mut AnyObject =
            objc2::msg_send![objc2::class!(NSPasteboard), generalPasteboard];
        if pasteboard.is_null() {
            return Err(ClipboardError::OperationFailed(
                "NSPasteboard.generalPasteboard returned nil".to_string(),
            ));
        }

        // Clear first so we replace the current clipboard item(s).
        let _: isize = objc2::msg_send![pasteboard, clearContents];

        // Multi-format write:
        // Writing via `-[NSPasteboard setString:forType:]` can clobber previous
        // representations. Use one `NSPasteboardItem` and publish it via
        // `-[NSPasteboard writeObjects:]` so all formats travel together.
        let item_ptr: *mut AnyObject = objc2::msg_send![objc2::class!(NSPasteboardItem), new];
        if item_ptr.is_null() {
            return Err(ClipboardError::OperationFailed(
                "NSPasteboardItem.new returned nil".to_string(),
            ));
        }
        let item: Id<AnyObject, Owned> = Id::from_retained_ptr(item_ptr);

        if let Some(text) = payload.text.as_deref() {
            let ty_string = nsstring_from_str(TYPE_STRING)?;
            let text_ns = nsstring_from_str(text)?;
            let ok: bool = objc2::msg_send![&*item, setString: &*text_ns forType: &*ty_string];
            if !ok {
                return Err(ClipboardError::OperationFailed(
                    "failed to set NSPasteboardTypeString".to_string(),
                ));
            }
        }

        if let Some(html) = &payload.html {
            let ty_html = nsstring_from_str(TYPE_HTML)?;
            let html_ns = nsstring_from_str(html)?;
            let ok: bool = objc2::msg_send![&*item, setString: &*html_ns forType: &*ty_html];
            if !ok {
                return Err(ClipboardError::OperationFailed(
                    "failed to set NSPasteboardTypeHTML".to_string(),
                ));
            }
        }

        if let Some(rtf) = &payload.rtf {
            let ty_rtf = nsstring_from_str(TYPE_RTF)?;
            let data = nsdata_from_bytes(rtf.as_bytes())?;
            let ok: bool = objc2::msg_send![&*item, setData: &*data forType: &*ty_rtf];
            if !ok {
                return Err(ClipboardError::OperationFailed(
                    "failed to set NSPasteboardTypeRTF".to_string(),
                ));
            }
        }

        if let Some(ref bytes) = png_bytes {
            let ty_png = nsstring_from_str(TYPE_PNG)?;
            let data = nsdata_from_bytes(bytes)?;
            let ok: bool = objc2::msg_send![&*item, setData: &*data forType: &*ty_png];
            if !ok {
                return Err(ClipboardError::OperationFailed(
                    "failed to set NSPasteboardTypePNG".to_string(),
                ));
            }
        }

        let objects: *mut AnyObject =
            objc2::msg_send![objc2::class!(NSArray), arrayWithObject: &*item];
        if objects.is_null() {
            return Err(ClipboardError::OperationFailed(
                "NSArray.arrayWithObject returned nil".to_string(),
            ));
        }

        let success: bool = objc2::msg_send![pasteboard, writeObjects: objects];
        if success {
            Ok(())
        } else {
            Err(ClipboardError::OperationFailed(
                "NSPasteboard.writeObjects returned false".to_string(),
            ))
        }
    })
}
