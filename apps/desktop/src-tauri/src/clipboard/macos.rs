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
    normalize_base64_str, string_within_limit, ClipboardContent, ClipboardError, ClipboardWritePayload,
    MAX_PNG_BYTES, MAX_TEXT_BYTES,
};

// Ensure the framework crates are linked (and silence `unused_crate_dependencies`).
use objc2_app_kit as _;
use objc2_foundation as _;

// NSPasteboard type constants (UTType identifiers).
const TYPE_STRING: &str = "public.utf8-plain-text"; // NSPasteboardTypeString
const TYPE_HTML: &str = "public.html"; // NSPasteboardTypeHTML
const TYPE_RTF: &str = "public.rtf"; // NSPasteboardTypeRTF
const TYPE_PNG: &str = "public.png"; // NSPasteboardTypePNG
const TYPE_TIFF: &str = "public.tiff"; // NSPasteboardTypeTIFF

// NSBitmapImageFileType values (from AppKit).
const NSBITMAP_IMAGE_FILE_TYPE_TIFF: usize = 0; // NSBitmapImageFileTypeTIFF
const NSBITMAP_IMAGE_FILE_TYPE_PNG: usize = 4; // NSBitmapImageFileTypePNG

// TIFF compression methods (NSTIFFCompression).
const NSTIFF_COMPRESSION_LZW: isize = 5; // NSTIFFCompressionLZW

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

unsafe fn pasteboard_string_for_type_limited(
    pasteboard: *mut AnyObject,
    ty: &AnyObject,
    max_bytes: usize,
) -> Option<String> {
    // -[NSPasteboard stringForType:]
    let ns_str: *mut AnyObject = objc2::msg_send![pasteboard, stringForType: ty];
    if ns_str.is_null() {
        return None;
    }

    // -[NSString lengthOfBytesUsingEncoding:]
    let len: usize = objc2::msg_send![
        ns_str,
        lengthOfBytesUsingEncoding: 4usize /* NSUTF8StringEncoding */
    ];
    if len > max_bytes {
        return None;
    }

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

unsafe fn tiff_to_png_bytes(tiff: &[u8]) -> Result<Vec<u8>, ClipboardError> {
    // [NSBitmapImageRep imageRepWithData:data]
    let data = nsdata_from_bytes(tiff)?;
    let rep: *mut AnyObject =
        objc2::msg_send![objc2::class!(NSBitmapImageRep), imageRepWithData: &*data];
    if rep.is_null() {
        return Err(ClipboardError::OperationFailed(
            "failed to decode TIFF via NSBitmapImageRep".to_string(),
        ));
    }

    // Pass an empty properties dictionary.
    let props: *mut AnyObject = objc2::msg_send![objc2::class!(NSDictionary), dictionary];
    let png_data: *mut AnyObject = objc2::msg_send![
        rep,
        representationUsingType: NSBITMAP_IMAGE_FILE_TYPE_PNG
        properties: props
    ];
    if png_data.is_null() {
        return Err(ClipboardError::OperationFailed(
            "failed to encode PNG via NSBitmapImageRep".to_string(),
        ));
    }

    let len: usize = objc2::msg_send![png_data, length];
    if len == 0 {
        return Err(ClipboardError::OperationFailed(
            "converted PNG was empty".to_string(),
        ));
    }
    if len > MAX_PNG_BYTES {
        return Err(ClipboardError::OperationFailed(format!(
            "converted PNG exceeds maximum size ({MAX_PNG_BYTES} bytes)"
        )));
    }

    let bytes = nsdata_to_vec(png_data, MAX_PNG_BYTES);
    if bytes.is_empty() {
        return Err(ClipboardError::OperationFailed(
            "failed to copy converted PNG bytes".to_string(),
        ));
    }

    Ok(bytes)
}

unsafe fn png_to_tiff_bytes(png: &[u8]) -> Result<Vec<u8>, ClipboardError> {
    // [NSBitmapImageRep imageRepWithData:data]
    let data = nsdata_from_bytes(png)?;
    let rep: *mut AnyObject =
        objc2::msg_send![objc2::class!(NSBitmapImageRep), imageRepWithData: &*data];
    if rep.is_null() {
        return Err(ClipboardError::OperationFailed(
            "failed to decode PNG via NSBitmapImageRep".to_string(),
        ));
    }

    // Use LZW compression by default to keep TIFF payload sizes reasonable.
    //
    // Many consumers (including built-in macOS apps) accept compressed TIFF, and the compression
    // can drastically reduce size compared to uncompressed RGBA (which would quickly exceed our
    // `MAX_PNG_BYTES` limit for moderately-sized images).
    let compression_key = nsstring_from_str("NSImageCompressionMethod")?;
    let compression_value: *mut AnyObject = objc2::msg_send![
        objc2::class!(NSNumber),
        numberWithInteger: NSTIFF_COMPRESSION_LZW
    ];
    if compression_value.is_null() {
        return Err(ClipboardError::OperationFailed(
            "failed to allocate NSNumber for TIFF compression".to_string(),
        ));
    }
    let props: *mut AnyObject = objc2::msg_send![
        objc2::class!(NSDictionary),
        dictionaryWithObject: compression_value
        forKey: &*compression_key
    ];
    if props.is_null() {
        return Err(ClipboardError::OperationFailed(
            "failed to allocate NSDictionary for TIFF compression".to_string(),
        ));
    }
    let tiff_data: *mut AnyObject = objc2::msg_send![
        rep,
        representationUsingType: NSBITMAP_IMAGE_FILE_TYPE_TIFF
        properties: props
    ];
    if tiff_data.is_null() {
        return Err(ClipboardError::OperationFailed(
            "failed to encode TIFF via NSBitmapImageRep".to_string(),
        ));
    }

    let len: usize = objc2::msg_send![tiff_data, length];
    if len == 0 {
        return Err(ClipboardError::OperationFailed(
            "converted TIFF was empty".to_string(),
        ));
    }
    if len > MAX_PNG_BYTES {
        return Err(ClipboardError::OperationFailed(format!(
            "converted TIFF exceeds maximum size ({MAX_PNG_BYTES} bytes)"
        )));
    }

    let bytes = nsdata_to_vec(tiff_data, MAX_PNG_BYTES);
    if bytes.is_empty() {
        return Err(ClipboardError::OperationFailed(
            "failed to copy converted TIFF bytes".to_string(),
        ));
    }

    Ok(bytes)
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
        let ty_tiff = nsstring_from_str(TYPE_TIFF)?;

        let text = pasteboard_string_for_type_limited(pasteboard, &*ty_string, MAX_TEXT_BYTES)
            .or_else(|| {
                pasteboard_data_for_type(pasteboard, &*ty_string, MAX_TEXT_BYTES)
                    .map(|bytes| String::from_utf8_lossy(&bytes).into_owned())
            })
            .and_then(|s| string_within_limit(s, MAX_TEXT_BYTES));

        let html = pasteboard_string_for_type_limited(pasteboard, &*ty_html, MAX_TEXT_BYTES)
            .or_else(|| {
                pasteboard_data_for_type(pasteboard, &*ty_html, MAX_TEXT_BYTES)
                    .map(|bytes| String::from_utf8_lossy(&bytes).into_owned())
            })
            .and_then(|s| string_within_limit(s, MAX_TEXT_BYTES));

        let rtf = pasteboard_data_for_type(pasteboard, &*ty_rtf, MAX_TEXT_BYTES)
            .map(|bytes| String::from_utf8_lossy(&bytes).into_owned())
            .and_then(|s| string_within_limit(s, MAX_TEXT_BYTES));

        // Prefer PNG when present, but fall back to TIFF (converted to PNG) for interoperability
        // with macOS apps that primarily put `public.tiff` on the pasteboard.
        let png_base64 = pasteboard_data_for_type(pasteboard, &*ty_png, MAX_PNG_BYTES)
            .map(|bytes| STANDARD.encode(&bytes))
            .or_else(|| {
                let tiff = pasteboard_data_for_type(pasteboard, &*ty_tiff, MAX_PNG_BYTES)?;
                let png = tiff_to_png_bytes(&tiff).ok()?;
                Some(STANDARD.encode(&png))
            });

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
        .map(normalize_base64_str)
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

            // Also provide `public.tiff` (NSPasteboardTypeTIFF) for better macOS interoperability.
            // Many AppKit apps prefer TIFF even when PNG is present.
            if bytes.len() <= MAX_PNG_BYTES {
                if let Ok(tiff) = png_to_tiff_bytes(bytes) {
                    if !tiff.is_empty() && tiff.len() <= MAX_PNG_BYTES {
                        let ty_tiff = nsstring_from_str(TYPE_TIFF)?;
                        let tiff_data = nsdata_from_bytes(&tiff)?;
                        let ok: bool =
                            objc2::msg_send![&*item, setData: &*tiff_data forType: &*ty_tiff];
                        if !ok {
                            return Err(ClipboardError::OperationFailed(
                                "failed to set NSPasteboardTypeTIFF".to_string(),
                            ));
                        }
                    }
                }
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

#[cfg(test)]
mod tests {
    use super::*;
    use base64::Engine as _;

    fn png_dimensions(png: &[u8]) -> Option<(u32, u32)> {
        const SIG: &[u8; 8] = b"\x89PNG\r\n\x1a\n";
        if png.len() < 24 {
            return None;
        }
        if &png[0..8] != SIG {
            return None;
        }
        if &png[12..16] != b"IHDR" {
            return None;
        }
        let w = u32::from_be_bytes(png[16..20].try_into().ok()?);
        let h = u32::from_be_bytes(png[20..24].try_into().ok()?);
        Some((w, h))
    }

    #[test]
    fn png_tiff_png_roundtrip_preserves_dimensions() {
        // 1x1 transparent PNG.
        let png = base64::engine::general_purpose::STANDARD
            .decode(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO9C9VwAAAAASUVORK5CYII=",
            )
            .unwrap();
        let dims_before = png_dimensions(&png).expect("valid png");

        let tiff = autoreleasepool(|_| unsafe { png_to_tiff_bytes(&png) }).expect("png -> tiff");
        let png2 =
            autoreleasepool(|_| unsafe { tiff_to_png_bytes(&tiff) }).expect("tiff -> png");
        let dims_after = png_dimensions(&png2).expect("valid png output");

        assert_eq!(dims_before, dims_after);
    }
}
