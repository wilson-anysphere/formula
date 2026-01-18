use super::{ClipboardContent, ClipboardError, ClipboardWritePayload};

/// Maximum number of bytes we're willing to accept for a decoded clipboard image pixel buffer on
/// Linux.
///
/// When `image/png` is not available we fall back to `gtk_clipboard_wait_for_image`, which decodes
/// images via GdkPixbuf loaders. Highly-compressible images can expand to very large decoded pixel
/// buffers (decompression bombs). Treat images as best-effort and skip anything that would require
/// an unreasonably large allocation.
const MAX_DECODED_IMAGE_BYTES: usize = 4 * super::MAX_PNG_BYTES;

fn decoded_pixbuf_len(rowstride: i32, height: i32) -> Option<usize> {
    if rowstride <= 0 || height <= 0 {
        return None;
    }
    let rowstride = usize::try_from(rowstride).ok()?;
    let height = usize::try_from(height).ok()?;
    rowstride.checked_mul(height)
}

fn normalize_target_name(target: &str) -> String {
    target.trim().to_ascii_lowercase()
}

fn latin1_encode_if_possible(text: &str) -> Option<Vec<u8>> {
    let mut bytes: Vec<u8> = Vec::new();
    let _ = bytes.try_reserve(text.len());
    for ch in text.chars() {
        let codepoint = ch as u32;
        if codepoint > 0xFF {
            return None;
        }
        bytes.push(codepoint as u8);
    }
    Some(bytes)
}

fn trim_trailing_nuls(bytes: &[u8]) -> &[u8] {
    let mut end = bytes.len();
    while end > 0 && bytes[end - 1] == 0 {
        end -= 1;
    }
    &bytes[..end]
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Utf8DecodeMode {
    /// Reject invalid UTF-8 (return `None`) so callers can fall back to other targets.
    Strict,
    /// Decode invalid UTF-8 using replacement characters (U+FFFD).
    Lossy,
}

fn decode_latin1(bytes: &[u8]) -> Option<String> {
    let bytes = trim_trailing_nuls(bytes);
    if bytes.is_empty() {
        return None;
    }

    let mut out = String::new();
    let _ = out.try_reserve(bytes.len());
    for &b in bytes {
        out.push(char::from(b));
    }
    (!out.is_empty()).then_some(out)
}

fn decode_utf8(bytes: &[u8], mode: Utf8DecodeMode) -> Option<String> {
    let bytes = trim_trailing_nuls(bytes);
    if bytes.is_empty() {
        return None;
    }
    match mode {
        Utf8DecodeMode::Strict => {
            let s = std::str::from_utf8(bytes).ok()?;
            (!s.is_empty()).then_some(s.to_string())
        }
        Utf8DecodeMode::Lossy => {
            let s = String::from_utf8_lossy(bytes);
            (!s.is_empty()).then_some(s.to_string())
        }
    }
}

/// Decode clipboard text bytes based on the advertised target name.
///
/// - `STRING` is specified as ISO-8859-1 (Latin-1) on X11, so we decode it as a direct
///   byte->Unicode mapping (U+00XX).
/// - `UTF8_STRING` and MIME-ish `text/*` targets are expected to contain UTF-8; we attempt strict
///   decoding first so invalid bytes don't silently turn into U+FFFD.
/// - `TEXT` is a legacy X11 target with unspecified locale-dependent encoding; we try UTF-8 first
///   (common in modern apps) and fall back to Latin-1 as a best-effort byte-preserving decode.
fn decode_text_for_target_impl(
    target: &str,
    bytes: &[u8],
    utf8_mode: Utf8DecodeMode,
) -> Option<String> {
    let normalized = normalize_target_name(target);
    if normalized == "string" {
        return decode_latin1(bytes);
    }
    if normalized == "text" {
        return decode_utf8(bytes, Utf8DecodeMode::Strict).or_else(|| decode_latin1(bytes));
    }

    if normalized == "utf8_string" || normalized.starts_with("text/") {
        return decode_utf8(bytes, utf8_mode);
    }

    // Unknown text targets: treat as UTF-8 by default, matching historical behavior, but keep
    // strictness configurable so callers can choose whether to accept replacement characters.
    decode_utf8(bytes, utf8_mode)
}

fn decode_text_for_target(target: &str, bytes: &[u8]) -> Option<String> {
    decode_text_for_target_impl(target, bytes, Utf8DecodeMode::Strict)
}

fn decode_text_for_target_lossy_utf8(target: &str, bytes: &[u8]) -> Option<String> {
    decode_text_for_target_impl(target, bytes, Utf8DecodeMode::Lossy)
}

/// Decode raw clipboard bytes as lossy UTF-8 and enforce a post-decode UTF-8 byte limit.
///
/// [`String::from_utf8_lossy`] can expand invalid input bytes into U+FFFD (3 UTF-8 bytes), so
/// callers that cap the *input* byte length must still apply the cap *after* decoding to avoid
/// oversized strings (and therefore oversized IPC payloads).
fn bytes_to_utf8_within_limit(bytes: &[u8], max_bytes: usize) -> Option<String> {
    // Use the existing UTF-8 decoding path (including NUL trimming) to keep behavior consistent.
    let s = decode_text_for_target_lossy_utf8("UTF8_STRING", bytes)?;
    super::string_within_limit(s, max_bytes)
}

fn target_prefers_utf8(target: &str) -> bool {
    let normalized = normalize_target_name(target);
    normalized == "utf8_string" || normalized.starts_with("text/")
}

/// Choose the "best" clipboard target from a list of advertised targets.
///
/// Linux clipboard targets are free-form atoms and different apps may advertise the same content
/// using slightly different target strings (e.g. `text/html;charset=utf-8`).
///
/// This helper prefers:
/// 1) Exact case-insensitive matches for each preferred prefix
/// 2) Otherwise, case-insensitive prefix matches (e.g. `text/html; charset=utf-8`)
fn choose_best_target<'a, T: AsRef<str>>(
    targets: &'a [T],
    preferred_prefixes: &[&str],
) -> Option<&'a str> {
    for preferred in preferred_prefixes {
        // First, prefer exact matches (ignoring case/whitespace).
        for target in targets {
            if normalize_target_name(target.as_ref()) == *preferred {
                return Some(target.as_ref());
            }
        }

        // Fall back to prefix matches (ignoring case/whitespace).
        for target in targets {
            if normalize_target_name(target.as_ref()).starts_with(preferred) {
                return Some(target.as_ref());
            }
        }
    }

    None
}

/// Returns all matching targets in preference order.
///
/// This is similar to [`choose_best_target`], but instead of returning only the single best match,
/// it returns a list of candidates so callers can try decoding multiple targets until one works
/// (e.g. when an app advertises `UTF8_STRING` but the payload contains invalid UTF-8).
fn choose_targets_in_order<'a, T: AsRef<str>>(
    targets: &'a [T],
    preferred_prefixes: &[&str],
) -> Vec<&'a str> {
    let mut out = Vec::new();

    for preferred in preferred_prefixes {
        // First, collect exact matches (ignoring case/whitespace).
        for target in targets {
            let target = target.as_ref();
            if normalize_target_name(target) == *preferred && !out.iter().any(|&t| t == target) {
                out.push(target);
            }
        }

        // Then, collect prefix matches (ignoring case/whitespace).
        for target in targets {
            let target = target.as_ref();
            if normalize_target_name(target).starts_with(preferred) && !out.iter().any(|&t| t == target)
            {
                out.push(target);
            }
        }
    }

    out
}

#[cfg(feature = "desktop")]
mod gtk_backend {
    use base64::{engine::general_purpose::STANDARD, Engine as _};

    use super::super::{
        debug_clipboard_log, normalize_base64_str, string_within_limit, MAX_PNG_BYTES, MAX_TEXT_BYTES,
    };
    use super::{
        bytes_to_utf8_within_limit, choose_best_target, decode_text_for_target,
        decoded_pixbuf_len, latin1_encode_if_possible, target_prefers_utf8, ClipboardContent,
        ClipboardError, ClipboardWritePayload, MAX_DECODED_IMAGE_BYTES,
    };
    use crate::clipboard_fallback;

    // GTK clipboard APIs must be called on the GTK main thread.
    //
    // Tauri commands may execute on a background thread (Tokio worker). GTK itself is not thread-safe
    // and `gtk-rs` asserts this at runtime via `assert_initialized_main_thread!()`.
    //
    // We dispatch operations onto the default GLib main context (the same one used by GTK) and wait
    // synchronously for the result.
    fn with_gtk_main_thread<R>(
        op: impl FnOnce() -> Result<R, ClipboardError> + Send + 'static,
    ) -> Result<R, ClipboardError>
    where
        R: Send + 'static,
    {
        if gtk::is_initialized_main_thread() {
            return op();
        }

        let ctx = glib::MainContext::default();
        // `g_main_context_invoke` runs inline when called from the context owner thread. This can
        // happen early in startup (before GTK has been initialized) and we must not block waiting
        // for the main loop in that case.
        if ctx.is_owner() {
            return op();
        }
        let (tx, rx) = std::sync::mpsc::channel::<Result<R, ClipboardError>>();

        ctx.invoke(move || {
            let res =
                std::panic::catch_unwind(std::panic::AssertUnwindSafe(op)).unwrap_or_else(|_| {
                    Err(ClipboardError::OperationFailed(
                        "GTK clipboard operation panicked".to_string(),
                    ))
                });
            let _ = tx.send(res);
        });

        use std::sync::mpsc::RecvTimeoutError;

        match rx.recv_timeout(std::time::Duration::from_secs(5)) {
            Ok(res) => res,
            Err(RecvTimeoutError::Timeout) => Err(ClipboardError::OperationFailed(
                "timed out waiting for GTK main thread clipboard operation (is the main loop running?)"
                    .to_string(),
            )),
            Err(RecvTimeoutError::Disconnected) => Err(ClipboardError::OperationFailed(
                "failed to receive result from GTK main thread".to_string(),
            )),
        }
    }

    fn ensure_gtk() -> Result<(), ClipboardError> {
        gtk::init().map_err(|e| ClipboardError::Unavailable(e.to_string()))
    }

    fn clipboard_target_names(clipboard: &gtk::Clipboard) -> Option<Vec<String>> {
        clipboard.wait_for_targets().map(|atoms| {
            atoms
                .into_iter()
                .map(|atom| atom.name().to_string())
                .collect::<Vec<_>>()
        })
    }

    fn wait_for_utf8_targets(
        clipboard: &gtk::Clipboard,
        targets: &[&str],
        max_bytes: usize,
    ) -> Option<String> {
        wait_for_utf8_targets_with_source(clipboard, targets, max_bytes).map(|(_, value)| value)
    }

    fn wait_for_utf8_targets_with_source<'a>(
        clipboard: &gtk::Clipboard,
        targets: &[&'a str],
        max_bytes: usize,
    ) -> Option<(&'a str, String)> {
        let mut lossy_utf8_fallback: Option<(&'a str, Vec<u8>)> = None;
        for &target in targets {
            let atom = gdk::Atom::intern(target);
            let Some(data) = clipboard.wait_for_contents(&atom) else {
                continue;
            };
            // `SelectionData::data()` copies the payload into a new Vec. Check the length first so
            // we don't duplicate huge clipboard contents in Rust memory.
            let len = data.length();
            if len <= 0 {
                continue;
            }
            let Ok(len) = usize::try_from(len) else {
                continue;
            };
            if len > max_bytes {
                continue;
            }
            let bytes = data.data();
            if let Some(s) = decode_text_for_target(target, &bytes) {
                return Some((target, s));
            }
            // As a last resort (only when no other target decodes successfully), allow lossy UTF-8
            // decoding for targets that are explicitly expected to be UTF-8.
            if lossy_utf8_fallback.is_none()
                && target_prefers_utf8(target)
                && !super::trim_trailing_nuls(&bytes).is_empty()
            {
                lossy_utf8_fallback = Some((target, bytes));
            }
        }
        lossy_utf8_fallback.and_then(|(target, bytes)| {
            bytes_to_utf8_within_limit(&bytes, max_bytes).map(|s| (target, s))
        })
    }

    fn wait_for_bytes_base64(
        clipboard: &gtk::Clipboard,
        target: &str,
        max_bytes: usize,
    ) -> Option<(String, usize)> {
        let atom = gdk::Atom::intern(target);
        let data = clipboard.wait_for_contents(&atom)?;
        // Avoid copying large clipboard payloads into a second buffer.
        let len = data.length();
        if len <= 0 {
            return None;
        }
        let Ok(len) = usize::try_from(len) else {
            return None;
        };
        if len > max_bytes {
            return None;
        }
        let bytes = data.data();
        if bytes.is_empty() {
            None
        } else {
            Some((STANDARD.encode(bytes), len))
        }
    }

    pub(super) fn read() -> Result<ClipboardContent, ClipboardError> {
        with_gtk_main_thread(|| {
            ensure_gtk()?;

            let read_from_clipboard = |clipboard: &gtk::Clipboard, selection: &'static str| {
                let targets = clipboard_target_names(clipboard);

                let mut text_target: Option<String> = None;
                let mut html_target: Option<String> = None;
                let mut rtf_target: Option<String> = None;
                let mut image_target: Option<String> = None;
                let mut image_bytes: Option<usize> = None;
                let mut image_pixbuf_fallback = false;

                // Read plain text in a size-limited way.
                //
                // Different apps advertise text using different targets (e.g. X11 atoms like
                // `UTF8_STRING`/`STRING`/`TEXT`, or MIME-like targets like
                // `text/plain;charset=utf-8`). We prefer using `wait_for_contents` so we can check
                // the `SelectionData` length before copying large buffers into Rust.
                let text = match targets.as_deref() {
                    Some(targets) => {
                        let candidates = super::choose_targets_in_order(
                            targets,
                            &[
                                "text/plain;charset=utf-8",
                                "text/plain; charset=utf-8",
                                "text/plain",
                                "utf8_string",
                                "string",
                                "text",
                            ],
                        );
                        if candidates.is_empty() {
                            None
                        } else {
                            wait_for_utf8_targets_with_source(clipboard, &candidates, MAX_TEXT_BYTES)
                                .map(|(target, value)| {
                                    text_target = Some(target.to_string());
                                    value
                                })
                        }
                    }
                    None => wait_for_utf8_targets_with_source(
                        clipboard,
                        &[
                            "text/plain;charset=utf-8",
                            "text/plain; charset=utf-8",
                            "text/plain",
                            "UTF8_STRING",
                            "STRING",
                            "TEXT",
                        ],
                        MAX_TEXT_BYTES,
                    )
                    .map(|(target, value)| {
                        text_target = Some(target.to_string());
                        value
                    }),
                }
                .and_then(|s| string_within_limit(s, MAX_TEXT_BYTES));

                let html = match targets.as_deref() {
                    Some(targets) => {
                        let candidates = super::choose_targets_in_order(targets, &["text/html"]);
                        if candidates.is_empty() {
                            None
                        } else {
                            wait_for_utf8_targets_with_source(clipboard, &candidates, MAX_TEXT_BYTES)
                                .map(|(target, value)| {
                                    html_target = Some(target.to_string());
                                    value
                                })
                        }
                    }
                    // If target enumeration isn't available, fall back to the canonical target.
                    None => wait_for_utf8_targets_with_source(
                        clipboard,
                        &[
                            "text/html",
                            "text/html;charset=utf-8",
                            "text/html; charset=utf-8",
                        ],
                        MAX_TEXT_BYTES,
                    )
                    .map(|(target, value)| {
                        html_target = Some(target.to_string());
                        value
                    }),
                }
                .and_then(|s| string_within_limit(s, MAX_TEXT_BYTES));
                let rtf = match targets.as_deref() {
                    Some(targets) => {
                        let candidates = super::choose_targets_in_order(
                            targets,
                            &["text/rtf", "application/rtf", "application/x-rtf"],
                        );
                        if candidates.is_empty() {
                            None
                        } else {
                            wait_for_utf8_targets_with_source(clipboard, &candidates, MAX_TEXT_BYTES)
                                .map(|(target, value)| {
                                    rtf_target = Some(target.to_string());
                                    value
                                })
                        }
                    }
                    None => wait_for_utf8_targets_with_source(
                        clipboard,
                        &[
                            "text/rtf",
                            "text/rtf;charset=utf-8",
                            "text/rtf; charset=utf-8",
                            "application/rtf",
                            "application/x-rtf",
                        ],
                        MAX_TEXT_BYTES,
                    )
                    .map(|(target, value)| {
                        rtf_target = Some(target.to_string());
                        value
                    }),
                }
                .and_then(|s| string_within_limit(s, MAX_TEXT_BYTES));
                let image_png_base64 = match targets.as_deref() {
                    Some(targets) => choose_best_target(targets, &["image/png"]).and_then(|t| {
                        wait_for_bytes_base64(clipboard, t, MAX_PNG_BYTES).map(|(b64, len)| {
                            image_target = Some(t.to_string());
                            image_bytes = Some(len);
                            b64
                        })
                    }),
                    None => wait_for_bytes_base64(clipboard, "image/png", MAX_PNG_BYTES).map(|(b64, len)| {
                        image_target = Some("image/png".to_string());
                        image_bytes = Some(len);
                        b64
                    }),
                }
                .or_else(|| {
                    // Some applications expose images on the clipboard without an `image/png` target.
                    // Fall back to GTK's pixbuf API and re-encode to PNG (requires image loaders).
                    let pixbuf = clipboard.wait_for_image()?;
                    // Guard against decompression bombs / huge decoded images.
                    let decoded_len = decoded_pixbuf_len(pixbuf.rowstride(), pixbuf.height())
                        .unwrap_or(usize::MAX);
                    if decoded_len > MAX_DECODED_IMAGE_BYTES {
                        return None;
                    }
                    let bytes = pixbuf.save_to_bufferv("png", &[]).ok()?;
                    if bytes.len() > MAX_PNG_BYTES {
                        return None;
                    }
                    image_target = Some("pixbuf->image/png".to_string());
                    image_bytes = Some(bytes.len());
                    image_pixbuf_fallback = true;
                    Some(STANDARD.encode(bytes))
                });

                let content = ClipboardContent {
                    text,
                    html,
                    rtf,
                    image_png_base64,
                };

                let text_bytes = content.text.as_ref().map(|s| s.as_bytes().len());
                let html_bytes = content.html.as_ref().map(|s| s.as_bytes().len());
                let rtf_bytes = content.rtf.as_ref().map(|s| s.as_bytes().len());
                debug_clipboard_log(format_args!(
                    "linux read({selection}): targets_len={:?} text_target={text_target:?} text_bytes={text_bytes:?} html_target={html_target:?} html_bytes={html_bytes:?} rtf_target={rtf_target:?} rtf_bytes={rtf_bytes:?} image_target={image_target:?} image_bytes={image_bytes:?} pixbuf_fallback={image_pixbuf_fallback} caps(text={MAX_TEXT_BYTES}, png={MAX_PNG_BYTES})",
                    targets.as_ref().map(|t| t.len())
                ));

                content
            };

            let clipboard = gtk::Clipboard::get(&gdk::SELECTION_CLIPBOARD);
            let content = read_from_clipboard(&clipboard, "CLIPBOARD");

            // On X11, some apps only populate PRIMARY selection (middle-click paste). When
            // CLIPBOARD has no usable content we may fall back to PRIMARY.
            //
            // By default we skip this fallback on Wayland to avoid changing semantics where
            // PRIMARY may not exist or behave differently, but users can override via:
            // `FORMULA_CLIPBOARD_PRIMARY_SELECTION=0/false/no` to disable or `=1/true/yes` to force-enable
            // (see `clipboard_fallback`).
            let has_usable_data = clipboard_fallback::has_usable_clipboard_data(
                content.text.as_deref(),
                content.html.as_deref(),
                content.rtf.as_deref(),
                content.image_png_base64.as_deref(),
            );
            if !has_usable_data && clipboard_fallback::should_attempt_primary_selection_from_env() {
                let primary = gtk::Clipboard::get(&gdk::SELECTION_PRIMARY);
                return Ok(read_from_clipboard(&primary, "PRIMARY"));
            }

            Ok(content)
        })
    }

    pub(super) fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
        use std::sync::Arc;

        let text = payload.text.clone().map(Arc::new);
        let html = payload.html.clone().map(Arc::new);
        let rtf = payload.rtf.clone().map(Arc::new);
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
            .transpose()?
            .map(Arc::new);

        with_gtk_main_thread(move || {
            ensure_gtk()?;

            let set_clipboard_data = |clipboard: &gtk::Clipboard| {
                const INFO_TEXT_UTF8: u32 = 1;
                const INFO_HTML: u32 = 2;
                const INFO_RTF: u32 = 3;
                const INFO_PNG: u32 = 4;
                const INFO_TEXT_STRING: u32 = 5;

                // Do not restrict targets based on app identity; we want both intra-app copy/paste and
                // interoperability with other apps (LibreOffice, browser, image editors, etc.).
                let flags = gtk::TargetFlags::empty();
                let mut targets: Vec<gtk::TargetEntry> = Vec::new();
                let text_string_bytes = text
                    .as_deref()
                    .and_then(|t| latin1_encode_if_possible(t))
                    .map(Arc::new);

                if text.is_some() {
                    // Provide common plaintext targets so other apps can pick their preferred flavor.
                    // (GTK's own `set_text` would also do this, but we need to set multiple targets at once.)
                    targets.push(gtk::TargetEntry::new("text/plain", flags, INFO_TEXT_UTF8));
                    targets.push(gtk::TargetEntry::new(
                        "text/plain;charset=utf-8",
                        flags,
                        INFO_TEXT_UTF8,
                    ));
                    targets.push(gtk::TargetEntry::new("UTF8_STRING", flags, INFO_TEXT_UTF8));

                    // X11 `STRING` is always ISO-8859-1 (Latin-1). Only advertise it when we can
                    // actually supply Latin-1 bytes; otherwise legacy consumers that request
                    // `STRING` would see mojibake.
                    if text_string_bytes.is_some() {
                        targets.push(gtk::TargetEntry::new(
                            "STRING",
                            flags,
                            INFO_TEXT_STRING,
                        ));
                    }

                    // `TEXT` is a legacy X11 target whose encoding depends on the current locale.
                    // On modern Linux desktops this is almost always UTF-8, so we treat it as
                    // equivalent to `UTF8_STRING` and provide UTF-8 bytes.
                    targets.push(gtk::TargetEntry::new("TEXT", flags, INFO_TEXT_UTF8));
                }

                if html.is_some() {
                    targets.push(gtk::TargetEntry::new("text/html", flags, INFO_HTML));
                    targets.push(gtk::TargetEntry::new(
                        "text/html;charset=utf-8",
                        flags,
                        INFO_HTML,
                    ));
                }

                if rtf.is_some() {
                    targets.push(gtk::TargetEntry::new("text/rtf", flags, INFO_RTF));
                    targets.push(gtk::TargetEntry::new("application/rtf", flags, INFO_RTF));
                    targets.push(gtk::TargetEntry::new("application/x-rtf", flags, INFO_RTF));
                }

                if png_bytes.is_some() {
                    targets.push(gtk::TargetEntry::new("image/png", flags, INFO_PNG));
                }

                let text_bytes = text.as_ref().map(|s| s.as_bytes().len());
                let html_bytes = html.as_ref().map(|s| s.as_bytes().len());
                let rtf_bytes = rtf.as_ref().map(|s| s.as_bytes().len());
                let png_len = png_bytes.as_ref().map(|b| b.len());
                debug_clipboard_log(format_args!(
                    "linux write: targets_len={} has_text={} has_html={} has_rtf={} has_png={} text_bytes={text_bytes:?} html_bytes={html_bytes:?} rtf_bytes={rtf_bytes:?} png_bytes={png_len:?} caps(text={MAX_TEXT_BYTES}, png={MAX_PNG_BYTES})",
                    targets.len(),
                    text.is_some(),
                    html.is_some(),
                    rtf.is_some(),
                    png_bytes.is_some(),
                ));

                // Note: the closure captures owned copies of the strings/bytes so the clipboard stays
                // valid after this function returns.
                let text = text.clone();
                let text_string_bytes = text_string_bytes.clone();
                let html = html.clone();
                let rtf = rtf.clone();
                let png_bytes = png_bytes.clone();
                let success =
                    clipboard.set_with_data(&targets, move |_clipboard, selection_data, info| {
                        match info {
                            INFO_TEXT_UTF8 => {
                                if let Some(ref text) = text {
                                    selection_data.set(
                                        &selection_data.target(),
                                        8,
                                        text.as_bytes(),
                                    );
                                }
                            }
                            INFO_TEXT_STRING => {
                                if let Some(ref bytes) = text_string_bytes {
                                    selection_data
                                        .set(&selection_data.target(), 8, bytes.as_slice());
                                }
                            }
                            INFO_HTML => {
                                if let Some(ref html) = html {
                                    selection_data.set(
                                        &selection_data.target(),
                                        8,
                                        html.as_bytes(),
                                    );
                                }
                            }
                            INFO_RTF => {
                                if let Some(ref rtf) = rtf {
                                    selection_data.set(&selection_data.target(), 8, rtf.as_bytes());
                                }
                            }
                            INFO_PNG => {
                                if let Some(ref bytes) = png_bytes {
                                    selection_data.set(&selection_data.target(), 8, bytes);
                                }
                            }
                            _ => {}
                        }
                    });

                if !success {
                    return Err(ClipboardError::OperationFailed(
                        "gtk_clipboard_set_with_data returned false".to_string(),
                    ));
                }

                Ok(())
            };

            let clipboard = gtk::Clipboard::get(&gdk::SELECTION_CLIPBOARD);
            set_clipboard_data(&clipboard)?;

            // Best-effort request to persist clipboard data via a clipboard manager (X11).
            clipboard.store();

            // Optional: also populate X11 PRIMARY selection (middle-click paste) when available.
            // This is controlled by the same env/heuristic gate as the read-time PRIMARY fallback.
            if clipboard_fallback::should_attempt_primary_selection_from_env() {
                let primary = gtk::Clipboard::get(&gdk::SELECTION_PRIMARY);
                let _ = set_clipboard_data(&primary);
            }

            Ok(())
        })
    }
}

#[cfg(feature = "desktop")]
pub fn read() -> Result<ClipboardContent, ClipboardError> {
    gtk_backend::read()
}

#[cfg(feature = "desktop")]
pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    gtk_backend::write(payload)
}

// The GTK-backed Linux clipboard implementation relies on system libraries that we intentionally
// keep behind the `desktop` feature. Provide a stub for unit tests so this module compiles under
// `cfg(test)` without enabling the full desktop toolchain.
#[cfg(not(feature = "desktop"))]
#[cfg_attr(test, allow(dead_code))]
pub fn read() -> Result<ClipboardContent, ClipboardError> {
    Err(ClipboardError::Unavailable(
        "GTK clipboard backend requires the `desktop` feature".to_string(),
    ))
}

#[cfg(not(feature = "desktop"))]
#[cfg_attr(test, allow(dead_code))]
pub fn write(_payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    Err(ClipboardError::Unavailable(
        "GTK clipboard backend requires the `desktop` feature".to_string(),
    ))
}

#[cfg(test)]
mod tests {
    use super::{
        bytes_to_utf8_within_limit, choose_best_target, decode_text_for_target,
        decode_text_for_target_lossy_utf8, decoded_pixbuf_len, normalize_target_name,
        target_prefers_utf8, MAX_DECODED_IMAGE_BYTES,
    };

    #[test]
    fn normalize_target_name_lowercases_and_trims() {
        assert_eq!(
            normalize_target_name("  Text/Html;Charset=UTF-8 \t"),
            "text/html;charset=utf-8"
        );
    }

    #[cfg(not(feature = "desktop"))]
    #[test]
    fn read_returns_unavailable_without_desktop_feature() {
        let err =
            super::read().expect_err("read should be unavailable without the `desktop` feature");
        match err {
            super::ClipboardError::Unavailable(msg) => {
                assert!(
                    msg.contains("GTK clipboard backend requires the `desktop` feature"),
                    "unexpected error message: {msg}"
                );
            }
            other => panic!("expected ClipboardError::Unavailable, got {other:?}"),
        }
    }

    #[cfg(not(feature = "desktop"))]
    #[test]
    fn write_returns_unavailable_without_desktop_feature() {
        let payload = super::ClipboardWritePayload::default();
        let err = super::write(&payload)
            .expect_err("write should be unavailable without the `desktop` feature");
        match err {
            super::ClipboardError::Unavailable(msg) => {
                assert!(
                    msg.contains("GTK clipboard backend requires the `desktop` feature"),
                    "unexpected error message: {msg}"
                );
            }
            other => panic!("expected ClipboardError::Unavailable, got {other:?}"),
        }
    }
    #[test]
    fn choose_best_target_prefers_exact_match() {
        let targets = vec!["text/html;charset=utf-8", "text/html"];
        let best = choose_best_target(&targets, &["text/html"]);
        assert_eq!(best, Some("text/html"));
    }

    #[test]
    fn choose_best_target_supports_text_plain_with_charset_suffix() {
        let targets = vec!["text/plain;charset=utf-8", "UTF8_STRING"];
        let best = choose_best_target(&targets, &["text/plain"]);
        assert_eq!(best, Some("text/plain;charset=utf-8"));
    }

    #[test]
    fn choose_best_target_supports_text_plain_with_charset_suffix_with_space() {
        let targets = vec!["text/plain; charset=utf-8", "UTF8_STRING"];
        let best = choose_best_target(&targets, &["text/plain"]);
        assert_eq!(best, Some("text/plain; charset=utf-8"));
    }

    #[test]
    fn choose_best_target_supports_image_png_with_parameters() {
        let targets = vec!["image/png;foo=bar", "image/jpeg"];
        let best = choose_best_target(&targets, &["image/png"]);
        assert_eq!(best, Some("image/png;foo=bar"));
    }

    #[test]
    fn choose_best_target_falls_back_to_prefix_match_case_insensitive() {
        let targets = vec!["UTF8_STRING", "TEXT/HTML;CHARSET=UTF-8", "text/plain"];
        let best = choose_best_target(&targets, &["text/html"]);
        assert_eq!(best, Some("TEXT/HTML;CHARSET=UTF-8"));
    }

    #[test]
    fn choose_best_target_supports_rtf_aliases() {
        let targets = vec!["application/rtf", "text/rtf;charset=utf-8"];
        let best = choose_best_target(
            &targets,
            &["text/rtf", "application/rtf", "application/x-rtf"],
        );
        assert_eq!(best, Some("text/rtf;charset=utf-8"));
    }

    #[test]
    fn choose_best_target_supports_x_rtf_alias() {
        let targets = vec!["application/x-rtf", "text/plain"];
        let best = choose_best_target(
            &targets,
            &["text/rtf", "application/rtf", "application/x-rtf"],
        );
        assert_eq!(best, Some("application/x-rtf"));
    }

    #[test]
    fn choose_best_target_returns_none_when_no_match() {
        let targets = vec!["text/plain", "UTF8_STRING"];
        let best = choose_best_target(&targets, &["text/html"]);
        assert_eq!(best, None);
    }

    #[test]
    fn latin1_encode_if_possible_encodes_ascii() {
        assert_eq!(
            super::latin1_encode_if_possible("hello"),
            Some(b"hello".to_vec())
        );
    }

    #[test]
    fn latin1_encode_if_possible_encodes_latin1_codepoints() {
        // "café" where "é" is U+00E9 (0xE9).
        assert_eq!(
            super::latin1_encode_if_possible("café"),
            Some(vec![b'c', b'a', b'f', 0xE9])
        );
    }

    #[test]
    fn latin1_encode_if_possible_returns_none_for_non_latin1() {
        // Euro sign is U+20AC which is not representable in ISO-8859-1.
        assert_eq!(super::latin1_encode_if_possible("€"), None);
    }

    #[test]
    fn decoded_pixbuf_len_computes_rowstride_times_height() {
        let len = decoded_pixbuf_len(400, 200).expect("expected valid rowstride/height");
        assert_eq!(len, 400 * 200);
        assert!(len < MAX_DECODED_IMAGE_BYTES);
    }

    #[test]
    fn decoded_pixbuf_len_rejects_non_positive_inputs() {
        assert_eq!(decoded_pixbuf_len(0, 10), None);
        assert_eq!(decoded_pixbuf_len(10, 0), None);
        assert_eq!(decoded_pixbuf_len(-1, 10), None);
        assert_eq!(decoded_pixbuf_len(10, -1), None);
    }

    #[test]
    fn decoded_pixbuf_len_handles_extreme_values_portably() {
        let len = decoded_pixbuf_len(i32::MAX, i32::MAX);
        if usize::BITS <= 32 {
            assert_eq!(len, None, "32-bit usize should overflow the multiplication");
        } else {
            assert!(len.is_some(), "64-bit usize should be able to represent the product");
        }
    }

    #[test]
    fn decode_text_for_target_decodes_string_as_latin1() {
        // X11 `STRING` uses ISO-8859-1; a standalone 0xE9 byte should decode to "é".
        let bytes = [0xE9u8];
        let decoded = decode_text_for_target("STRING", &bytes);
        assert_eq!(decoded.as_deref(), Some("é"));
    }

    #[test]
    fn decode_text_for_target_allows_fallback_when_utf8_string_is_invalid() {
        // Some legacy apps advertise `UTF8_STRING` but provide invalid UTF-8. We must not turn
        // those bytes into U+FFFD if a better (e.g. `STRING`) target is available.
        let invalid_utf8 = [0xE9u8]; // Invalid UTF-8 when interpreted as a standalone byte.
        let latin1 = [0xE9u8];

        let decoded = [("UTF8_STRING", &invalid_utf8[..]), ("STRING", &latin1[..])]
            .into_iter()
            .find_map(|(target, bytes)| decode_text_for_target(target, bytes));

        assert_eq!(decoded.as_deref(), Some("é"));
    }

    #[test]
    fn decode_text_for_target_lossy_utf8_can_replace_invalid_sequences() {
        let invalid_utf8 = [0xE9u8];
        let decoded = decode_text_for_target_lossy_utf8("UTF8_STRING", &invalid_utf8);
        assert_eq!(decoded.as_deref(), Some("\u{FFFD}"));
    }

    #[test]
    fn target_prefers_utf8_identifies_utf8_like_targets() {
        assert!(target_prefers_utf8("UTF8_STRING"));
        assert!(target_prefers_utf8("text/plain"));
        assert!(!target_prefers_utf8("STRING"));
        assert!(!target_prefers_utf8("TEXT"));
    }

    #[test]
    fn choose_targets_in_order_returns_multiple_candidates_in_preference_order() {
        let targets = vec!["STRING", "UTF8_STRING", "text/plain;charset=utf-8"];
        let candidates =
            super::choose_targets_in_order(&targets, &["text/plain", "utf8_string", "string", "text"]);
        assert_eq!(candidates, vec!["text/plain;charset=utf-8", "UTF8_STRING", "STRING"]);
    }

    #[test]
    fn bytes_to_utf8_within_limit_drops_strings_that_expand_over_limit() {
        // Each invalid byte becomes U+FFFD (3 UTF-8 bytes) when decoded lossily.
        let bytes = [0xFF, 0xFF, 0xFF];
        assert_eq!(bytes_to_utf8_within_limit(&bytes, 3), None);
    }

    #[test]
    fn bytes_to_utf8_within_limit_accepts_valid_utf8_html_and_rtf() {
        let html = b"<p>Hello</p>";
        assert_eq!(
            bytes_to_utf8_within_limit(html, 64),
            Some("<p>Hello</p>".to_string())
        );

        let rtf = b"{\\rtf1\\ansi Hello}";
        assert_eq!(
            bytes_to_utf8_within_limit(rtf, 64),
            Some("{\\rtf1\\ansi Hello}".to_string())
        );
    }
}
