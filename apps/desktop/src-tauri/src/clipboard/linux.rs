#[cfg(feature = "desktop")]
use super::{ClipboardContent, ClipboardError, ClipboardWritePayload};

fn normalize_target_name(target: &str) -> String {
    target.trim().to_ascii_lowercase()
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

#[cfg(feature = "desktop")]
fn bytes_to_utf8(bytes: &[u8]) -> Option<String> {
    if bytes.is_empty() {
        return None;
    }
    // Some clipboard sources include NUL termination.
    let s = String::from_utf8_lossy(bytes);
    let s = s.trim_end_matches('\0');
    if s.is_empty() {
        None
    } else {
        Some(s.to_string())
    }
}

#[cfg(feature = "desktop")]
mod gtk_backend {
    use base64::{engine::general_purpose::STANDARD, Engine as _};

    use super::super::{normalize_base64_str, MAX_IMAGE_BYTES, MAX_RICH_TEXT_BYTES};
    use super::{
        bytes_to_utf8, choose_best_target, ClipboardContent, ClipboardError, ClipboardWritePayload,
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
        for target in targets {
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
            if let Some(s) = bytes_to_utf8(&bytes) {
                return Some(s);
            }
        }
        None
    }

    fn wait_for_bytes_base64(
        clipboard: &gtk::Clipboard,
        target: &str,
        max_bytes: usize,
    ) -> Option<String> {
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
            Some(STANDARD.encode(bytes))
        }
    }

    pub(super) fn read() -> Result<ClipboardContent, ClipboardError> {
        with_gtk_main_thread(|| {
            ensure_gtk()?;

            let read_from_clipboard = |clipboard: &gtk::Clipboard| {
                let targets = clipboard_target_names(clipboard);

                // `wait_for_text` covers many common plaintext targets (UTF8_STRING, STRING, TEXT),
                // but some producers only advertise MIME-like targets (e.g. `text/plain;charset=utf-8`).
                // Fall back to reading `text/plain*` via the same target-selection logic used for
                // rich formats so we don't incorrectly treat the clipboard as empty and fall back
                // to PRIMARY.
                let text = clipboard
                    .wait_for_text()
                    .map(|s| s.to_string())
                    .filter(|s| !s.is_empty())
                    .or_else(|| match targets.as_deref() {
                        Some(targets) => choose_best_target(targets, &["text/plain"])
                            .and_then(|t| wait_for_utf8_targets(clipboard, &[t], usize::MAX)),
                        None => wait_for_utf8_targets(
                            clipboard,
                            &["text/plain", "text/plain;charset=utf-8"],
                            usize::MAX,
                        ),
                    });
                let html = match targets.as_deref() {
                    Some(targets) => choose_best_target(targets, &["text/html"])
                        .and_then(|t| wait_for_utf8_targets(clipboard, &[t], MAX_RICH_TEXT_BYTES)),
                    // If target enumeration isn't available, fall back to the canonical target.
                    None => wait_for_utf8_targets(
                        clipboard,
                        &["text/html", "text/html;charset=utf-8"],
                        MAX_RICH_TEXT_BYTES,
                    ),
                };
                let rtf = match targets.as_deref() {
                    Some(targets) => choose_best_target(
                        targets,
                        &["text/rtf", "application/rtf", "application/x-rtf"],
                    )
                    .and_then(|t| wait_for_utf8_targets(clipboard, &[t], MAX_RICH_TEXT_BYTES)),
                    None => wait_for_utf8_targets(
                        clipboard,
                        &[
                            "text/rtf",
                            "text/rtf;charset=utf-8",
                            "application/rtf",
                            "application/x-rtf",
                        ],
                        MAX_RICH_TEXT_BYTES,
                    ),
                };
                let png_base64 = match targets.as_deref() {
                    Some(targets) => choose_best_target(targets, &["image/png"])
                        .and_then(|t| wait_for_bytes_base64(clipboard, t, MAX_IMAGE_BYTES)),
                    None => wait_for_bytes_base64(clipboard, "image/png", MAX_IMAGE_BYTES),
                }
                .or_else(|| {
                    // Some applications expose images on the clipboard without an `image/png` target.
                    // Fall back to GTK's pixbuf API and re-encode to PNG (requires image loaders).
                    let pixbuf = clipboard.wait_for_image()?;
                    let bytes = pixbuf.save_to_bufferv("png", &[]).ok()?;
                    if bytes.len() > MAX_IMAGE_BYTES {
                        return None;
                    }
                    Some(STANDARD.encode(bytes))
                });

                ClipboardContent {
                    text,
                    html,
                    rtf,
                    png_base64,
                }
            };

            let clipboard = gtk::Clipboard::get(&gdk::SELECTION_CLIPBOARD);
            let content = read_from_clipboard(&clipboard);

            // On X11, some apps only populate PRIMARY selection (middle-click paste).
            // Only fall back to PRIMARY when CLIPBOARD has no usable content, and skip on Wayland to
            // avoid changing semantics where PRIMARY may not exist or behave differently.
            let has_usable_data = clipboard_fallback::has_usable_clipboard_data(
                content.text.as_deref(),
                content.html.as_deref(),
                content.rtf.as_deref(),
                content.png_base64.as_deref(),
            );
            if !has_usable_data && clipboard_fallback::should_attempt_primary_selection_from_env() {
                let primary = gtk::Clipboard::get(&gdk::SELECTION_PRIMARY);
                return Ok(read_from_clipboard(&primary));
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
            .png_base64
            .as_deref()
            .map(normalize_base64_str)
            .filter(|s| !s.is_empty())
            .map(|s| {
                STANDARD
                    .decode(s)
                    .map_err(|e| ClipboardError::InvalidPayload(format!("invalid pngBase64: {e}")))
            })
            .transpose()?
            .map(Arc::new);

        with_gtk_main_thread(move || {
            ensure_gtk()?;

            let set_clipboard_data = |clipboard: &gtk::Clipboard| {
                const INFO_TEXT: u32 = 1;
                const INFO_HTML: u32 = 2;
                const INFO_RTF: u32 = 3;
                const INFO_PNG: u32 = 4;

                // Do not restrict targets based on app identity; we want both intra-app copy/paste and
                // interoperability with other apps (LibreOffice, browser, image editors, etc.).
                let flags = gtk::TargetFlags::empty();
                let mut targets: Vec<gtk::TargetEntry> = Vec::new();

                if text.is_some() {
                    // Provide common plaintext targets so other apps can pick their preferred flavor.
                    // (GTK's own `set_text` would also do this, but we need to set multiple targets at once.)
                    targets.push(gtk::TargetEntry::new("text/plain", flags, INFO_TEXT));
                    targets.push(gtk::TargetEntry::new(
                        "text/plain;charset=utf-8",
                        flags,
                        INFO_TEXT,
                    ));
                    targets.push(gtk::TargetEntry::new("UTF8_STRING", flags, INFO_TEXT));
                    targets.push(gtk::TargetEntry::new("STRING", flags, INFO_TEXT));
                    targets.push(gtk::TargetEntry::new("TEXT", flags, INFO_TEXT));
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

                // Note: the closure captures owned copies of the strings/bytes so the clipboard stays
                // valid after this function returns.
                let text = text.clone();
                let html = html.clone();
                let rtf = rtf.clone();
                let png_bytes = png_bytes.clone();
                let success =
                    clipboard.set_with_data(&targets, move |_clipboard, selection_data, info| {
                        match info {
                            INFO_TEXT => {
                                if let Some(ref text) = text {
                                    selection_data.set(
                                        &selection_data.target(),
                                        8,
                                        text.as_bytes(),
                                    );
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

#[cfg(test)]
mod tests {
    use super::{choose_best_target, normalize_target_name};

    #[test]
    fn normalize_target_name_lowercases_and_trims() {
        assert_eq!(
            normalize_target_name("  Text/Html;Charset=UTF-8 \t"),
            "text/html;charset=utf-8"
        );
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
}
