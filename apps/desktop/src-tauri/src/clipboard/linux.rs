use base64::{engine::general_purpose::STANDARD, Engine as _};

use super::{ClipboardContent, ClipboardError, ClipboardWritePayload};

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
    let (tx, rx) = std::sync::mpsc::channel::<Result<R, ClipboardError>>();

    ctx.invoke(move || {
        let res = op();
        let _ = tx.send(res);
    });

    rx.recv().map_err(|_| {
        ClipboardError::OperationFailed("failed to receive result from GTK main thread".to_string())
    })?
}

fn ensure_gtk() -> Result<(), ClipboardError> {
    gtk::init().map_err(|e| ClipboardError::Unavailable(e.to_string()))
}

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

fn wait_for_utf8_targets(clipboard: &gtk::Clipboard, targets: &[&str]) -> Option<String> {
    for target in targets {
        let atom = gdk::Atom::intern(target);
        let Some(data) = clipboard.wait_for_contents(&atom) else {
            continue;
        };
        if let Some(s) = bytes_to_utf8(&data.data()) {
            return Some(s);
        }
    }
    None
}

fn wait_for_bytes_base64(clipboard: &gtk::Clipboard, target: &str) -> Option<String> {
    let atom = gdk::Atom::intern(target);
    let data = clipboard.wait_for_contents(&atom)?;
    let bytes = data.data();
    if bytes.is_empty() {
        None
    } else {
        Some(STANDARD.encode(bytes))
    }
}

pub fn read() -> Result<ClipboardContent, ClipboardError> {
    with_gtk_main_thread(|| {
        ensure_gtk()?;

        let clipboard = gtk::Clipboard::get(&gdk::SELECTION_CLIPBOARD);

        let text = clipboard.wait_for_text().map(|s| s.to_string());
        let html = wait_for_utf8_targets(&clipboard, &["text/html", "text/html;charset=utf-8"]);
        let rtf = wait_for_utf8_targets(&clipboard, &["text/rtf", "application/rtf"]);
        let image_png_base64 = wait_for_bytes_base64(&clipboard, "image/png");

        Ok(ClipboardContent {
            text,
            html,
            rtf,
            image_png_base64,
        })
    })
}

pub fn write(payload: &ClipboardWritePayload) -> Result<(), ClipboardError> {
    let text = payload.text.clone();
    let html = payload.html.clone();
    let rtf = payload.rtf.clone();
    let png_bytes = payload
        .image_png_base64
        .as_deref()
        .filter(|s| !s.is_empty())
        .map(|s| {
            STANDARD
                .decode(s)
                .map_err(|e| ClipboardError::InvalidPayload(format!("invalid pngBase64: {e}")))
        })
        .transpose()?;

    with_gtk_main_thread(move || {
        ensure_gtk()?;

        let clipboard = gtk::Clipboard::get(&gdk::SELECTION_CLIPBOARD);

        const INFO_TEXT: u32 = 1;
        const INFO_HTML: u32 = 2;
        const INFO_RTF: u32 = 3;
        const INFO_PNG: u32 = 4;

        let flags = gtk::TargetFlags::OTHER_APP;
        let mut targets: Vec<gtk::TargetEntry> = Vec::new();

        // Provide common plaintext targets so other apps can pick their preferred flavor.
        // (GTK's own `set_text` would also do this, but we need to set multiple targets at once.)
        targets.push(gtk::TargetEntry::new("text/plain", flags, INFO_TEXT));
        targets.push(gtk::TargetEntry::new("text/plain;charset=utf-8", flags, INFO_TEXT));
        targets.push(gtk::TargetEntry::new("UTF8_STRING", flags, INFO_TEXT));
        targets.push(gtk::TargetEntry::new("STRING", flags, INFO_TEXT));
        targets.push(gtk::TargetEntry::new("TEXT", flags, INFO_TEXT));

        if html.is_some() {
            targets.push(gtk::TargetEntry::new("text/html", flags, INFO_HTML));
            targets.push(gtk::TargetEntry::new("text/html;charset=utf-8", flags, INFO_HTML));
        }

        if rtf.is_some() {
            targets.push(gtk::TargetEntry::new("text/rtf", flags, INFO_RTF));
            targets.push(gtk::TargetEntry::new("application/rtf", flags, INFO_RTF));
        }

        if png_bytes.is_some() {
            targets.push(gtk::TargetEntry::new("image/png", flags, INFO_PNG));
        }

        // Note: the closure captures owned copies of the strings/bytes so the clipboard stays
        // valid after this function returns.
        let success = clipboard.set_with_data(&targets, move |_clipboard, selection_data, info| {
            match info {
                INFO_TEXT => {
                    selection_data.set(&selection_data.target(), 8, text.as_bytes());
                }
                INFO_HTML => {
                    if let Some(ref html) = html {
                        selection_data.set(&selection_data.target(), 8, html.as_bytes());
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

        // Best-effort request to persist clipboard data via a clipboard manager (X11).
        clipboard.store();

        Ok(())
    })
}
