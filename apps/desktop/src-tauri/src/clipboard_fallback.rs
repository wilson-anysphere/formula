use std::ffi::OsStr;

pub(crate) fn has_usable_clipboard_data(
    text: Option<&str>,
    html: Option<&str>,
    rtf: Option<&str>,
    image_png_base64: Option<&str>,
) -> bool {
    [text, html, rtf, image_png_base64]
        .into_iter()
        .any(|v| matches!(v, Some(s) if !s.is_empty()))
}

pub(crate) fn should_attempt_primary_selection(
    xdg_session_type: Option<&str>,
    wayland_display: Option<&OsStr>,
) -> bool {
    // PRIMARY selection is an X11 concept. On Wayland, the protocol may not be available (or may
    // behave differently), so avoid changing behavior by default.
    if wayland_display.is_some() {
        return false;
    }
    if let Some(session_type) = xdg_session_type {
        if session_type.eq_ignore_ascii_case("wayland") {
            return false;
        }
    }
    true
}

#[cfg(feature = "desktop")]
pub(crate) fn should_attempt_primary_selection_from_env() -> bool {
    let xdg_session_type = std::env::var("XDG_SESSION_TYPE").ok();
    let wayland_display = std::env::var_os("WAYLAND_DISPLAY");
    should_attempt_primary_selection(xdg_session_type.as_deref(), wayland_display.as_deref())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn has_usable_clipboard_data_requires_non_empty_value() {
        assert!(!has_usable_clipboard_data(None, None, None, None));
        assert!(!has_usable_clipboard_data(Some(""), None, None, None));
        assert!(!has_usable_clipboard_data(None, Some(""), None, None));
        assert!(has_usable_clipboard_data(Some("hello"), None, None, None));
        assert!(has_usable_clipboard_data(
            None,
            None,
            Some("{\\rtf1}"),
            None
        ));
        assert!(has_usable_clipboard_data(
            None,
            None,
            None,
            Some("aGVsbG8=")
        ));
    }

    #[test]
    fn should_attempt_primary_selection_skips_wayland_sessions() {
        assert!(!should_attempt_primary_selection(
            Some("wayland"),
            Some(OsStr::new("wayland-0"))
        ));
        assert!(!should_attempt_primary_selection(Some("wayland"), None));
        assert!(!should_attempt_primary_selection(
            None,
            Some(OsStr::new("wayland-0"))
        ));
    }

    #[test]
    fn should_attempt_primary_selection_allows_x11() {
        assert!(should_attempt_primary_selection(Some("x11"), None));
        // Some environments don't set XDG_SESSION_TYPE; treat that as "not Wayland".
        assert!(should_attempt_primary_selection(None, None));
    }
}
