use std::ffi::OsStr;

#[cfg(any(feature = "desktop", test))]
const PRIMARY_SELECTION_OVERRIDE_ENV_VAR: &str = "FORMULA_CLIPBOARD_PRIMARY_SELECTION";

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

fn parse_primary_selection_override(value: &str) -> Option<bool> {
    let value = value.trim();
    if value == "1" || value.eq_ignore_ascii_case("true") || value.eq_ignore_ascii_case("yes") {
        return Some(true);
    }
    if value == "0" || value.eq_ignore_ascii_case("false") || value.eq_ignore_ascii_case("no") {
        return Some(false);
    }
    None
}

fn should_attempt_primary_selection_with_override(
    primary_selection_override: Option<&str>,
    xdg_session_type: Option<&str>,
    wayland_display: Option<&OsStr>,
) -> bool {
    primary_selection_override
        .and_then(parse_primary_selection_override)
        .unwrap_or_else(|| should_attempt_primary_selection(xdg_session_type, wayland_display))
}

#[cfg(any(feature = "desktop", test))]
pub(crate) fn should_attempt_primary_selection_from_env() -> bool {
    let primary_selection_override = std::env::var(PRIMARY_SELECTION_OVERRIDE_ENV_VAR).ok();
    let xdg_session_type = std::env::var("XDG_SESSION_TYPE").ok();
    let wayland_display = std::env::var_os("WAYLAND_DISPLAY");
    should_attempt_primary_selection_with_override(
        primary_selection_override.as_deref(),
        xdg_session_type.as_deref(),
        wayland_display.as_deref(),
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::ffi::OsString;
    use std::sync::{Mutex, OnceLock};

    static ENV_MUTEX: OnceLock<Mutex<()>> = OnceLock::new();

    fn env_mutex() -> &'static Mutex<()> {
        ENV_MUTEX.get_or_init(|| Mutex::new(()))
    }

    struct EnvVarGuard {
        key: &'static str,
        prev: Option<OsString>,
    }

    impl EnvVarGuard {
        fn set(key: &'static str, value: impl AsRef<std::ffi::OsStr>) -> Self {
            let prev = std::env::var_os(key);
            std::env::set_var(key, value);
            Self { key, prev }
        }

        fn remove(key: &'static str) -> Self {
            let prev = std::env::var_os(key);
            std::env::remove_var(key);
            Self { key, prev }
        }
    }

    impl Drop for EnvVarGuard {
        fn drop(&mut self) {
            match &self.prev {
                Some(prev) => std::env::set_var(self.key, prev),
                None => std::env::remove_var(self.key),
            }
        }
    }

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

    #[test]
    fn parse_primary_selection_override_accepts_disable_values() {
        for value in ["0", "false", "FALSE", "no", "No", "  no\t"] {
            assert_eq!(
                parse_primary_selection_override(value),
                Some(false),
                "unexpected parse result for {value:?}"
            );
        }
    }

    #[test]
    fn parse_primary_selection_override_accepts_enable_values() {
        for value in ["1", "true", "TRUE", "yes", "YeS", "  yes\t"] {
            assert_eq!(
                parse_primary_selection_override(value),
                Some(true),
                "unexpected parse result for {value:?}"
            );
        }
    }

    #[test]
    fn parse_primary_selection_override_rejects_unknown_values() {
        for value in ["", "maybe", "2", "enable", "disable"] {
            assert_eq!(
                parse_primary_selection_override(value),
                None,
                "unexpected parse result for {value:?}"
            );
        }
    }

    #[test]
    fn should_attempt_primary_selection_with_override_respects_explicit_values() {
        let wayland = Some(OsStr::new("wayland-0"));
        assert!(!should_attempt_primary_selection(Some("wayland"), wayland));

        // Allow PRIMARY explicitly even in Wayland sessions.
        assert!(should_attempt_primary_selection_with_override(
            Some("1"),
            Some("wayland"),
            wayland
        ));

        // Deny PRIMARY explicitly even on X11.
        assert!(!should_attempt_primary_selection_with_override(
            Some("no"),
            Some("x11"),
            None
        ));
    }

    #[test]
    fn should_attempt_primary_selection_from_env_honors_disable_override() {
        let _lock = env_mutex()
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let _override_var = EnvVarGuard::set(PRIMARY_SELECTION_OVERRIDE_ENV_VAR, "0");
        let _session = EnvVarGuard::set("XDG_SESSION_TYPE", "x11");
        let _wayland = EnvVarGuard::remove("WAYLAND_DISPLAY");

        assert!(!should_attempt_primary_selection_from_env());
    }

    #[test]
    fn should_attempt_primary_selection_from_env_honors_enable_override_even_on_wayland() {
        let _lock = env_mutex()
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let _override_var = EnvVarGuard::set(PRIMARY_SELECTION_OVERRIDE_ENV_VAR, "yes");
        let _session = EnvVarGuard::set("XDG_SESSION_TYPE", "wayland");
        let _wayland = EnvVarGuard::set("WAYLAND_DISPLAY", "wayland-0");

        assert!(should_attempt_primary_selection_from_env());
    }

    #[test]
    fn should_attempt_primary_selection_from_env_defaults_to_heuristic_when_unset() {
        let _lock = env_mutex()
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());

        let _override_var = EnvVarGuard::remove(PRIMARY_SELECTION_OVERRIDE_ENV_VAR);
        let _session_wayland = EnvVarGuard::set("XDG_SESSION_TYPE", "wayland");
        let _wayland = EnvVarGuard::remove("WAYLAND_DISPLAY");

        assert!(!should_attempt_primary_selection_from_env());

        // Switch to X11 and verify heuristic re-enables.
        let _session_x11 = EnvVarGuard::set("XDG_SESSION_TYPE", "x11");
        assert!(should_attempt_primary_selection_from_env());
    }
}
