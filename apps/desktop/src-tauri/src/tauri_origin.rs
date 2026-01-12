use url::Url;

/// Minimal platform model for computing Tauri webview origins.
///
/// We keep this in the `desktop` library crate (instead of the Tauri binary)
/// so it can be unit tested without enabling the `desktop` feature (which
/// pulls in the system WebView toolchain on Linux).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DesktopPlatform {
    Windows,
    MacOS,
    Linux,
    Other,
}

impl DesktopPlatform {
    pub const fn current() -> Self {
        #[cfg(target_os = "windows")]
        return Self::Windows;
        #[cfg(target_os = "macos")]
        return Self::MacOS;
        #[cfg(target_os = "linux")]
        return Self::Linux;
        #[cfg(not(any(target_os = "windows", target_os = "macos", target_os = "linux")))]
        return Self::Other;
    }
}

/// Computes the stable CORS origin for a Tauri webview.
///
/// This intentionally mirrors Tauri's upstream `window_origin` behavior:
/// the origin is derived from the *initial* webview URL (not the current one),
/// so a later navigation cannot gain CORS access to privileged `asset://`
/// resources.
///
/// Desktop rules (Tauri v2):
/// - dev: origin comes from `build.devUrl` (including the port)
/// - prod: origin is `tauri://localhost` on macOS/Linux; `http(s)://tauri.localhost` on Windows
pub fn stable_webview_origin(
    is_dev: bool,
    dev_url: Option<&str>,
    use_https_scheme: bool,
    platform: DesktopPlatform,
) -> String {
    let initial_webview_url = if is_dev {
        dev_url.and_then(|raw| Url::parse(raw).ok())
    } else {
        // Packaged builds always use the internal Tauri origin.
        Url::parse("tauri://localhost").ok()
    };

    let Some(initial_webview_url) = initial_webview_url else {
        return "null".to_string();
    };

    if initial_webview_url.scheme() == "data" {
        return "null".to_string();
    }

    if platform == DesktopPlatform::Windows
        && initial_webview_url.scheme() != "http"
        && initial_webview_url.scheme() != "https"
    {
        // On Windows, WebView2 exposes custom schemes (like `tauri:`) as
        // `http(s)://<scheme>.localhost`.
        let scheme = if use_https_scheme { "https" } else { "http" };
        return format!("{scheme}://{}.localhost", initial_webview_url.scheme());
    }

    if let Some(host) = initial_webview_url.host() {
        return format!(
            "{}://{}{}",
            initial_webview_url.scheme(),
            host,
            initial_webview_url
                .port()
                .map(|p| format!(":{p}"))
                .unwrap_or_default()
        );
    }

    "null".to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn stable_webview_origin_dev_uses_dev_url_origin() {
        let origin = stable_webview_origin(
            true,
            Some("http://localhost:4174/some/path?x=1#y"),
            false,
            DesktopPlatform::Linux,
        );
        assert_eq!(origin, "http://localhost:4174");
    }

    #[test]
    fn stable_webview_origin_prod_unix_is_tauri_localhost() {
        let origin = stable_webview_origin(false, None, false, DesktopPlatform::Linux);
        assert_eq!(origin, "tauri://localhost");
    }

    #[test]
    fn stable_webview_origin_prod_windows_honors_use_https_scheme() {
        let http_origin = stable_webview_origin(false, None, false, DesktopPlatform::Windows);
        assert_eq!(http_origin, "http://tauri.localhost");

        let https_origin = stable_webview_origin(false, None, true, DesktopPlatform::Windows);
        assert_eq!(https_origin, "https://tauri.localhost");
    }
}

