use url::Url;

use crate::ipc_origin;
use crate::tauri_origin::DesktopPlatform;

/// Returns whether the `asset://` protocol should be available for a given window URL.
///
/// Security model:
/// - `asset://` serves local files within the configured Tauri scope.
/// - Only trusted app-local origins should be able to access it.
/// - If the webview navigates to an untrusted origin (remote content), it must not be able to
///   continue reading local files through `asset://`.
///
/// The decision is derived from:
/// - `stable_origin`: the "stable" app origin (computed once from config/platform rules), and
/// - `window_url`: the *current* window URL.
///
/// The stable origin is intentionally not derived from the current URL so that an attacker cannot
/// change the effective CORS policy via navigation.
pub fn is_asset_protocol_allowed(
    stable_origin: &str,
    window_url: Option<&Url>,
    use_https_scheme: bool,
    platform: DesktopPlatform,
) -> bool {
    let Some(window_url) = window_url else {
        return false;
    };

    ipc_origin::matches_webview_origin(window_url, stable_origin, use_https_scheme, platform)
}

#[cfg(test)]
mod tests {
    use super::is_asset_protocol_allowed;
    use crate::tauri_origin::DesktopPlatform;
    use url::Url;

    #[test]
    fn allows_tauri_localhost_prod_unix() {
        let stable_origin = "tauri://localhost";
        let window_url = Url::parse("tauri://localhost/index.html").unwrap();
        assert!(is_asset_protocol_allowed(
            stable_origin,
            Some(&window_url),
            false,
            DesktopPlatform::Linux,
        ));
    }

    #[test]
    fn allows_tauri_localhost_mapped_origin_windows_prod() {
        let stable_origin = "http://tauri.localhost";
        let window_url = Url::parse("http://tauri.localhost/index.html").unwrap();
        assert!(is_asset_protocol_allowed(
            stable_origin,
            Some(&window_url),
            false,
            DesktopPlatform::Windows,
        ));
    }

    #[test]
    fn allows_dev_origin_exact_match() {
        let stable_origin = "http://localhost:4174";
        let window_url = Url::parse("http://localhost:4174/some/path").unwrap();
        assert!(is_asset_protocol_allowed(
            stable_origin,
            Some(&window_url),
            false,
            DesktopPlatform::Linux,
        ));
    }

    #[test]
    fn denies_remote_origin() {
        let stable_origin = "tauri://localhost";
        let window_url = Url::parse("https://example.com/some/path").unwrap();
        assert!(!is_asset_protocol_allowed(
            stable_origin,
            Some(&window_url),
            false,
            DesktopPlatform::Linux,
        ));
    }

    #[test]
    fn denies_localhost_port_mismatch() {
        let stable_origin = "http://localhost:4174";
        let window_url = Url::parse("http://localhost:9999/").unwrap();
        assert!(!is_asset_protocol_allowed(
            stable_origin,
            Some(&window_url),
            false,
            DesktopPlatform::Linux,
        ));
    }

    #[test]
    fn denies_file_scheme_even_if_ipc_origin_allows_it() {
        let stable_origin = "tauri://localhost";
        let window_url = Url::parse("file:///Users/alice/secrets.txt").unwrap();
        assert!(!is_asset_protocol_allowed(
            stable_origin,
            Some(&window_url),
            false,
            DesktopPlatform::Linux,
        ));
    }
}
