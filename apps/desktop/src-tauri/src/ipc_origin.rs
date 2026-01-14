use url::Url;

use crate::tauri_origin::{self, DesktopPlatform};

/// Returns whether a WebView origin is considered trusted for privileged IPC commands.
///
/// This is a defense-in-depth guard: even though Tauri's security model should prevent remote
/// origins from accessing the invoke API by default, we keep privileged commands resilient in case
/// remote content is ever loaded into a WebView.
///
/// We treat app-local content as trusted:
/// - production: `tauri://localhost` (macOS/Linux) or `http(s)://tauri.localhost` (Windows)
/// - dev: the configured `build.devUrl` origin (typically `http://localhost:<port>`)
///
/// Loopback variants (`127.0.0.1`, `::1`) are also allowed.
///
/// NOTE:
/// Prefer [`matches_stable_webview_origin`] for privileged IPC commands. This helper is intentionally
/// conservative (host-based) and does **not** encode the full expected origin for the current
/// build/config (e.g. the dev server port).
pub fn is_trusted_app_origin(url: &Url) -> bool {
    match url.scheme() {
        "data" => return false,
        // `file://` is explicitly *not* trusted by default. If the WebView ever navigates to local
        // HTML, we don't want that origin to be able to invoke privileged commands.
        //
        // In debug builds only: set `FORMULA_TRUST_FILE_IPC_ORIGIN=1` to opt into trusting
        // `file://` origins (primarily for local/dev debugging).
        "file" => return trust_file_ipc_origin_enabled(),
        _ => {}
    }

    match url.host() {
        Some(url::Host::Domain(host)) => {
            // `localhost` is used for dev-server origins (`build.devUrl`) across platforms.
            host == "localhost"
                // On Windows, WebView2 maps custom schemes (like `tauri:`) onto
                // `http(s)://<scheme>.localhost`. Restrict this to the specific host we expect
                // (Tauri's internal scheme is `tauri:` => `tauri.localhost`).
                || host == "tauri.localhost"
        }
        Some(url::Host::Ipv4(ip)) => ip == std::net::Ipv4Addr::LOCALHOST,
        Some(url::Host::Ipv6(ip)) => ip == std::net::Ipv6Addr::LOCALHOST,
        None => false,
    }
}

fn trust_file_ipc_origin_enabled() -> bool {
    // Never trust `file://` in release builds: local HTML should not be able to invoke privileged
    // IPC in production.
    if !cfg!(debug_assertions) {
        return false;
    }
    match std::env::var("FORMULA_TRUST_FILE_IPC_ORIGIN") {
        Ok(raw) => {
            let v = raw.trim().to_ascii_lowercase();
            !(v.is_empty() || v == "0" || v == "false")
        }
        Err(_) => false,
    }
}

/// Returns `true` if `url` matches the expected, stable webview origin for this build/config.
///
/// This ties privileged IPC authorization to the *specific* origin Tauri expects for the WebView,
/// instead of broad host allowlists like `*.localhost`.
///
/// The expected origin is computed from the app's runtime config using
/// [`tauri_origin::stable_webview_origin`], and compared against the current window URL's origin
/// computed via [`tauri_origin::webview_origin_from_url`].
pub fn matches_stable_webview_origin(
    url: &Url,
    is_dev: bool,
    dev_url: Option<&str>,
    use_https_scheme: bool,
    platform: DesktopPlatform,
) -> bool {
    let expected =
        tauri_origin::stable_webview_origin(is_dev, dev_url, use_https_scheme, platform);
    matches_webview_origin(url, &expected, use_https_scheme, platform)
}

/// Returns `true` if `url` maps to `expected_origin` under Tauri's origin rules.
pub fn matches_webview_origin(
    url: &Url,
    expected_origin: &str,
    use_https_scheme: bool,
    platform: DesktopPlatform,
) -> bool {
    // Fail closed: if we can't compute a stable origin, do not treat the "null" origin as trusted.
    if expected_origin == "null" || expected_origin.trim().is_empty() {
        return false;
    }
    tauri_origin::webview_origin_from_url(url, use_https_scheme, platform) == expected_origin
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Verb {
    Is,
    Are,
}

impl Verb {
    fn as_str(self) -> &'static str {
        match self {
            Verb::Is => "is",
            Verb::Are => "are",
        }
    }
}

pub fn ensure_main_window(label: &str, subject: &str, verb: Verb) -> Result<(), String> {
    if label != "main" {
        return Err(format!(
            "{subject} {} only allowed from the main window",
            verb.as_str()
        ));
    }
    Ok(())
}

pub fn ensure_trusted_origin(url: &Url, subject: &str, verb: Verb) -> Result<(), String> {
    if !is_trusted_app_origin(url) {
        return Err(format!(
            "{subject} {} not allowed from this origin",
            verb.as_str()
        ));
    }
    Ok(())
}

#[cfg(feature = "desktop")]
fn use_https_scheme_for_window(window: &tauri::WebviewWindow) -> bool {
    window
        .app_handle()
        .config()
        .app
        .windows
        .iter()
        .find(|w| w.label == window.label())
        .map(|w| w.use_https_scheme)
        .unwrap_or(false)
}

#[cfg(feature = "desktop")]
fn dev_url_from_config(config: &tauri::Config) -> Option<&str> {
    // `dev_url` is represented as either `String` or `Url` depending on the Tauri version.
    // Both expose `as_str()`, so prefer that over `as_deref()` to avoid tying this code to a
    // specific config representation.
    config.build.dev_url.as_ref().map(|url| url.as_str())
}

/// Enforce that the invoking window is on the expected, stable webview origin.
///
/// Prefer this over [`ensure_trusted_origin`].
#[cfg(feature = "desktop")]
pub fn ensure_stable_origin(
    window: &tauri::WebviewWindow,
    subject: &str,
    verb: Verb,
) -> Result<(), String> {
    let url = window.url().map_err(|err| err.to_string())?;
    let config = window.app_handle().config();

    let platform = DesktopPlatform::current();
    let use_https_scheme = use_https_scheme_for_window(window);
    let dev_url = dev_url_from_config(&config);

    if !matches_stable_webview_origin(
        &url,
        tauri::is_dev(),
        dev_url,
        use_https_scheme,
        platform,
    ) {
        return Err(format!(
            "{subject} {} not allowed from this origin",
            verb.as_str()
        ));
    }
    Ok(())
}

#[cfg(feature = "desktop")]
pub fn ensure_main_window_and_stable_origin(
    window: &tauri::WebviewWindow,
    subject: &str,
    verb: Verb,
) -> Result<(), String> {
    ensure_main_window(window.label(), subject, verb)?;
    ensure_stable_origin(window, subject, verb)?;
    Ok(())
}

pub fn ensure_main_window_and_trusted_origin(
    label: &str,
    url: &Url,
    subject: &str,
    verb: Verb,
) -> Result<(), String> {
    ensure_main_window(label, subject, verb)?;
    ensure_trusted_origin(url, subject, verb)?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{is_trusted_app_origin, matches_stable_webview_origin};
    use crate::tauri_origin::DesktopPlatform;
    use url::Url;

    #[test]
    fn allows_localhost() {
        let url = Url::parse("http://localhost:1420/").unwrap();
        assert!(is_trusted_app_origin(&url));
    }

    #[test]
    fn allows_windows_tauri_localhost() {
        let url = Url::parse("https://tauri.localhost/").unwrap();
        assert!(is_trusted_app_origin(&url));
    }

    #[test]
    fn denies_arbitrary_localhost_subdomains() {
        let url = Url::parse("https://evil.localhost/").unwrap();
        assert!(!is_trusted_app_origin(&url));

        let url = Url::parse("https://foo.bar.localhost/some/path").unwrap();
        assert!(!is_trusted_app_origin(&url));
    }

    #[test]
    fn allows_loopback_hosts() {
        let url = Url::parse("http://127.0.0.1:1234/").unwrap();
        assert!(is_trusted_app_origin(&url));

        let url = Url::parse("http://[::1]:1234/").unwrap();
        assert!(is_trusted_app_origin(&url));
    }

    #[test]
    fn denies_file_scheme_by_default() {
        let url = Url::parse("file:///foo/bar.txt").unwrap();
        assert!(!is_trusted_app_origin(&url));
    }

    #[test]
    fn denies_data_scheme() {
        let url = Url::parse("data:text/plain,hello").unwrap();
        assert!(!is_trusted_app_origin(&url));
    }

    #[test]
    fn denies_remote_hosts() {
        let url = Url::parse("https://example.com/").unwrap();
        assert!(!is_trusted_app_origin(&url));
    }

    #[test]
    fn denies_empty_host() {
        // Valid URL with a scheme but no host.
        let url = Url::parse("http:///path").unwrap();
        assert!(!is_trusted_app_origin(&url));
    }

    #[test]
    fn stable_origin_dev_allows_configured_localhost_port() {
        let window_url = Url::parse("http://localhost:4174/some/path").unwrap();
        assert!(matches_stable_webview_origin(
            &window_url,
            true,
            Some("http://localhost:4174/"),
            false,
            DesktopPlatform::Linux
        ));

        let other_port = Url::parse("http://localhost:9999/").unwrap();
        assert!(!matches_stable_webview_origin(
            &other_port,
            true,
            Some("http://localhost:4174/"),
            false,
            DesktopPlatform::Linux
        ));
    }

    #[test]
    fn stable_origin_prod_unix_is_tauri_localhost() {
        let window_url = Url::parse("tauri://localhost/index.html").unwrap();
        assert!(matches_stable_webview_origin(
            &window_url,
            false,
            None,
            false,
            DesktopPlatform::Linux
        ));
    }

    #[test]
    fn stable_origin_prod_windows_accepts_tauri_localhost_http_and_https() {
        // WebView2 reports the internal `tauri:` origin as `http(s)://tauri.localhost`.
        let http_url = Url::parse("http://tauri.localhost/index.html").unwrap();
        assert!(matches_stable_webview_origin(
            &http_url,
            false,
            None,
            false,
            DesktopPlatform::Windows
        ));

        let https_url = Url::parse("https://tauri.localhost/index.html").unwrap();
        assert!(matches_stable_webview_origin(
            &https_url,
            false,
            None,
            true,
            DesktopPlatform::Windows
        ));

        // Some code paths may still surface the underlying custom-scheme URL.
        let tauri_url = Url::parse("tauri://localhost/index.html").unwrap();
        assert!(matches_stable_webview_origin(
            &tauri_url,
            false,
            None,
            false,
            DesktopPlatform::Windows
        ));
        assert!(matches_stable_webview_origin(
            &tauri_url,
            false,
            None,
            true,
            DesktopPlatform::Windows
        ));
    }
}
