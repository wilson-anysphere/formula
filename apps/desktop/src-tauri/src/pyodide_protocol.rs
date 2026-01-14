use std::path::{Component, Path, PathBuf};

use tauri::http::{
    header::CONTENT_TYPE,
    Request, Response, StatusCode,
};
use tauri::{Manager, Runtime, UriSchemeContext};
use url::Url;

/// Custom `pyodide:` protocol handler for serving cached Pyodide assets under COOP/COEP.
///
/// Security model:
/// - Only serves files from the app-controlled Pyodide cache directory.
/// - Only available to trusted app-local webview origins (mirrors `asset:` protocol gating).
///
/// Why:
/// - Packaged desktop builds run with `Cross-Origin-Embedder-Policy: require-corp` so Chromium
///   enables `SharedArrayBuffer` (required by Pyodide's worker backend).
/// - Pyodide assets must therefore be served with `Cross-Origin-Resource-Policy: cross-origin`
///   so they can be embedded as cross-origin subresources.
pub fn handler<R: Runtime>(
    ctx: UriSchemeContext<'_, R>,
    request: Request<Vec<u8>>,
) -> Response<Vec<u8>> {
    let window_origin = stable_window_origin(&ctx);

    // Security boundary: `pyodide://` serves local files within the Pyodide cache directory. If the
    // webview ever navigates to remote/untrusted content, do not allow that origin to read the
    // cache.
    let window_url = current_window_url(&ctx);
    let is_trusted_origin = desktop::asset_protocol_policy::is_asset_protocol_allowed(
        &window_origin,
        window_url.as_ref(),
        use_https_scheme(&ctx),
        desktop::tauri_origin::DesktopPlatform::current(),
    );
    if !is_trusted_origin {
        let url_for_log = window_url
            .as_ref()
            .map(|u| u.to_string())
            .unwrap_or_else(|| "<unknown>".to_string());
        eprintln!("[pyodide protocol] blocked request from untrusted origin: {url_for_log}");
        return Response::builder()
            .status(StatusCode::FORBIDDEN)
            .header(CONTENT_TYPE, "text/plain")
            .header("Access-Control-Allow-Origin", &window_origin)
            .header("Cross-Origin-Resource-Policy", "cross-origin")
            .body(b"pyodide protocol is only available from trusted app-local origins".to_vec())
            .unwrap();
    }

    let cache_root = match desktop::pyodide_assets::pyodide_cache_root() {
        Ok(p) => p,
        Err(err) => {
            return Response::builder()
                .status(StatusCode::INTERNAL_SERVER_ERROR)
                .header(CONTENT_TYPE, "text/plain")
                .header("Access-Control-Allow-Origin", &window_origin)
                .header("Cross-Origin-Resource-Policy", "cross-origin")
                .body(format!("failed to determine pyodide cache directory: {err}").into_bytes())
                .unwrap();
        }
    };

    get_response(request, &cache_root, &window_origin)
}

fn current_window_url<R: Runtime>(ctx: &UriSchemeContext<'_, R>) -> Option<Url> {
    let window = ctx.app_handle().get_webview_window(ctx.webview_label())?;
    window.as_ref().url().ok()
}

fn use_https_scheme<R: Runtime>(ctx: &UriSchemeContext<'_, R>) -> bool {
    ctx.app_handle()
        .config()
        .app
        .windows
        .iter()
        .find(|w| w.label == ctx.webview_label())
        .map(|w| w.use_https_scheme)
        .unwrap_or(false)
}

fn stable_window_origin<R: Runtime>(ctx: &UriSchemeContext<'_, R>) -> String {
    // Mirror the logic used by the `asset:` protocol handler: compute the stable origin from the
    // app config + platform so a compromised navigation cannot change the effective CORS policy.
    let config = ctx.app_handle().config();

    let use_https_scheme = use_https_scheme(ctx);

    let dev_url = config.build.dev_url.as_ref().map(|url| url.as_str());

    desktop::tauri_origin::stable_webview_origin(
        tauri::is_dev(),
        dev_url,
        use_https_scheme,
        desktop::tauri_origin::DesktopPlatform::current(),
    )
}

fn validate_relative_path(path: &str) -> bool {
    if path.is_empty() || path.contains('\0') {
        return false;
    }

    for c in Path::new(path).components() {
        match c {
            Component::Normal(_) => {}
            _ => return false,
        }
    }

    true
}

fn resolve_request_path(cache_root: &Path, request: &Request<Vec<u8>>) -> Result<PathBuf, String> {
    let uri = request.uri();
    let authority = uri
        .authority()
        .map(|a| a.as_str())
        .unwrap_or_default();
    let authority = authority
        .split('@')
        .last()
        .unwrap_or("")
        .split(':')
        .next()
        .unwrap_or("")
        .trim();

    if authority.is_empty() {
        return Err("missing pyodide:// authority".to_string());
    }

    // Skip leading `/` from the path.
    let raw_path = uri.path().strip_prefix('/').unwrap_or("");
    let decoded = percent_encoding::percent_decode(raw_path.as_bytes())
        .decode_utf8_lossy()
        .to_string();

    let rel = format!("{authority}/{decoded}");
    if !validate_relative_path(&rel) {
        return Err("invalid pyodide:// path".to_string());
    }

    Ok(cache_root.join(rel))
}

fn get_response(
    request: Request<Vec<u8>>,
    cache_root: &Path,
    window_origin: &str,
) -> Response<Vec<u8>> {
    use desktop::asset_protocol_core::{handle_asset_request, AssetHeaderValue, AssetMethod, AssetRequest};

    let method = if request.method() == tauri::http::Method::GET {
        AssetMethod::Get
    } else if request.method() == tauri::http::Method::HEAD {
        AssetMethod::Head
    } else if request.method() == tauri::http::Method::OPTIONS {
        AssetMethod::Options
    } else {
        AssetMethod::Other
    };

    let range_header_value = request.headers().get("range");
    let range = range_header_value.map(|hv| AssetHeaderValue {
        raw: hv.as_bytes(),
        value: hv.to_str().ok(),
    });

    let path = if method == AssetMethod::Get || method == AssetMethod::Head {
        match resolve_request_path(cache_root, &request) {
            Ok(p) => p.to_string_lossy().to_string(),
            Err(err) => {
                return Response::builder()
                    .status(StatusCode::FORBIDDEN)
                    .header(CONTENT_TYPE, "text/plain")
                    .header("Access-Control-Allow-Origin", window_origin)
                    .header("Cross-Origin-Resource-Policy", "cross-origin")
                    .body(err.into_bytes())
                    .unwrap();
            }
        }
    } else {
        String::new()
    };

    let cache_root = cache_root.to_path_buf();
    let core_resp = handle_asset_request(
        AssetRequest {
            method,
            path: &path,
            range,
        },
        |path| desktop::pyodide_assets::pyodide_cache_path_is_allowed(Path::new(path), &cache_root),
    );

    let mut builder = Response::builder()
        .status(StatusCode::from_u16(core_resp.status).unwrap_or(StatusCode::INTERNAL_SERVER_ERROR))
        .header("Access-Control-Allow-Origin", window_origin)
        .header("Cross-Origin-Resource-Policy", "cross-origin");

    for (k, v) in core_resp.headers {
        builder = builder.header(k, v);
    }

    builder.body(core_resp.body).unwrap()
}
