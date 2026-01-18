use tauri::http::{
    header::CONTENT_TYPE,
    Request, Response, StatusCode,
};
use tauri::scope::fs::Scope;
use tauri::{Manager, Runtime, UriSchemeContext};
use url::Url;

/// Custom `asset:` protocol handler with COEP-friendly headers.
///
/// Why:
/// With `Cross-Origin-Embedder-Policy: require-corp` enabled on the main document,
/// cross-origin subresources (like `asset://...` images produced by `convertFileSrc`)
/// must opt into being embedded. We do that by adding:
/// `Cross-Origin-Resource-Policy: cross-origin` to *all* `asset:` responses.
///
/// Note:
/// We intentionally do **not** set `Access-Control-Allow-Origin: *` to avoid making
/// arbitrary local files readable via `fetch()` from unexpected origins; instead we
/// mirror Tauri's upstream behavior of reflecting the *initial* webview origin.
///
/// Important security property:
/// The origin is computed from configuration/platform rules instead of the current
/// webview URL so that an external navigation cannot gain CORS access to `asset://`
/// resources.
pub fn handler<R: Runtime>(
    ctx: UriSchemeContext<'_, R>,
    request: Request<Vec<u8>>,
) -> Response<Vec<u8>> {
    let window_origin = stable_window_origin(&ctx);

    if !ctx.app_handle().config().app.security.asset_protocol.enable {
        // Match the intent of Tauri's built-in asset protocol: if it's not enabled,
        // deny all requests.
        let mut response = crate::http_response::response(StatusCode::FORBIDDEN, Vec::new());
        crate::http_response::insert_header(
            &mut response,
            "Access-Control-Allow-Origin",
            &window_origin,
        );
        crate::http_response::insert_header(
            &mut response,
            "Cross-Origin-Resource-Policy",
            "cross-origin",
        );
        return response;
    }

    // Security boundary: `asset://` is effectively "read a local file inside the configured scope".
    // If the webview ever navigates to remote/untrusted content, we must not allow that origin to
    // access `asset://` at all (even as a no-cors subresource).
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
        crate::stdio::stderrln(format_args!(
            "[asset protocol] blocked request from untrusted origin: {url_for_log}"
        ));
        let mut response = crate::http_response::response_with_content_type(
            StatusCode::FORBIDDEN,
            "text/plain",
            b"asset protocol is only available from trusted app-local origins".to_vec(),
        );
        crate::http_response::insert_header(
            &mut response,
            "Access-Control-Allow-Origin",
            &window_origin,
        );
        crate::http_response::insert_header(
            &mut response,
            "Cross-Origin-Resource-Policy",
            "cross-origin",
        );
        return response;
    }

    let scope = match Scope::new(
        ctx.app_handle(),
        &ctx.app_handle().config().app.security.asset_protocol.scope,
    ) {
        Ok(scope) => scope,
        Err(err) => {
            let mut response = crate::http_response::response_with_content_type(
                StatusCode::INTERNAL_SERVER_ERROR,
                "text/plain",
                format!("failed to initialize asset protocol scope: {err}").into_bytes(),
            );
            crate::http_response::insert_header(
                &mut response,
                "Access-Control-Allow-Origin",
                &window_origin,
            );
            crate::http_response::insert_header(
                &mut response,
                "Cross-Origin-Resource-Policy",
                "cross-origin",
            );
            return response;
        }
    };

    get_response(request, &scope, &window_origin)
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
    // Mirror Tauri upstream behavior: the window origin is computed once from the
    // *initial* webview URL and then used by the protocol handler.
    //
    // We compute the equivalent stable origin from the app config + platform so an
    // arbitrary navigation cannot change the effective CORS policy.
    let config = ctx.app_handle().config();

    let use_https_scheme = use_https_scheme(ctx);

    // `dev_url` is represented as either `String` or `Url` depending on the Tauri version.
    // Both expose `as_str()`, so prefer that over `as_deref()` to avoid tying this code to a
    // specific config representation.
    let dev_url = config.build.dev_url.as_ref().map(|url| url.as_str());

    desktop::tauri_origin::stable_webview_origin(
        tauri::is_dev(),
        dev_url,
        use_https_scheme,
        desktop::tauri_origin::DesktopPlatform::current(),
    )
}

fn get_response(
    request: Request<Vec<u8>>,
    scope: &Scope,
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
        let raw = request.uri().path().strip_prefix('/').unwrap_or("");
        percent_encoding::percent_decode(raw.as_bytes())
            .decode_utf8_lossy()
            .to_string()
    } else {
        String::new()
    };

    let core_resp = handle_asset_request(
        AssetRequest {
            method,
            path: &path,
            range,
        },
        |path| scope.is_allowed(path),
    );

    let mut response = crate::http_response::response(
        StatusCode::from_u16(core_resp.status).unwrap_or(StatusCode::INTERNAL_SERVER_ERROR),
        core_resp.body,
    );
    crate::http_response::insert_header(&mut response, "Access-Control-Allow-Origin", window_origin);
    crate::http_response::insert_header(
        &mut response,
        "Cross-Origin-Resource-Policy",
        "cross-origin",
    );
    for (k, v) in core_resp.headers {
        crate::http_response::insert_header(&mut response, &k, &v);
    }
    response
}
