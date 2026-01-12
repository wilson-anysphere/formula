use std::io::{Read, Seek, SeekFrom, Write};

use rand_core::RngCore;
use tauri::http::{
    header::{
        ACCEPT_RANGES, ACCESS_CONTROL_EXPOSE_HEADERS, CONTENT_LENGTH, CONTENT_RANGE, CONTENT_TYPE,
    },
    Request, Response, StatusCode,
};
use tauri::path::SafePathBuf;
use tauri::scope::fs::Scope;
use tauri::utils::mime_type::MimeType;
use tauri::{Manager, Runtime, UriSchemeContext};
use url::Url;

/// Max bytes allowed for non-range `asset://...` responses.
///
/// Range requests are already clamped to ~1 MiB, but non-range requests historically read the
/// entire file into memory, which can OOM the backend if a compromised webview requests a very
/// large file.
const MAX_NON_RANGE_ASSET_BYTES: u64 = 10 * 1024 * 1024; // 10 MiB

/// Max length for the `Range:` request header.
///
/// This is a defense-in-depth limit to avoid pathological multi-range headers causing excessive
/// CPU/memory usage during parsing.
const MAX_RANGE_HEADER_BYTES: usize = 4 * 1024;

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
        return Response::builder()
            .status(StatusCode::FORBIDDEN)
            .header("Access-Control-Allow-Origin", &window_origin)
            .header("Cross-Origin-Resource-Policy", "cross-origin")
            .body(Vec::new())
            .unwrap();
    }

    // Security boundary: `asset://` is effectively "read a local file inside the configured scope".
    // If the webview ever navigates to remote/untrusted content, we must not allow that origin to
    // access `asset://` at all (even as a no-cors subresource).
    let window_url = current_window_url(&ctx);
    if !window_url
        .as_ref()
        .is_some_and(desktop::ipc_origin::is_trusted_app_origin)
    {
        let url_for_log = window_url
            .as_ref()
            .map(|u| u.to_string())
            .unwrap_or_else(|| "<unknown>".to_string());
        eprintln!("[asset protocol] blocked request from untrusted origin: {url_for_log}");
        return Response::builder()
            .status(StatusCode::FORBIDDEN)
            .header(CONTENT_TYPE, "text/plain")
            .header("Access-Control-Allow-Origin", &window_origin)
            .header("Cross-Origin-Resource-Policy", "cross-origin")
            .body(b"asset protocol is only available from trusted app-local origins".to_vec())
            .unwrap();
    }

    let scope = match Scope::new(
        ctx.app_handle(),
        &ctx.app_handle().config().app.security.asset_protocol.scope,
    ) {
        Ok(scope) => scope,
        Err(err) => {
            return Response::builder()
                .status(StatusCode::INTERNAL_SERVER_ERROR)
                .header(CONTENT_TYPE, "text/plain")
                .header("Access-Control-Allow-Origin", &window_origin)
                .header("Cross-Origin-Resource-Policy", "cross-origin")
                .body(format!("failed to initialize asset protocol scope: {err}").into_bytes())
                .unwrap();
        }
    };

    match get_response(request, &scope, &window_origin) {
        Ok(response) => response,
        Err(err) => Response::builder()
            .status(StatusCode::INTERNAL_SERVER_ERROR)
            .header(CONTENT_TYPE, "text/plain")
            .header("Access-Control-Allow-Origin", &window_origin)
            .header("Cross-Origin-Resource-Policy", "cross-origin")
            .body(err.to_string().into_bytes())
            .unwrap(),
    }
}

fn current_window_url<R: Runtime>(ctx: &UriSchemeContext<'_, R>) -> Option<Url> {
    let window = ctx.app_handle().get_webview_window(ctx.webview_label())?;
    window.as_ref().url().ok()
}

fn stable_window_origin<R: Runtime>(ctx: &UriSchemeContext<'_, R>) -> String {
    // Mirror Tauri upstream behavior: the window origin is computed once from the
    // *initial* webview URL and then used by the protocol handler.
    //
    // We compute the equivalent stable origin from the app config + platform so an
    // arbitrary navigation cannot change the effective CORS policy.
    let config = ctx.app_handle().config();

    let use_https_scheme = config
        .app
        .windows
        .iter()
        .find(|w| w.label == ctx.webview_label())
        .map(|w| w.use_https_scheme)
        .unwrap_or(false);

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
) -> Result<Response<Vec<u8>>, Box<dyn std::error::Error>> {
    // skip leading `/`
    let path = percent_encoding::percent_decode(&request.uri().path().as_bytes()[1..])
        .decode_utf8_lossy()
        .to_string();

    let mut resp = Response::builder()
        .header("Access-Control-Allow-Origin", window_origin)
        .header("Cross-Origin-Resource-Policy", "cross-origin");

    if let Err(e) = SafePathBuf::new(path.clone().into()) {
        eprintln!("[asset protocol] invalid path \"{path}\": {e}");
        return resp
            .status(StatusCode::FORBIDDEN)
            .body(Vec::new())
            .map_err(Into::into);
    }

    if !scope.is_allowed(&path) {
        eprintln!("[asset protocol] path not allowed by scope: {path}");
        return resp
            .status(StatusCode::FORBIDDEN)
            .body(Vec::new())
            .map_err(Into::into);
    }

    let mut file = match std::fs::File::open(&path) {
        Ok(file) => file,
        Err(err) => {
            return match err.kind() {
                std::io::ErrorKind::NotFound => resp
                    .status(StatusCode::NOT_FOUND)
                    .body(Vec::new())
                    .map_err(Into::into),
                std::io::ErrorKind::PermissionDenied => resp
                    .status(StatusCode::FORBIDDEN)
                    .body(Vec::new())
                    .map_err(Into::into),
                _ => Err(err.into()),
            };
        }
    };

    let metadata = file.metadata()?;
    if !metadata.is_file() {
        return resp
            .status(StatusCode::NOT_FOUND)
            .body(Vec::new())
            .map_err(Into::into);
    }

    // File length.
    let len = metadata.len();

    // We only enforce this on non-range (full-body) requests. Range requests are clamped to
    // `MAX_LEN` below so they remain bounded per request.
    let range_header_value = request.headers().get("range");
    if let Some(range_value) = range_header_value {
        let nbytes = range_value.as_bytes().len();
        if nbytes > MAX_RANGE_HEADER_BYTES {
            eprintln!(
                "[asset protocol] refusing overly large Range header: {path} (bytes={nbytes}, limit={MAX_RANGE_HEADER_BYTES})"
            );
            return resp
                .status(StatusCode::PAYLOAD_TOO_LARGE)
                .header(CONTENT_TYPE, "text/plain")
                .body(b"range header is too large".to_vec())
                .map_err(Into::into);
        }
    }

    let range_header = range_header_value.and_then(|r| r.to_str().ok());

    if range_header.is_none()
        && request.method() != tauri::http::Method::HEAD
        && len > MAX_NON_RANGE_ASSET_BYTES
    {
        eprintln!(
            "[asset protocol] refusing to serve large file without Range: {path} (size={len}, limit={MAX_NON_RANGE_ASSET_BYTES})"
        );
        return resp
            .status(StatusCode::PAYLOAD_TOO_LARGE)
            .header(CONTENT_TYPE, "text/plain")
            .body(b"asset file is too large; use Range requests".to_vec())
            .map_err(Into::into);
    }

    // MIME type detection.
    let (mime_type, read_bytes) = {
        let nbytes = len.min(8192);
        let mut magic_buf = Vec::with_capacity(nbytes as usize);
        let old_pos = file.stream_position()?;
        (&mut file).take(nbytes).read_to_end(&mut magic_buf)?;
        file.seek(SeekFrom::Start(old_pos))?;
        (
            MimeType::parse(&magic_buf, &path),
            // Return the magic bytes if we already read the whole file so we can reuse them below.
            if len < 8192 { Some(magic_buf) } else { None },
        )
    };

    resp = resp.header(CONTENT_TYPE, mime_type.to_string());

    // Handle 206 (partial range) requests.
    if let Some(range_header) = range_header {
        resp = resp.header(ACCEPT_RANGES, "bytes");
        resp = resp.header(ACCESS_CONTROL_EXPOSE_HEADERS, "content-range");

        let not_satisfiable = || {
            Response::builder()
                .status(StatusCode::RANGE_NOT_SATISFIABLE)
                .header(CONTENT_RANGE, format!("bytes */{len}"))
                .header("Access-Control-Allow-Origin", window_origin)
                .header("Cross-Origin-Resource-Policy", "cross-origin")
                .body(Vec::new())
                .map_err(Into::into)
        };

        let ranges = if let Ok(ranges) = http_range::HttpRange::parse(range_header, len) {
            ranges
                .iter()
                // Map to <start-end>, example: 0-499
                .map(|r| (r.start, r.start + r.length - 1))
                .collect::<Vec<_>>()
        } else {
            return not_satisfiable();
        };

        // Avoid unbounded allocations in multi-range requests. Range headers can contain many
        // comma-separated ranges; even with per-range clamping, building a multipart response could
        // otherwise allocate an arbitrarily large `Vec<u8>`.
        const MAX_RANGES: usize = 32;
        if ranges.len() > MAX_RANGES {
            eprintln!(
                "[asset protocol] too many ranges requested: {path} (count={})",
                ranges.len()
            );
            return Response::builder()
                .status(StatusCode::PAYLOAD_TOO_LARGE)
                .header("Access-Control-Allow-Origin", window_origin)
                .header("Cross-Origin-Resource-Policy", "cross-origin")
                .body(Vec::new())
                .map_err(Into::into);
        }

        /// Max bytes we send in one range.
        const MAX_LEN: u64 = 1000 * 1024;

        // Single range.
        if ranges.len() == 1 {
            let (start, mut end) = ranges[0];
            if start >= len || end >= len || end < start {
                return not_satisfiable();
            }

            // Clamp to MAX_LEN.
            end = start + (end - start).min(len - start).min(MAX_LEN - 1);
            let nbytes = end + 1 - start;

            let mut buf = Vec::with_capacity(nbytes as usize);
            file.seek(SeekFrom::Start(start))?;
            (&mut file).take(nbytes).read_to_end(&mut buf)?;

            resp = resp.header(CONTENT_RANGE, format!("bytes {start}-{end}/{len}"));
            resp = resp.header(CONTENT_LENGTH, nbytes.to_string());
            resp = resp.status(StatusCode::PARTIAL_CONTENT);
            return resp.body(buf).map_err(Into::into);
        }

        // Multi-range support (rare; kept for compatibility).
        // We implement it correctly (Tauri's internal implementation historically had a few quirks).
        let ranges = ranges
            .into_iter()
            .filter_map(|(start, mut end)| {
                if start >= len || end >= len || end < start {
                    None
                } else {
                    end = start + (end - start).min(len - start).min(MAX_LEN - 1);
                    Some((start, end))
                }
            })
            .collect::<Vec<_>>();

        if ranges.is_empty() {
            return not_satisfiable();
        }

        let mut total_range_bytes = 0u64;
        for (start, end) in &ranges {
            // (end + 1 - start) should be safe because we validated start/end above.
            total_range_bytes = total_range_bytes.saturating_add(end + 1 - start);
        }
        if total_range_bytes > MAX_NON_RANGE_ASSET_BYTES {
            eprintln!(
                "[asset protocol] refusing to serve multi-range response larger than limit: {path} (bytes={total_range_bytes}, limit={MAX_NON_RANGE_ASSET_BYTES})"
            );
            return Response::builder()
                .status(StatusCode::PAYLOAD_TOO_LARGE)
                .header("Access-Control-Allow-Origin", window_origin)
                .header("Cross-Origin-Resource-Policy", "cross-origin")
                .body(Vec::new())
                .map_err(Into::into);
        }

        let boundary = random_boundary();
        let boundary_sep = format!("\r\n--{boundary}\r\n");
        let boundary_closer = format!("\r\n--{boundary}--\r\n");

        resp = resp.header(
            CONTENT_TYPE,
            format!("multipart/byteranges; boundary={boundary}"),
        );
        resp = resp.status(StatusCode::PARTIAL_CONTENT);

        let mut buf = Vec::new();
        for (start, end) in ranges {
            buf.write_all(boundary_sep.as_bytes())?;
            buf.write_all(format!("{CONTENT_TYPE}: {mime_type}\r\n").as_bytes())?;
            buf.write_all(format!("{CONTENT_RANGE}: bytes {start}-{end}/{len}\r\n").as_bytes())?;
            buf.write_all(b"\r\n")?;

            let nbytes = end + 1 - start;
            let mut local_buf = Vec::with_capacity(nbytes as usize);
            file.seek(SeekFrom::Start(start))?;
            (&mut file).take(nbytes).read_to_end(&mut local_buf)?;
            buf.extend_from_slice(&local_buf);
        }
        buf.write_all(boundary_closer.as_bytes())?;
        return resp.body(buf).map_err(Into::into);
    }

    if request.method() == tauri::http::Method::HEAD {
        // If HEAD, don't return a body.
        resp = resp.header(CONTENT_LENGTH, len.to_string());
        return resp.body(Vec::new()).map_err(Into::into);
    }

    // Avoid reading the file if we already read it as part of MIME detection.
    let buf = if let Some(bytes) = read_bytes {
        bytes
    } else {
        let mut local_buf = Vec::with_capacity(len as usize);
        file.read_to_end(&mut local_buf)?;
        local_buf
    };
    resp = resp.header(CONTENT_LENGTH, len.to_string());
    resp.body(buf).map_err(Into::into)
}

fn random_boundary() -> String {
    let mut x = [0_u8; 30];
    // `rand_core` is already a dependency of the desktop backend.
    rand_core::OsRng.fill_bytes(&mut x);
    x.iter()
        .map(|b| format!("{b:x}"))
        .collect::<Vec<_>>()
        .join("")
}
