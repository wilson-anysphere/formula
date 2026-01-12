use std::io::{Read, Seek, SeekFrom, Write};

use rand_core::RngCore;
use tauri::http::{
    header::{ACCEPT_RANGES, ACCESS_CONTROL_EXPOSE_HEADERS, CONTENT_LENGTH, CONTENT_RANGE, CONTENT_TYPE},
    Request, Response, StatusCode,
};
use tauri::path::SafePathBuf;
use tauri::scope::fs::Scope;
use tauri::utils::mime_type::MimeType;
use tauri::{Manager, Runtime, UriSchemeContext};

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
/// mirror Tauri's default behavior of reflecting the current window origin.
pub fn handler<R: Runtime>(ctx: UriSchemeContext<'_, R>, request: Request<Vec<u8>>) -> Response<Vec<u8>> {
    let window_origin = window_origin(&ctx);

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

    let scope = match Scope::new(ctx.app_handle(), &ctx.app_handle().config().app.security.asset_protocol.scope) {
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

fn window_origin<R: Runtime>(ctx: &UriSchemeContext<'_, R>) -> String {
    // This mirrors Tauri's internal computation from `manager/webview.rs`.
    //
    // We need this so `fetch()` can request `asset://` resources with the correct CORS origin,
    // and so we avoid enabling `Access-Control-Allow-Origin: *`.
    let Some(window) = ctx.app_handle().get_webview_window(ctx.webview_label()) else {
        return "null".to_string();
    };

    let Ok(window_url) = window.as_ref().url() else {
        return "null".to_string();
    };

    if window_url.scheme() == "data" {
        return "null".to_string();
    }

    let use_https_scheme = ctx
        .app_handle()
        .config()
        .app
        .windows
        .iter()
        .find(|w| w.label == ctx.webview_label())
        .map(|w| w.use_https_scheme)
        .unwrap_or(false);

    if (cfg!(windows) || cfg!(target_os = "android"))
        && window_url.scheme() != "http"
        && window_url.scheme() != "https"
    {
        let scheme = if use_https_scheme { "https" } else { "http" };
        return format!("{scheme}://{}.localhost", window_url.scheme());
    }

    if let Some(host) = window_url.host() {
        return format!(
            "{}://{}{}",
            window_url.scheme(),
            host,
            window_url
                .port()
                .map(|p| format!(":{p}"))
                .unwrap_or_default()
        );
    }

    "null".to_string()
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
        return resp.status(StatusCode::FORBIDDEN).body(Vec::new()).map_err(Into::into);
    }

    if !scope.is_allowed(&path) {
        eprintln!("[asset protocol] path not allowed by scope: {path}");
        return resp.status(StatusCode::FORBIDDEN).body(Vec::new()).map_err(Into::into);
    }

    let mut file = match std::fs::File::open(&path) {
        Ok(file) => file,
        Err(err) => {
            return match err.kind() {
                std::io::ErrorKind::NotFound => resp.status(StatusCode::NOT_FOUND).body(Vec::new()).map_err(Into::into),
                std::io::ErrorKind::PermissionDenied => {
                    resp.status(StatusCode::FORBIDDEN).body(Vec::new()).map_err(Into::into)
                }
                _ => Err(err.into()),
            };
        }
    };

    let metadata = file.metadata()?;
    if !metadata.is_file() {
        return resp.status(StatusCode::NOT_FOUND).body(Vec::new()).map_err(Into::into);
    }

    // File length.
    let len = metadata.len();

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
    if let Some(range_header) = request
        .headers()
        .get("range")
        .and_then(|r| r.to_str().ok())
        .map(|r| r.to_string())
    {
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

        let ranges = if let Ok(ranges) = http_range::HttpRange::parse(&range_header, len) {
            ranges
                .iter()
                // Map to <start-end>, example: 0-499
                .map(|r| (r.start, r.start + r.length - 1))
                .collect::<Vec<_>>()
        } else {
            return not_satisfiable();
        };

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

        let boundary = random_boundary();
        let boundary_sep = format!("\r\n--{boundary}\r\n");
        let boundary_closer = format!("\r\n--{boundary}--\r\n");

        resp = resp.header(CONTENT_TYPE, format!("multipart/byteranges; boundary={boundary}"));
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
