use std::collections::BTreeMap;
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::{Component, Path};
use std::{fs, io};

use rand_core::RngCore;

/// Max bytes allowed for non-range `asset://...` responses.
///
/// Range requests are already clamped to ~1 MiB, but non-range requests historically read the
/// entire file into memory, which can OOM the backend if a compromised webview requests a very
/// large file.
pub const MAX_NON_RANGE_ASSET_BYTES: u64 = 10 * 1024 * 1024; // 10 MiB

/// Max length for the `Range:` request header.
///
/// This is a defense-in-depth limit to avoid pathological multi-range headers causing excessive
/// CPU/memory usage during parsing.
pub const MAX_RANGE_HEADER_BYTES: usize = 4 * 1024;

/// Avoid unbounded allocations in multi-range requests.
///
/// Range headers can contain many comma-separated ranges; even with per-range clamping, building a
/// multipart response could otherwise allocate an arbitrarily large `Vec<u8>`.
pub const MAX_RANGES: usize = 32;

/// Max bytes we send in one range.
pub const MAX_RANGE_LEN_BYTES: u64 = 1000 * 1024;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum AssetMethod {
    Get,
    Head,
    Options,
    Other,
}

#[derive(Clone, Copy, Debug)]
pub struct AssetHeaderValue<'a> {
    /// Raw header bytes, used for size limits even when the value isn't valid UTF-8.
    pub raw: &'a [u8],
    /// Parsed UTF-8 representation, if available.
    pub value: Option<&'a str>,
}

#[derive(Clone, Copy, Debug)]
pub struct AssetRequest<'a> {
    /// Only `GET`, `HEAD`, and `OPTIONS` are supported.
    pub method: AssetMethod,
    /// Percent-decoded filesystem path (as rendered in the `asset://` URL).
    pub path: &'a str,
    /// Optional `Range:` request header.
    pub range: Option<AssetHeaderValue<'a>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AssetResponse {
    pub status: u16,
    /// Response headers.
    ///
    /// Stored as a deterministic map for stable tests. Header names are case-insensitive per HTTP,
    /// but we preserve a lower-case canonical form.
    pub headers: BTreeMap<String, String>,
    pub body: Vec<u8>,
}

impl AssetResponse {
    pub fn header(&self, name: &str) -> Option<&str> {
        self.headers
            .get(&name.to_ascii_lowercase())
            .map(|s| s.as_str())
    }
}

fn header_name(name: &str) -> String {
    name.to_ascii_lowercase()
}

fn resp(status: u16) -> AssetResponse {
    AssetResponse {
        status,
        headers: BTreeMap::new(),
        body: Vec::new(),
    }
}

fn resp_with_body(status: u16, content_type: &str, body: impl Into<Vec<u8>>) -> AssetResponse {
    let mut r = resp(status);
    r.headers
        .insert(header_name("content-type"), content_type.to_string());
    r.body = body.into();
    r
}

fn validate_asset_path(path: &str) -> bool {
    // Defense in depth: reject NULs early so we don't accidentally pass a truncated string to
    // platform APIs.
    if path.is_empty() || path.contains('\0') {
        return false;
    }

    // Reject any explicit `.` or `..` segments.
    //
    // `Path::components()` performs platform-specific parsing (e.g. `C:` prefixes on Windows),
    // which keeps this check consistent across OSes.
    for c in Path::new(path).components() {
        match c {
            Component::CurDir | Component::ParentDir => return false,
            _ => {}
        }
    }

    true
}

fn detect_mime_type(file: &mut fs::File, path: &str, len: u64) -> io::Result<(String, Option<Vec<u8>>)> {
    // Read a small prefix for magic bytes; this is bounded and safe even for very large files.
    let nbytes = len.min(8192);
    let mut magic_buf: Vec<u8> = Vec::new();
    let _ = magic_buf.try_reserve(usize::try_from(nbytes).unwrap_or(usize::MAX));
    let old_pos = file.stream_position()?;
    (&mut *file).take(nbytes).read_to_end(&mut magic_buf)?;
    file.seek(SeekFrom::Start(old_pos))?;

    // Prefer magic-byte detection, then fall back to extension-based guessing.
    let mime = infer::get(&magic_buf)
        .map(|t| t.mime_type().to_string())
        .or_else(|| {
            let guess = mime_guess::from_path(path).first_or_octet_stream();
            Some(guess.essence_str().to_string())
        })
        .unwrap_or_else(|| "application/octet-stream".to_string());

    let read_bytes = if len < 8192 { Some(magic_buf) } else { None };
    Ok((mime, read_bytes))
}

fn random_boundary() -> String {
    let mut bytes = [0_u8; 30];
    // `rand_core` is already a dependency of the desktop backend.
    rand_core::OsRng.fill_bytes(&mut bytes);
    hex::encode(bytes)
}

/// Pure-Rust `asset://` file responder.
///
/// This function contains the security/performance-critical hardening logic: request method
/// handling, path validation, range parsing/clamping, and size limits.
///
/// The Tauri-facing wrapper is expected to:
/// - percent-decode the request URI path,
/// - enforce trusted origin gating, and
/// - add cross-origin headers like `Access-Control-Allow-Origin` and
///   `Cross-Origin-Resource-Policy`.
pub fn handle_asset_request(
    request: AssetRequest<'_>,
    is_allowed_by_scope: impl Fn(&str) -> bool,
) -> AssetResponse {
    // Support CORS preflight requests so `fetch()` can opt into non-simple headers (e.g. `Range`)
    // without forcing the backend to read any file bytes.
    if request.method == AssetMethod::Options {
        let mut r = resp(204);
        r.headers.insert(
            header_name("access-control-allow-methods"),
            "GET, HEAD, OPTIONS".to_string(),
        );
        // `Range` is not a simple request header, so it requires preflight when used from
        // `fetch()`.
        r.headers.insert(
            header_name("access-control-allow-headers"),
            "range".to_string(),
        );
        return r;
    }

    // The asset protocol is a read-only file responder. Reject unexpected methods early so callers
    // cannot trigger surprising behavior via non-GET verbs.
    if request.method != AssetMethod::Get && request.method != AssetMethod::Head {
        let mut r = resp(405);
        r.headers
            .insert(header_name("allow"), "GET, HEAD, OPTIONS".to_string());
        return r;
    }

    let path = request.path;
    if !validate_asset_path(path) {
        crate::stdio::stderrln(format_args!("[asset protocol] invalid path \"{path}\""));
        return resp(403);
    }

    if !is_allowed_by_scope(path) {
        crate::stdio::stderrln(format_args!(
            "[asset protocol] path not allowed by scope: {path}"
        ));
        return resp(403);
    }

    let mut file = match fs::File::open(path) {
        Ok(file) => file,
        Err(err) => {
            return match err.kind() {
                io::ErrorKind::NotFound => resp(404),
                io::ErrorKind::PermissionDenied => resp(403),
                _ => resp_with_body(
                    500,
                    "text/plain",
                    format!("failed to open file: {err}").into_bytes(),
                ),
            };
        }
    };

    let metadata = match file.metadata() {
        Ok(metadata) => metadata,
        Err(err) => {
            return resp_with_body(
                500,
                "text/plain",
                format!("failed to stat file: {err}").into_bytes(),
            );
        }
    };
    if !metadata.is_file() {
        return resp(404);
    }

    // File length.
    let len = metadata.len();

    // We only enforce this on non-range (full-body) requests. Range requests are clamped below so
    // they remain bounded per request.
    let range_header = request.range;
    if let Some(range_value) = range_header {
        let nbytes = range_value.raw.len();
        if nbytes > MAX_RANGE_HEADER_BYTES {
            crate::stdio::stderrln(format_args!(
                "[asset protocol] refusing overly large Range header: {path} (bytes={nbytes}, limit={MAX_RANGE_HEADER_BYTES})"
            ));
            return resp_with_body(413, "text/plain", b"range header is too large".to_vec());
        }
    }

    let range_header = range_header.and_then(|r| r.value);

    if range_header.is_none() && request.method != AssetMethod::Head && len > MAX_NON_RANGE_ASSET_BYTES {
        crate::stdio::stderrln(format_args!(
            "[asset protocol] refusing to serve large file without Range: {path} (size={len}, limit={MAX_NON_RANGE_ASSET_BYTES})"
        ));
        return resp_with_body(
            413,
            "text/plain",
            b"asset file is too large; use Range requests".to_vec(),
        );
    }

    // MIME type detection.
    let (mime_type, read_bytes) = match detect_mime_type(&mut file, path, len) {
        Ok(x) => x,
        Err(err) => {
            return resp_with_body(
                500,
                "text/plain",
                format!("failed to detect mime type: {err}").into_bytes(),
            );
        }
    };

    // Handle 206 (partial range) requests.
    if let Some(range_header) = range_header {
        let mut base_headers = BTreeMap::new();
        base_headers.insert(header_name("accept-ranges"), "bytes".to_string());
        base_headers.insert(
            header_name("access-control-expose-headers"),
            "content-range".to_string(),
        );
        base_headers.insert(header_name("content-type"), mime_type.clone());

        let not_satisfiable = || {
            let mut r = resp(416);
            r.headers.insert(
                header_name("content-range"),
                format!("bytes */{len}"),
            );
            r
        };

        let ranges = if let Ok(ranges) = http_range::HttpRange::parse(range_header, len) {
            ranges
                .iter()
                // Map to <start-end>, example: 0-499
                .map(|r| (r.start, r.start.saturating_add(r.length.saturating_sub(1))))
                .collect::<Vec<_>>()
        } else {
            return not_satisfiable();
        };

        if ranges.len() > MAX_RANGES {
            crate::stdio::stderrln(format_args!(
                "[asset protocol] too many ranges requested: {path} (count={})",
                ranges.len()
            ));
            return resp(413);
        }

        // Single range.
        if ranges.len() == 1 {
            let (start, mut end) = ranges[0];
            if start >= len || end >= len || end < start {
                return not_satisfiable();
            }

            // Clamp to MAX_RANGE_LEN_BYTES.
            end = start + (end - start).min(len - start).min(MAX_RANGE_LEN_BYTES - 1);
            let nbytes = end + 1 - start;

            let mut buf: Vec<u8> = Vec::new();
            let _ = buf.try_reserve(usize::try_from(nbytes).unwrap_or(usize::MAX));
            if let Err(err) = file.seek(SeekFrom::Start(start)) {
                return resp_with_body(
                    500,
                    "text/plain",
                    format!("failed to seek: {err}").into_bytes(),
                );
            }
            if let Err(err) = (&mut file).take(nbytes).read_to_end(&mut buf) {
                return resp_with_body(
                    500,
                    "text/plain",
                    format!("failed to read range: {err}").into_bytes(),
                );
            }

            let mut r = resp(206);
            r.headers = base_headers;
            r.headers.insert(
                header_name("content-range"),
                format!("bytes {start}-{end}/{len}"),
            );
            r.headers
                .insert(header_name("content-length"), nbytes.to_string());
            r.body = buf;
            return r;
        }

        // Multi-range support (rare; kept for compatibility).
        let ranges = ranges
            .into_iter()
            .filter_map(|(start, mut end)| {
                if start >= len || end >= len || end < start {
                    None
                } else {
                    end = start + (end - start).min(len - start).min(MAX_RANGE_LEN_BYTES - 1);
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
            crate::stdio::stderrln(format_args!(
                "[asset protocol] refusing to serve multi-range response larger than limit: {path} (bytes={total_range_bytes}, limit={MAX_NON_RANGE_ASSET_BYTES})"
            ));
            return resp(413);
        }

        let boundary = random_boundary();
        let boundary_sep = format!("\r\n--{boundary}\r\n");
        let boundary_closer = format!("\r\n--{boundary}--\r\n");

        let mut r = resp(206);
        r.headers.insert(
            header_name("content-type"),
            format!("multipart/byteranges; boundary={boundary}"),
        );
        // Keep the range-related headers from the base (accept-ranges + expose-headers).
        r.headers
            .insert(header_name("accept-ranges"), "bytes".to_string());
        r.headers.insert(
            header_name("access-control-expose-headers"),
            "content-range".to_string(),
        );

        // Pre-size to avoid repeated reallocations. The multipart framing overhead is small and is
        // intentionally excluded from the size cap to preserve existing behavior.
        let mut buf: Vec<u8> = Vec::new();
        let _ = buf.try_reserve(usize::try_from(total_range_bytes).unwrap_or(usize::MAX));
        for (start, end) in ranges {
            if let Err(err) = buf.write_all(boundary_sep.as_bytes()) {
                return resp_with_body(
                    500,
                    "text/plain",
                    format!("failed to write multipart boundary: {err}").into_bytes(),
                );
            }
            if let Err(err) = buf.write_all(format!("content-type: {mime_type}\r\n").as_bytes()) {
                return resp_with_body(
                    500,
                    "text/plain",
                    format!("failed to write multipart headers: {err}").into_bytes(),
                );
            }
            if let Err(err) = buf.write_all(
                format!("content-range: bytes {start}-{end}/{len}\r\n").as_bytes(),
            ) {
                return resp_with_body(
                    500,
                    "text/plain",
                    format!("failed to write multipart headers: {err}").into_bytes(),
                );
            }
            if let Err(err) = buf.write_all(b"\r\n") {
                return resp_with_body(
                    500,
                    "text/plain",
                    format!("failed to write multipart header terminator: {err}").into_bytes(),
                );
            }

            let nbytes = end + 1 - start;
            let mut local_buf: Vec<u8> = Vec::new();
            let _ = local_buf.try_reserve(usize::try_from(nbytes).unwrap_or(usize::MAX));
            if let Err(err) = file.seek(SeekFrom::Start(start)) {
                return resp_with_body(
                    500,
                    "text/plain",
                    format!("failed to seek: {err}").into_bytes(),
                );
            }
            if let Err(err) = (&mut file).take(nbytes).read_to_end(&mut local_buf) {
                return resp_with_body(
                    500,
                    "text/plain",
                    format!("failed to read range: {err}").into_bytes(),
                );
            }
            buf.extend_from_slice(&local_buf);
        }
        if let Err(err) = buf.write_all(boundary_closer.as_bytes()) {
            return resp_with_body(
                500,
                "text/plain",
                format!("failed to write multipart boundary closer: {err}").into_bytes(),
            );
        }
        r.body = buf;
        return r;
    }

    let mut r = resp(200);
    r.headers
        .insert(header_name("content-type"), mime_type.to_string());

    if request.method == AssetMethod::Head {
        // If HEAD, don't return a body.
        r.headers
            .insert(header_name("content-length"), len.to_string());
        return r;
    }

    // Avoid reading the file if we already read it as part of MIME detection.
    let buf = if let Some(bytes) = read_bytes {
        bytes
    } else {
        let mut local_buf: Vec<u8> = Vec::new();
        let _ = local_buf.try_reserve(usize::try_from(len).unwrap_or(usize::MAX));
        if let Err(err) = file.read_to_end(&mut local_buf) {
            return resp_with_body(
                500,
                "text/plain",
                format!("failed to read file: {err}").into_bytes(),
            );
        }
        local_buf
    };
    r.headers
        .insert(header_name("content-length"), len.to_string());
    r.body = buf;
    r
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;

    fn allow_all(_: &str) -> bool {
        true
    }

    fn range_header(value: &str) -> AssetHeaderValue<'_> {
        AssetHeaderValue {
            raw: value.as_bytes(),
            value: Some(value),
        }
    }

    fn get_tempfile_with_len(len: u64) -> tempfile::NamedTempFile {
        let file = tempfile::NamedTempFile::new().expect("tempfile");
        file.as_file().set_len(len).expect("set_len");
        file
    }

    #[test]
    fn options_preflight_returns_204_and_cors_headers() {
        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Options,
                path: "/does/not/matter",
                range: None,
            },
            allow_all,
        );
        assert_eq!(resp.status, 204);
        assert!(resp.body.is_empty());
        assert_eq!(
            resp.header("access-control-allow-methods"),
            Some("GET, HEAD, OPTIONS")
        );
        assert_eq!(resp.header("access-control-allow-headers"), Some("range"));
    }

    #[test]
    fn method_not_allowed_returns_405_and_allow_header() {
        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Other,
                path: "/does/not/matter",
                range: None,
            },
            allow_all,
        );
        assert_eq!(resp.status, 405);
        assert_eq!(resp.header("allow"), Some("GET, HEAD, OPTIONS"));
        assert!(resp.body.is_empty());
    }

    #[test]
    fn non_range_get_over_10_mib_returns_413() {
        let file = get_tempfile_with_len(MAX_NON_RANGE_ASSET_BYTES + 1);
        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Get,
                path: file.path().to_str().unwrap(),
                range: None,
            },
            allow_all,
        );
        assert_eq!(resp.status, 413);
        assert_eq!(resp.header("content-type"), Some("text/plain"));
        assert_eq!(
            std::str::from_utf8(&resp.body).unwrap(),
            "asset file is too large; use Range requests"
        );
    }

    #[test]
    fn range_header_longer_than_4kib_returns_413() {
        let file = get_tempfile_with_len(128);
        let long_value = format!("bytes=0-0{}", "a".repeat(MAX_RANGE_HEADER_BYTES + 1));
        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Get,
                path: file.path().to_str().unwrap(),
                range: Some(AssetHeaderValue {
                    raw: long_value.as_bytes(),
                    value: Some(&long_value),
                }),
            },
            allow_all,
        );
        assert_eq!(resp.status, 413);
        assert_eq!(resp.header("content-type"), Some("text/plain"));
        assert_eq!(
            std::str::from_utf8(&resp.body).unwrap(),
            "range header is too large"
        );
    }

    #[test]
    fn single_range_clamps_to_1000_kib_and_sets_content_range() {
        let file_len = 2 * 1024 * 1024;
        let file = get_tempfile_with_len(file_len);

        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Get,
                path: file.path().to_str().unwrap(),
                range: Some(range_header("bytes=0-2000000")),
            },
            allow_all,
        );
        assert_eq!(resp.status, 206);
        let expected_content_range = format!("bytes 0-1023999/{file_len}");
        assert_eq!(
            resp.header("content-range"),
            Some(expected_content_range.as_str())
        );
        assert_eq!(resp.header("content-length"), Some("1024000"));
        assert_eq!(resp.body.len(), 1024000);
    }

    #[test]
    fn multi_range_with_more_than_32_ranges_returns_413() {
        let file = get_tempfile_with_len(128);
        let ranges = (0..33)
            .map(|i| format!("{i}-{i}"))
            .collect::<Vec<_>>()
            .join(",");
        let header_value = format!("bytes={ranges}");
        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Get,
                path: file.path().to_str().unwrap(),
                range: Some(range_header(&header_value)),
            },
            allow_all,
        );
        assert_eq!(resp.status, 413);
        assert!(resp.body.is_empty());
    }

    #[test]
    fn multi_range_total_bytes_over_10_mib_returns_413() {
        // Large sparse file, but we never read from it because the size cap triggers first.
        let file = get_tempfile_with_len(25 * 1024 * 1024);

        // 11 ranges, each larger than MAX_RANGE_LEN_BYTES, so each clamps to MAX_RANGE_LEN_BYTES.
        // 11 * 1000 KiB = 11,000 KiB > 10 MiB => 413.
        let mut parts = Vec::new();
        for i in 0..11u64 {
            let start = i * 2 * 1024 * 1024;
            let end = start + 2 * 1024 * 1024 - 1;
            parts.push(format!("{start}-{end}"));
        }
        let header_value = format!("bytes={}", parts.join(","));

        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Get,
                path: file.path().to_str().unwrap(),
                range: Some(range_header(&header_value)),
            },
            allow_all,
        );
        assert_eq!(resp.status, 413);
        assert!(resp.body.is_empty());
    }

    #[test]
    fn head_returns_content_length_and_empty_body() {
        let file = get_tempfile_with_len(123);
        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Head,
                path: file.path().to_str().unwrap(),
                range: None,
            },
            allow_all,
        );
        assert_eq!(resp.status, 200);
        assert_eq!(resp.header("content-length"), Some("123"));
        assert!(resp.body.is_empty());
    }

    #[test]
    fn invalid_path_traversal_returns_403() {
        // Even if the scope would allow it, traversal is rejected before touching the filesystem.
        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Get,
                path: "../etc/passwd",
                range: None,
            },
            allow_all,
        );
        assert_eq!(resp.status, 403);
        assert!(resp.body.is_empty());
    }

    #[test]
    fn out_of_scope_returns_403() {
        let file = get_tempfile_with_len(16);
        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Get,
                path: file.path().to_str().unwrap(),
                range: None,
            },
            |_| false,
        );
        assert_eq!(resp.status, 403);
        assert!(resp.body.is_empty());
    }

    #[test]
    fn missing_file_returns_404() {
        let dir = tempfile::tempdir().unwrap();
        let missing = dir.path().join("missing.txt");
        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Get,
                path: missing.to_str().unwrap(),
                range: None,
            },
            allow_all,
        );
        assert_eq!(resp.status, 404);
    }

    #[test]
    fn non_file_returns_404() {
        let dir = tempfile::tempdir().unwrap();
        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Get,
                path: dir.path().to_str().unwrap(),
                range: None,
            },
            allow_all,
        );
        assert_eq!(resp.status, 404);
    }

    #[test]
    fn range_not_satisfiable_returns_416_and_content_range() {
        let file = get_tempfile_with_len(100);
        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Get,
                path: file.path().to_str().unwrap(),
                range: Some(range_header("bytes=200-300")),
            },
            allow_all,
        );
        assert_eq!(resp.status, 416);
        assert_eq!(resp.header("content-range"), Some("bytes */100"));
        assert!(resp.body.is_empty());
    }

    #[test]
    fn full_get_returns_body_with_content_length() {
        let file = tempfile::NamedTempFile::new().unwrap();
        fs::write(file.path(), b"hello world").unwrap();
        let resp = handle_asset_request(
            AssetRequest {
                method: AssetMethod::Get,
                path: file.path().to_str().unwrap(),
                range: None,
            },
            allow_all,
        );
        assert_eq!(resp.status, 200);
        assert_eq!(resp.header("content-length"), Some("11"));
        assert_eq!(&resp.body, b"hello world");
    }
}
