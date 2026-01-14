use directories::ProjectDirs;
use rand_core::RngCore;
use serde::Serialize;
use sha2::{Digest, Sha256};
use std::io::Read;
use std::path::{Component, Path, PathBuf};
use std::time::{Duration, Instant};
use tokio::io::AsyncWriteExt;
use uuid::Uuid;

/// Pyodide version used by the desktop app.
///
/// Keep in sync with:
/// - `apps/desktop/scripts/ensure-pyodide-assets.mjs`
/// - `packages/python-runtime/src/pyodide-main-thread.js` (DEFAULT_INDEX_URL)
const PYODIDE_VERSION: &str = "0.25.1";

const PYODIDE_CDN_BASE_URL: &str = "https://cdn.jsdelivr.net/pyodide/";

/// Upper bound for any single downloaded Pyodide asset.
///
/// Defense-in-depth: even though the download URLs are fixed, we treat network responses as
/// untrusted and cap allocations.
const MAX_SINGLE_PYODIDE_ASSET_BYTES: usize = 50 * 1024 * 1024; // 50 MiB

pub const PYODIDE_DOWNLOAD_PROGRESS_EVENT: &str = "pyodide-download-progress";

#[derive(Clone, Copy, Debug)]
pub struct PyodideAssetSpec<'a> {
    pub file_name: &'a str,
    pub sha256: &'a str,
}

/// Minimal Pyodide files required to initialize the runtime.
///
/// Keep in sync with the `requiredFiles` list in `ensure-pyodide-assets.mjs`.
const PYODIDE_REQUIRED_FILES: &[PyodideAssetSpec<'static>] = &[
    PyodideAssetSpec {
        file_name: "pyodide.js",
        sha256: "b9cb64a73cc4127eef7cdc75f0cd8307db9e90e93b66b1b6a789319511e1937c",
    },
    PyodideAssetSpec {
        file_name: "pyodide.asm.js",
        sha256: "512042dfdd406971c6fc920b6932e1a8eb5dd2ab3521aa89a020980e4a08bd4b",
    },
    PyodideAssetSpec {
        file_name: "pyodide.asm.wasm",
        sha256: "aa920641c032c3db42eb1fb018eec611dbef96f0fa4dbdfa6fe3cb1b335aed3c",
    },
    PyodideAssetSpec {
        file_name: "python_stdlib.zip",
        sha256: "52866039fa3097e549649b9a62ffae8a1125f01ace7b2d077f34e3cbaff8d0ca",
    },
    PyodideAssetSpec {
        file_name: "pyodide-lock.json",
        sha256: "6526dae570ab7db75019fe2c7ccc6b7b82765c56417a498a7b57e1aaebec39f5",
    },
];

#[derive(Clone, Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PyodideDownloadProgress {
    pub kind: PyodideDownloadProgressKind,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub file_name: Option<String>,
    pub completed_files: u32,
    pub total_files: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bytes_downloaded: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bytes_total: Option<u64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub message: Option<String>,
}

#[derive(Clone, Copy, Debug, Serialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub enum PyodideDownloadProgressKind {
    Checking,
    DownloadStart,
    DownloadProgress,
    DownloadComplete,
    Ready,
}

pub fn pyodide_version_tag() -> String {
    format!("v{PYODIDE_VERSION}")
}

pub fn pyodide_index_url() -> String {
    // Use the host segment to encode the version so requests map cleanly to:
    // `<cache-root>/<host>/<path>`.
    format!("pyodide://{}/full/", pyodide_version_tag())
}

fn pyodide_cdn_base_url() -> String {
    format!("{PYODIDE_CDN_BASE_URL}v{PYODIDE_VERSION}/full/")
}

fn default_pyodide_cache_root() -> Option<PathBuf> {
    let proj = ProjectDirs::from("com", "formula", "Formula")?;
    Some(proj.data_local_dir().join("pyodide"))
}

pub fn pyodide_cache_root() -> Result<PathBuf, String> {
    default_pyodide_cache_root().ok_or_else(|| "could not determine app data directory".to_string())
}

pub fn pyodide_cache_dir() -> Result<PathBuf, String> {
    Ok(pyodide_cache_root()?.join(pyodide_version_tag()).join("full"))
}

/// Returns true if `path` is safe to serve from the `pyodide://` protocol.
///
/// Security goal: prevent arbitrary file reads by ensuring the requested path resolves within the
/// app-controlled Pyodide cache directory (including after resolving symlinks).
pub fn pyodide_cache_path_is_allowed(path: &Path, cache_root: &Path) -> bool {
    if !path.starts_with(cache_root) {
        return false;
    }

    // Canonicalize the cache root so we compare paths in a stable format (and so symlinked roots
    // can't be used to bypass the scope check).
    let Ok(cache_root) = dunce::canonicalize(cache_root) else {
        return false;
    };

    // If the target exists, canonicalize it directly to ensure symlinks can't escape.
    if let Ok(canon) = dunce::canonicalize(path) {
        return canon.starts_with(&cache_root);
    }

    // If the file doesn't exist, fall back to canonicalizing the parent directory so we still
    // detect symlink escapes in the directory chain. Allowing this keeps missing files behaving as
    // 404s (instead of 403s) when they are within the cache scope.
    let Some(parent) = path.parent() else {
        return false;
    };
    let Ok(parent) = dunce::canonicalize(parent) else {
        return false;
    };
    if !parent.starts_with(&cache_root) {
        return false;
    }

    // Treat dangling symlinks as out-of-scope. If the entry exists as a symlink but its target is
    // missing/unreadable, we deny rather than serving it as an allowed missing file.
    if let Ok(meta) = std::fs::symlink_metadata(path) {
        if meta.file_type().is_symlink() {
            return false;
        }
    }

    true
}

fn validate_relative_path(path: &str) -> bool {
    if path.is_empty() || path.contains('\0') {
        return false;
    }

    // Reject any explicit `.` or `..` segments and any absolute-path markers (Windows drive
    // prefixes, Unix root, etc).
    for c in Path::new(path).components() {
        match c {
            Component::Normal(_) => {}
            _ => return false,
        }
    }

    true
}

fn sha256_bytes(bytes: &[u8]) -> String {
    let mut hasher = Sha256::new();
    hasher.update(bytes);
    hex::encode(hasher.finalize())
}

fn sha256_file(path: &Path) -> Result<String, String> {
    let mut file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    let mut hasher = Sha256::new();
    // Stream the hash computation to avoid large allocations for big Pyodide assets.
    let mut buf = [0u8; 1024 * 1024];
    loop {
        let n = file.read(&mut buf).map_err(|e| e.to_string())?;
        if n == 0 {
            break;
        }
        hasher.update(&buf[..n]);
    }
    Ok(hex::encode(hasher.finalize()))
}

#[cfg(windows)]
fn is_windows_reparse_point(meta: &std::fs::Metadata) -> bool {
    use std::os::windows::fs::MetadataExt;
    const FILE_ATTRIBUTE_REPARSE_POINT: u32 = 0x0000_0400;
    meta.file_attributes() & FILE_ATTRIBUTE_REPARSE_POINT != 0
}

#[cfg(not(windows))]
fn is_windows_reparse_point(_meta: &std::fs::Metadata) -> bool {
    false
}

fn file_has_expected_hash(path: &Path, expected_sha256: &str) -> Result<bool, String> {
    // Use `symlink_metadata` so we can treat symlinks as invalid cache entries without ever
    // following them (defense-in-depth: avoid symlink-based scope escapes and avoid hashing
    // attacker-controlled targets outside the cache directory).
    let meta = match std::fs::symlink_metadata(path) {
        Ok(m) => m,
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => return Ok(false),
        Err(err) => return Err(err.to_string()),
    };

    if meta.file_type().is_symlink() {
        return Ok(false);
    }

    // On Windows, directory junctions are not always classified as symlinks, but they are still
    // reparse points that can escape the cache scope. Treat them as invalid cache entries.
    if is_windows_reparse_point(&meta) {
        return Ok(false);
    }

    if !meta.is_file() || meta.len() == 0 {
        return Ok(false);
    }

    if meta.len() > MAX_SINGLE_PYODIDE_ASSET_BYTES as u64 {
        return Err(format!(
            "cached Pyodide asset is too large to hash safely (limit {} bytes, size {} bytes): {}",
            MAX_SINGLE_PYODIDE_ASSET_BYTES,
            meta.len(),
            path.display()
        ));
    }

    let actual = sha256_file(path)?;
    Ok(actual == expected_sha256)
}

async fn remove_existing_cache_entry(path: &Path) -> Result<(), String> {
    match tokio::fs::symlink_metadata(path).await {
        Ok(meta) => {
            // Avoid following symlinks/reparse points. On Windows, directory junctions can appear
            // as directories and `remove_dir_all()` would recurse into the target, so remove the
            // entry itself instead.
            let is_link_like = meta.file_type().is_symlink() || is_windows_reparse_point(&meta);

            let res = if is_link_like {
                if meta.is_dir() {
                    tokio::fs::remove_dir(path).await
                } else {
                    tokio::fs::remove_file(path).await
                }
            } else if meta.is_dir() {
                remove_dir_all_without_following_links(path).await
            } else {
                tokio::fs::remove_file(path).await
            };

            res.map_err(|e| e.to_string())
        }
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(err) => Err(err.to_string()),
    }
}

async fn remove_dir_all_without_following_links(path: &Path) -> std::io::Result<()> {
    use std::io::ErrorKind;

    let mut dirs_to_remove: Vec<PathBuf> = Vec::new();
    let mut stack: Vec<PathBuf> = vec![path.to_path_buf()];

    while let Some(dir) = stack.pop() {
        let mut rd = match tokio::fs::read_dir(&dir).await {
            Ok(rd) => rd,
            Err(err) if err.kind() == ErrorKind::NotFound => continue,
            Err(err) => return Err(err),
        };

        // Remove children before the directory itself (post-order).
        dirs_to_remove.push(dir.clone());

        while let Some(entry) = rd.next_entry().await? {
            let entry_path = entry.path();
            let meta = match tokio::fs::symlink_metadata(&entry_path).await {
                Ok(meta) => meta,
                Err(err) if err.kind() == ErrorKind::NotFound => continue,
                Err(err) => return Err(err),
            };

            let is_link_like = meta.file_type().is_symlink() || is_windows_reparse_point(&meta);
            if meta.is_dir() && !is_link_like {
                stack.push(entry_path);
                continue;
            }

            let res = if meta.is_dir() {
                // Directory symlink/junction: remove the entry itself without recursing.
                tokio::fs::remove_dir(&entry_path).await
            } else {
                tokio::fs::remove_file(&entry_path).await
            };

            match res {
                Ok(()) => {}
                Err(err) if err.kind() == ErrorKind::NotFound => {}
                Err(err) => return Err(err),
            }
        }
    }

    for dir in dirs_to_remove.into_iter().rev() {
        match tokio::fs::remove_dir(&dir).await {
            Ok(()) => {}
            Err(err) if err.kind() == ErrorKind::NotFound => {}
            Err(err) => return Err(err),
        }
    }

    Ok(())
}

async fn download_to_path_with_progress(
    client: &reqwest::Client,
    url: &str,
    dest_path: &Path,
    expected_sha256: &str,
    mut progress: impl FnMut(PyodideDownloadProgress) + Send,
    completed_files: u32,
    total_files: u32,
) -> Result<(), String> {
    let response = client.get(url).send().await.map_err(|e| e.to_string())?;
    if !response.status().is_success() {
        return Err(format!(
            "Failed to download Pyodide asset (HTTP {}): {url}",
            response.status()
        ));
    }

    if let Some(content_length) = response.content_length() {
        if content_length > MAX_SINGLE_PYODIDE_ASSET_BYTES as u64 {
            return Err(format!(
                "Pyodide asset is too large (limit {} bytes, Content-Length {} bytes): {url}",
                MAX_SINGLE_PYODIDE_ASSET_BYTES, content_length
            ));
        }
    }

    // Download to a temp file next to the destination, then rename into place so we never leave
    // partially-written files around if interrupted.
    let parent = dest_path
        .parent()
        .ok_or_else(|| "invalid Pyodide cache path".to_string())?;
    tokio::fs::create_dir_all(parent)
        .await
        .map_err(|e| e.to_string())?;

    let mut tmp_path = None;
    let mut file = None;
    for _ in 0..10 {
        let mut suffix = [0u8; 8];
        rand_core::OsRng.fill_bytes(&mut suffix);
        let suffix = format!("{}-{}", Uuid::new_v4(), hex::encode(suffix));
        let candidate = dest_path.with_extension(format!("tmp-{suffix}"));

        match tokio::fs::OpenOptions::new()
            .create_new(true)
            .write(true)
            .open(&candidate)
            .await
        {
            Ok(f) => {
                tmp_path = Some(candidate);
                file = Some(f);
                break;
            }
            Err(err) if err.kind() == std::io::ErrorKind::AlreadyExists => continue,
            Err(err) => return Err(err.to_string()),
        }
    }
    let tmp_path =
        tmp_path.ok_or_else(|| "failed to create a temporary file for Pyodide download".to_string())?;
    let mut file = file.unwrap();

    let mut hasher = Sha256::new();
    let mut downloaded: u64 = 0;
    let total = response.content_length();
    let mut response = response;

    let mut last_emit = Instant::now() - Duration::from_secs(60);

    loop {
        let Some(chunk) = response.chunk().await.map_err(|e| e.to_string())? else {
            break;
        };

        downloaded = downloaded.saturating_add(chunk.len() as u64);
        if downloaded > MAX_SINGLE_PYODIDE_ASSET_BYTES as u64 {
            let _ = tokio::fs::remove_file(&tmp_path).await;
            return Err(format!(
                "Pyodide asset download exceeded limit (limit {} bytes): {url}",
                MAX_SINGLE_PYODIDE_ASSET_BYTES
            ));
        }

        hasher.update(&chunk);
        file.write_all(&chunk).await.map_err(|e| e.to_string())?;

        // Throttle progress events to keep renderer overhead bounded.
        if last_emit.elapsed() >= Duration::from_millis(200) {
            last_emit = Instant::now();
            progress(PyodideDownloadProgress {
                kind: PyodideDownloadProgressKind::DownloadProgress,
                file_name: Some(dest_path.file_name().unwrap_or_default().to_string_lossy().to_string()),
                completed_files,
                total_files,
                bytes_downloaded: Some(downloaded),
                bytes_total: total,
                message: None,
            });
        }
    }

    file.flush().await.map_err(|e| e.to_string())?;
    drop(file);

    let actual_sha256 = hex::encode(hasher.finalize());
    if actual_sha256 != expected_sha256 {
        let _ = tokio::fs::remove_file(&tmp_path).await;
        return Err(format!(
            "Downloaded Pyodide asset has unexpected sha256 (expected {expected_sha256}, got {actual_sha256}): {url}",
        ));
    }

    // Replace any existing destination entry (Windows rename doesn't overwrite).
    if let Err(err) = remove_existing_cache_entry(dest_path).await {
        let _ = tokio::fs::remove_file(&tmp_path).await;
        return Err(err);
    }

    if let Err(err) = tokio::fs::rename(&tmp_path, dest_path).await {
        let _ = tokio::fs::remove_file(&tmp_path).await;
        return Err(err.to_string());
    }

    Ok(())
}

/// Ensure Pyodide assets are present in `cache_dir`.
///
/// This is intentionally independent of Tauri types so it can be tested in CI without enabling
/// the `desktop` feature (which pulls in the system WebView toolchain on Linux).
pub async fn ensure_pyodide_assets_in_dir(
    cache_dir: &Path,
    cdn_base_url: &str,
    required_files: &[PyodideAssetSpec<'_>],
    download_if_missing: bool,
    mut progress: impl FnMut(PyodideDownloadProgress) + Send,
) -> Result<bool, String> {
    tokio::fs::create_dir_all(cache_dir)
        .await
        .map_err(|e| e.to_string())?;

    let total_files = required_files.len() as u32;

    progress(PyodideDownloadProgress {
        kind: PyodideDownloadProgressKind::Checking,
        file_name: None,
        completed_files: 0,
        total_files,
        bytes_downloaded: None,
        bytes_total: None,
        message: None,
    });

    let mut missing: Vec<PyodideAssetSpec<'_>> = Vec::new();
    for spec in required_files {
        let dest = cache_dir.join(spec.file_name);
        match file_has_expected_hash(&dest, spec.sha256) {
            Ok(true) => {}
            Ok(false) => missing.push(*spec),
            Err(err) => {
                // Corrupt/unreadable file: treat as missing and allow re-download.
                eprintln!("[pyodide] failed to hash cached asset {:?}: {err}", dest);
                missing.push(*spec);
            }
        }
    }

    if missing.is_empty() {
        progress(PyodideDownloadProgress {
            kind: PyodideDownloadProgressKind::Ready,
            file_name: None,
            completed_files: total_files,
            total_files,
            bytes_downloaded: None,
            bytes_total: None,
            message: Some("Pyodide assets ready.".to_string()),
        });
        return Ok(true);
    }

    if !download_if_missing {
        return Ok(false);
    }

    // Security: treat any redirects as untrusted. Only follow redirects within the same origin as
    // `cdn_base_url` (scheme + host + port) so a compromised/malicious CDN cannot bounce us to an
    // unexpected host.
    let parsed_base = reqwest::Url::parse(cdn_base_url).map_err(|e| e.to_string())?;
    let allowed_scheme = parsed_base.scheme().to_string();
    let allowed_host = parsed_base
        .host_str()
        .ok_or_else(|| "Invalid Pyodide CDN base URL (missing host)".to_string())?
        .to_string();
    let allowed_port = parsed_base.port_or_known_default();

    let client = reqwest::Client::builder()
        .redirect(reqwest::redirect::Policy::custom(move |attempt| {
            if attempt.previous().len() >= 10 {
                return attempt.stop();
            }
            let url = attempt.url();
            let same_origin = url.scheme() == allowed_scheme
                && url.host_str() == Some(allowed_host.as_str())
                && url.port_or_known_default() == allowed_port;
            if same_origin {
                attempt.follow()
            } else {
                attempt.stop()
            }
        }))
        .build()
        .map_err(|e| e.to_string())?;

    // Track how many we already had cached for progress reporting.
    let mut completed_files = total_files.saturating_sub(missing.len() as u32);

    for spec in missing {
        if !validate_relative_path(spec.file_name) {
            return Err(format!("Invalid Pyodide asset name: {}", spec.file_name));
        }

        progress(PyodideDownloadProgress {
            kind: PyodideDownloadProgressKind::DownloadStart,
            file_name: Some(spec.file_name.to_string()),
            completed_files,
            total_files,
            bytes_downloaded: None,
            bytes_total: None,
            message: Some(format!(
                "Downloading Python runtime ({}/{})…",
                completed_files + 1,
                total_files
            )),
        });

        let url = format!("{cdn_base_url}{}", spec.file_name);
        let dest = cache_dir.join(spec.file_name);
        download_to_path_with_progress(
            &client,
            &url,
            &dest,
            spec.sha256,
            &mut progress,
            completed_files,
            total_files,
        )
        .await?;

        completed_files = completed_files.saturating_add(1);
        progress(PyodideDownloadProgress {
            kind: PyodideDownloadProgressKind::DownloadComplete,
            file_name: Some(spec.file_name.to_string()),
            completed_files,
            total_files,
            bytes_downloaded: None,
            bytes_total: None,
            message: None,
        });
    }

    progress(PyodideDownloadProgress {
        kind: PyodideDownloadProgressKind::Ready,
        file_name: None,
        completed_files: total_files,
        total_files,
        bytes_downloaded: None,
        bytes_total: None,
        message: Some("Pyodide assets ready.".to_string()),
    });

    Ok(true)
}

/// Resolve the local `pyodide://...` index URL if the cache is present (and valid), optionally
/// downloading missing/corrupt assets.
pub async fn pyodide_index_url_from_cache(
    download_if_missing: bool,
    mut progress: impl FnMut(PyodideDownloadProgress) + Send,
) -> Result<Option<String>, String> {
    let cache_dir = pyodide_cache_dir()?;
    let ok = ensure_pyodide_assets_in_dir(
        &cache_dir,
        &pyodide_cdn_base_url(),
        PYODIDE_REQUIRED_FILES,
        download_if_missing,
        &mut progress,
    )
    .await?;

    if ok {
        Ok(Some(pyodide_index_url()))
    } else {
        Ok(None)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tokio::io::{AsyncReadExt, AsyncWriteExt};
    use tokio::net::TcpListener;

    #[test]
    fn pyodide_cache_path_scope_allows_paths_inside_root() {
        let tmp = tempfile::tempdir().unwrap();
        let root = tmp.path().join("root");
        std::fs::create_dir_all(&root).unwrap();
        let file = root.join("file.txt");
        std::fs::write(&file, b"hello").unwrap();

        assert!(pyodide_cache_path_is_allowed(&file, &root));
        assert!(pyodide_cache_path_is_allowed(&root.join("missing.txt"), &root));
        assert!(!pyodide_cache_path_is_allowed(&tmp.path().join("other.txt"), &root));
    }

    #[cfg(unix)]
    #[test]
    fn pyodide_cache_path_scope_denies_symlink_escape() {
        use std::os::unix::fs::symlink;

        let tmp = tempfile::tempdir().unwrap();
        let root = tmp.path().join("root");
        std::fs::create_dir_all(&root).unwrap();

        let outside = tmp.path().join("outside.txt");
        std::fs::write(&outside, b"secret").unwrap();

        let link = root.join("escape.txt");
        symlink(&outside, &link).unwrap();

        assert!(!pyodide_cache_path_is_allowed(&link, &root));
    }

    async fn serve_files_once(files: Vec<(String, Vec<u8>)>) -> String {
        let listener = TcpListener::bind(("127.0.0.1", 0)).await.unwrap();
        let addr = listener.local_addr().unwrap();

        // Keep the server alive for a handful of requests.
        tokio::spawn(async move {
            let mut remaining = 32usize;
            while remaining > 0 {
                remaining -= 1;
                let Ok((mut socket, _peer)) = listener.accept().await else {
                    return;
                };

                // Read the request line.
                let mut req = Vec::new();
                let mut buf = [0u8; 1024];
                loop {
                    match socket.read(&mut buf).await {
                        Ok(0) => break,
                        Ok(n) => {
                            req.extend_from_slice(&buf[..n]);
                            if req.windows(4).any(|w| w == b"\r\n\r\n") || req.len() > 16 * 1024 {
                                break;
                            }
                        }
                        Err(_) => break,
                    }
                }

                let path = req
                    .split(|b| *b == b' ')
                    .nth(1)
                    .and_then(|s| std::str::from_utf8(s).ok())
                    .unwrap_or("/");
                let name = path.trim_start_matches('/').to_string();

                let body = files
                    .iter()
                    .find(|(n, _)| n == &name)
                    .map(|(_, b)| b.clone());

                match body {
                    Some(body) => {
                        let mut headers = String::from("HTTP/1.1 200 OK\r\n");
                        headers.push_str(&format!("Content-Length: {}\r\n", body.len()));
                        headers.push_str("Connection: close\r\n\r\n");
                        let _ = socket.write_all(headers.as_bytes()).await;
                        let _ = socket.write_all(&body).await;
                    }
                    None => {
                        let _ = socket
                            .write_all(b"HTTP/1.1 404 Not Found\r\nConnection: close\r\n\r\n")
                            .await;
                    }
                }

                let _ = socket.shutdown().await;
            }
        });

        format!("http://{addr}/")
    }

    #[cfg(unix)]
    #[tokio::test]
    async fn symlinked_cached_assets_are_treated_as_missing_and_replaced() {
        use std::os::unix::fs::symlink;

        let files = vec![("a.txt".to_string(), b"hello".to_vec())];
        let base_url = serve_files_once(files.clone()).await;

        let tmp = tempfile::tempdir().unwrap();
        let cache_dir = tmp.path().join("cache");
        std::fs::create_dir_all(&cache_dir).unwrap();

        // Create an out-of-cache target with the *same* content and hash. If we follow the symlink
        // when checking cached assets, we'd treat this as valid and never replace it.
        let outside = tmp.path().join("outside");
        std::fs::create_dir_all(&outside).unwrap();
        let outside_file = outside.join("a.txt");
        std::fs::write(&outside_file, b"hello").unwrap();

        let link = cache_dir.join("a.txt");
        symlink(&outside_file, &link).unwrap();

        let expected_sha = sha256_bytes(b"hello");
        let specs = vec![PyodideAssetSpec {
            file_name: "a.txt",
            sha256: expected_sha.as_str(),
        }];

        let ok = ensure_pyodide_assets_in_dir(&cache_dir, &base_url, &specs, true, |_| {})
            .await
            .unwrap();
        assert!(ok);

        let meta = std::fs::symlink_metadata(&link).unwrap();
        assert!(
            !meta.file_type().is_symlink(),
            "expected cached asset symlink to be replaced with a real file"
        );
    }

    async fn serve_redirect_once(target_base_url: String) -> String {
        let listener = TcpListener::bind(("127.0.0.1", 0)).await.unwrap();
        let addr = listener.local_addr().unwrap();

        tokio::spawn(async move {
            let mut remaining = 8usize;
            while remaining > 0 {
                remaining -= 1;
                let Ok((mut socket, _peer)) = listener.accept().await else {
                    return;
                };

                let mut req = Vec::new();
                let mut buf = [0u8; 1024];
                loop {
                    match socket.read(&mut buf).await {
                        Ok(0) => break,
                        Ok(n) => {
                            req.extend_from_slice(&buf[..n]);
                            if req.windows(4).any(|w| w == b"\r\n\r\n") || req.len() > 16 * 1024 {
                                break;
                            }
                        }
                        Err(_) => break,
                    }
                }

                let path = req
                    .split(|b| *b == b' ')
                    .nth(1)
                    .and_then(|s| std::str::from_utf8(s).ok())
                    .unwrap_or("/");
                let name = path.trim_start_matches('/').to_string();
                let location = format!("{target_base_url}{name}");

                let headers = format!(
                    "HTTP/1.1 302 Found\r\nLocation: {location}\r\nConnection: close\r\n\r\n"
                );
                let _ = socket.write_all(headers.as_bytes()).await;
                let _ = socket.shutdown().await;
            }
        });

        format!("http://{addr}/")
    }

    #[tokio::test]
    async fn directory_cached_assets_are_treated_as_missing_and_replaced() {
        let files = vec![("a.txt".to_string(), b"hello".to_vec())];
        let base_url = serve_files_once(files.clone()).await;

        let tmp = tempfile::tempdir().unwrap();
        let cache_dir = tmp.path().join("cache");
        std::fs::create_dir_all(&cache_dir).unwrap();

        // Simulate a corrupted cache entry where a directory exists at the expected file path.
        let bad_entry = cache_dir.join("a.txt");
        std::fs::create_dir_all(&bad_entry).unwrap();

        let expected_sha = sha256_bytes(b"hello");
        let specs = vec![PyodideAssetSpec {
            file_name: "a.txt",
            sha256: expected_sha.as_str(),
        }];

        let ok = ensure_pyodide_assets_in_dir(&cache_dir, &base_url, &specs, true, |_| {})
            .await
            .unwrap();
        assert!(ok);
        assert!(bad_entry.is_file(), "expected cached directory to be replaced");
        assert_eq!(std::fs::read(&bad_entry).unwrap(), b"hello");
    }

    #[cfg(unix)]
    #[tokio::test]
    async fn directory_cached_assets_do_not_follow_nested_symlinks_when_replaced() {
        use std::os::unix::fs::symlink;

        let files = vec![("a.txt".to_string(), b"hello".to_vec())];
        let base_url = serve_files_once(files.clone()).await;

        let tmp = tempfile::tempdir().unwrap();
        let cache_dir = tmp.path().join("cache");
        std::fs::create_dir_all(&cache_dir).unwrap();

        // Simulate a corrupted cache entry where a directory exists at the expected file path and
        // contains a symlink to an out-of-cache target directory.
        let bad_entry = cache_dir.join("a.txt");
        std::fs::create_dir_all(&bad_entry).unwrap();

        let outside_dir = tmp.path().join("outside_dir");
        std::fs::create_dir_all(&outside_dir).unwrap();
        let outside_file = outside_dir.join("victim.txt");
        std::fs::write(&outside_file, b"secret").unwrap();
        symlink(&outside_dir, bad_entry.join("escape")).unwrap();

        let expected_sha = sha256_bytes(b"hello");
        let specs = vec![PyodideAssetSpec {
            file_name: "a.txt",
            sha256: expected_sha.as_str(),
        }];

        let ok = ensure_pyodide_assets_in_dir(&cache_dir, &base_url, &specs, true, |_| {})
            .await
            .unwrap();
        assert!(ok);
        assert!(bad_entry.is_file(), "expected cached directory to be replaced");
        assert_eq!(std::fs::read(&bad_entry).unwrap(), b"hello");
        assert_eq!(
            std::fs::read(&outside_file).unwrap(),
            b"secret",
            "expected cached directory cleanup to not delete symlink targets"
        );
    }

    #[tokio::test]
    async fn oversized_cached_assets_are_treated_as_missing_and_replaced() {
        let files = vec![("a.txt".to_string(), b"hello".to_vec())];
        let base_url = serve_files_once(files.clone()).await;

        let tmp = tempfile::tempdir().unwrap();
        let cache_dir = tmp.path().join("cache");
        std::fs::create_dir_all(&cache_dir).unwrap();

        // Create an oversized file so `file_has_expected_hash` refuses to hash it.
        let bad_entry = cache_dir.join("a.txt");
        let file = std::fs::File::create(&bad_entry).unwrap();
        file.set_len((MAX_SINGLE_PYODIDE_ASSET_BYTES as u64) + 1).unwrap();
        drop(file);

        let expected_sha = sha256_bytes(b"hello");
        let specs = vec![PyodideAssetSpec {
            file_name: "a.txt",
            sha256: expected_sha.as_str(),
        }];

        let ok = ensure_pyodide_assets_in_dir(&cache_dir, &base_url, &specs, true, |_| {})
            .await
            .unwrap();
        assert!(ok);
        assert!(bad_entry.is_file(), "expected oversized cached file to be replaced");
        assert_eq!(std::fs::read(&bad_entry).unwrap(), b"hello");
    }

    #[tokio::test]
    async fn downloads_and_verifies_sha256() {
        let files = vec![
            ("a.txt".to_string(), b"hello".to_vec()),
            ("b.bin".to_string(), b"world".to_vec()),
        ];

        let base_url = serve_files_once(files.clone()).await;

        let tmp = tempfile::tempdir().unwrap();
        let cache_dir = tmp.path().join("cache");

        let specs: Vec<(String, String)> = files
            .iter()
            .map(|(name, body)| (name.clone(), sha256_bytes(body)))
            .collect();
        let spec_refs: Vec<PyodideAssetSpec<'_>> = specs
            .iter()
            .map(|(name, sha)| PyodideAssetSpec {
                file_name: name.as_str(),
                sha256: sha.as_str(),
            })
            .collect();

        let mut events = Vec::<PyodideDownloadProgressKind>::new();
        let ok = ensure_pyodide_assets_in_dir(
            &cache_dir,
            &base_url,
            &spec_refs,
            true,
            |p| events.push(p.kind.clone()),
        )
        .await
        .unwrap();
        assert!(ok);
        assert!(cache_dir.join("a.txt").is_file());
        assert!(cache_dir.join("b.bin").is_file());

        // Spot-check event ordering.
        assert_eq!(events.first(), Some(&PyodideDownloadProgressKind::Checking));
        assert!(events.iter().any(|k| matches!(k, PyodideDownloadProgressKind::DownloadStart)));
        assert!(events.iter().any(|k| matches!(k, PyodideDownloadProgressKind::Ready)));
    }

    #[tokio::test]
    async fn sha_mismatch_returns_error_and_does_not_cache_file() {
        let files = vec![("bad.txt".to_string(), b"wrong".to_vec())];
        let base_url = serve_files_once(files.clone()).await;

        let tmp = tempfile::tempdir().unwrap();
        let cache_dir = tmp.path().join("cache");

        let specs = vec![PyodideAssetSpec {
            file_name: "bad.txt",
            // Intentionally wrong.
            sha256: "0000000000000000000000000000000000000000000000000000000000000000",
        }];

        let err = ensure_pyodide_assets_in_dir(&cache_dir, &base_url, &specs, true, |_| {})
            .await
            .unwrap_err();
        assert!(err.contains("unexpected sha256"), "unexpected error: {err}");
        assert!(
            !cache_dir.join("bad.txt").exists(),
            "expected corrupt file not to be cached"
        );
    }

    #[tokio::test]
    async fn rejects_redirects_to_other_origins() {
        let target = serve_files_once(vec![("a.txt".to_string(), b"hello".to_vec())]).await;
        let redirect = serve_redirect_once(target.clone()).await;

        let tmp = tempfile::tempdir().unwrap();
        let cache_dir = tmp.path().join("cache");

        let sha = sha256_bytes(b"hello");
        let specs = vec![PyodideAssetSpec {
            file_name: "a.txt",
            sha256: &sha,
        }];

        let err = ensure_pyodide_assets_in_dir(&cache_dir, &redirect, &specs, true, |_| {})
            .await
            .unwrap_err();
        assert!(err.contains("HTTP 302"), "unexpected error: {err}");
        assert!(!cache_dir.join("a.txt").exists());
    }
}
