use std::collections::HashSet;
use std::path::{Path, PathBuf};

use crate::file_io::looks_like_workbook;
use crate::open_file_ipc::{
    MAX_PENDING_BYTES as MAX_OPEN_FILE_PENDING_BYTES,
    MAX_PENDING_PATHS as MAX_OPEN_FILE_PENDING_PATHS,
};
use url::Url;

const SUPPORTED_EXTENSIONS: &[&str] = &[
    "xlsx",
    "xls",
    "xlt",
    "xla",
    "xlsm",
    "xltx",
    "xltm",
    "xlam",
    "xlsb",
    "csv",
    #[cfg(feature = "parquet")]
    "parquet",
];

/// Extract supported spreadsheet paths from a process argv list.
///
/// This is used for:
/// - cold start: initial `std::env::args()`
/// - warm start: argv forwarded by the single-instance plugin
/// - macOS open-document events (converted to `file://...` strings)
///
/// Normalization rules:
/// - accepts: `xlsx`, `xls`, `xlt`, `xla`, `xlsm`, `xltx`, `xltm`, `xlam`, `xlsb`, `csv` (case-insensitive)
///   - plus `parquet` when compiled with the `parquet` feature
/// - if the extension is missing/unsupported, performs a lightweight file signature sniff to
///   detect valid workbooks (OLE `.xls`, ZIP-based `.xlsx`/`.xlsm`/`.xlsb`, plus Parquet when the
///   `parquet` feature is enabled) so downloads and renamed files can still be opened via OS
///   open-file events
/// - handles `file://...` URLs (via [`Url::to_file_path`])
/// - resolves relative paths using `cwd` when provided (falls back to `std::env::current_dir()`)
/// - ignores args that look like flags (start with `-`)
/// - caps the number of returned paths so a huge argv list cannot cause unbounded allocations
pub fn extract_open_file_paths_from_argv(argv: &[String], cwd: Option<&Path>) -> Vec<PathBuf> {
    // Treat argv as untrusted: a malicious sender can invoke the app with a huge argv list to
    // force unbounded allocations and/or expensive file signature sniffing. Bound the output to
    // the same cap as the pending open-file IPC queue.
    let mut out_rev = Vec::with_capacity(MAX_OPEN_FILE_PENDING_PATHS.min(argv.len()));
    for arg in argv.iter().rev() {
        if out_rev.len() >= MAX_OPEN_FILE_PENDING_PATHS {
            break;
        }
        let Some(path) = normalize_open_file_candidate(arg, cwd) else {
            continue;
        };
        out_rev.push(path);
    }
    out_rev.reverse();
    out_rev
}

/// Normalize an "open file" request payload.
///
/// The OS can provide file-open events via argv (cold start) or via the single-instance plugin
/// (warm start). Those payloads should be treated as untrusted: they can be arbitrarily large if a
/// malicious sender invokes the app with a huge argv list.
///
/// This function:
/// - drops empty/whitespace-only entries
/// - de-dupes entries (best-effort)
/// - bounds the output size to keep memory usage deterministic
///
/// When the cap is exceeded, we drop the **oldest** paths and keep the most recent ones ("latest
/// user action wins"). The caps are aligned with the pending open-file IPC queue enforced by
/// [`crate::open_file_ipc::OpenFileState`].
pub fn normalize_open_file_request_paths(paths: Vec<String>) -> Vec<String> {
    let mut seen = HashSet::<String>::with_capacity(
        MAX_OPEN_FILE_PENDING_PATHS.min(paths.len()),
    );
    let mut out_rev = Vec::with_capacity(MAX_OPEN_FILE_PENDING_PATHS.min(paths.len()));
    let mut bytes = 0usize;

    // Walk backwards so we keep the most recent file-open requests, then reverse at the end to
    // preserve the original order among kept paths.
    for path in paths.into_iter().rev() {
        if path.trim().is_empty() {
            continue;
        }

        if seen.contains(&path) {
            continue;
        }

        let len = path.len();
        if len > MAX_OPEN_FILE_PENDING_BYTES {
            // Single oversized entry; skip rather than exceeding the deterministic cap.
            continue;
        }

        if out_rev.len() >= MAX_OPEN_FILE_PENDING_PATHS {
            break;
        }

        if bytes.saturating_add(len) > MAX_OPEN_FILE_PENDING_BYTES {
            // Adding this (older) entry would exceed the byte cap; keep scanning older entries in
            // case smaller ones still fit.
            continue;
        }

        bytes += len;
        seen.insert(path.clone());
        out_rev.push(path);
    }

    out_rev.reverse();
    out_rev
}

fn normalize_open_file_candidate(arg: &str, cwd: Option<&Path>) -> Option<PathBuf> {
    let arg = arg.trim().trim_matches('"');
    if arg.is_empty() {
        return None;
    }

    if arg.len() > MAX_OPEN_FILE_PENDING_BYTES {
        // Treat argv as untrusted; avoid allocating and processing huge strings that will be
        // rejected by the later open-file request normalization anyway.
        return None;
    }

    // Deep links (e.g. OAuth redirects) are delivered via argv on some platforms. Ignore them
    // here so we don't attempt to treat `scheme://...` as a filesystem path.
    if crate::deep_link_schemes::is_deep_link_url(arg) {
        return None;
    }

    // Finder launches can include args like `-psn_0_12345`. Tauri can also be passed flags.
    if arg.starts_with('-') {
        return None;
    }

    let mut path = if arg
        .get(..5)
        .map_or(false, |prefix| prefix.eq_ignore_ascii_case("file:"))
    {
        let url = Url::parse(arg).ok()?;
        if url.scheme() != "file" {
            return None;
        }
        url.to_file_path().ok()?
    } else {
        PathBuf::from(arg)
    };

    if path.is_relative() {
        // Prefer a caller-provided cwd (single-instance plugin supplies one), but fall back to
        // the process cwd if needed so cold-start CLI invocations work.
        let base = cwd
            .map(PathBuf::from)
            .or_else(|| std::env::current_dir().ok())?;
        path = base.join(path);
    }

    if !path.is_absolute() {
        // Shouldn't happen (cwd join above), but keep a best-effort fallback so we never emit
        // relative paths to the frontend.
        let base = std::env::current_dir().ok()?;
        path = base.join(path);
    }

    if let Some(ext) = path.extension().and_then(|ext| ext.to_str()) {
        // Treat OS-delivered paths as untrusted: avoid allocating a full lowercase copy of a
        // potentially large extension string just to compare against our small allowlist.
        if SUPPORTED_EXTENSIONS
            .iter()
            .any(|supported| ext.eq_ignore_ascii_case(supported))
        {
            return Some(path);
        }
    }

    // If the extension is missing or unsupported, attempt a lightweight content sniff so
    // workbooks that were downloaded/renamed without an extension (or with a wrong one) can
    // still be opened via OS open-file events.
    if looks_like_workbook(&path) {
        return Some(path);
    }

    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::BTreeSet;

    use serde_json::Value;
    use tempfile::tempdir;

    #[test]
    fn ignores_flags_and_unsupported_extensions() {
        let dir = tempdir().unwrap();
        let cwd = dir.path();

        let argv = vec![
            "formula-desktop".to_string(),
            "-psn_0_12345".to_string(),
            "--flag".to_string(),
            "notes.txt".to_string(),
            "workbook.xlsx.tmp".to_string(),
        ];

        let paths = extract_open_file_paths_from_argv(&argv, Some(cwd));
        assert!(paths.is_empty());
    }

    #[test]
    fn ignores_formula_deep_links() {
        let argv = vec![
            "formula-desktop".to_string(),
            "formula://oauth/callback?code=abc".to_string(),
            "FORMULA://oauth/callback?code=def".to_string(),
            "\"formula://oauth/callback?code=ghi\"".to_string(),
        ];
        let paths = extract_open_file_paths_from_argv(&argv, None);
        assert!(paths.is_empty());
    }

    #[test]
    fn supported_extensions_cover_tauri_file_associations() {
        let conf = include_str!(concat!(env!("CARGO_MANIFEST_DIR"), "/tauri.conf.json"));
        let conf: Value = serde_json::from_str(conf).expect("parse tauri.conf.json");

        let mut advertised: BTreeSet<String> = BTreeSet::new();
        let associations = conf
            .get("bundle")
            .and_then(|b| b.get("fileAssociations"))
            .and_then(|fa| fa.as_array())
            .expect("tauri.conf.json must contain bundle.fileAssociations");

        for assoc in associations {
            let exts = assoc.get("ext").and_then(|ext| ext.as_array());
            let Some(exts) = exts else { continue };
            for ext in exts {
                let Some(ext) = ext.as_str() else { continue };
                let ext = ext.trim();
                if ext.is_empty() {
                    continue;
                }
                advertised.insert(ext.to_ascii_lowercase());
            }
        }

        // Desktop builds enable Parquet support, but unit tests may compile without the `parquet`
        // feature. Only assert on parquet when it is enabled.
        #[cfg(not(feature = "parquet"))]
        advertised.remove("parquet");

        for ext in advertised {
            assert!(
                SUPPORTED_EXTENSIONS.contains(&ext.as_str()),
                "tauri.conf.json advertises extension `{ext}` but it is missing from SUPPORTED_EXTENSIONS"
            );
        }
    }

    #[test]
    fn resolves_relative_paths_against_cwd_and_filters_extensions_case_insensitively() {
        let dir = tempdir().unwrap();
        let cwd = dir.path();

        let argv = vec![
            "formula-desktop".to_string(),
            "Report.XLSX".to_string(),
            "Report.XLT".to_string(),
            "Report.XLA".to_string(),
            "Report.XLTX".to_string(),
            "Report.XLTM".to_string(),
            "Report.XLAM".to_string(),
            "data.csv".to_string(),
            "ignore.md".to_string(),
        ];

        let paths = extract_open_file_paths_from_argv(&argv, Some(cwd));
        assert_eq!(
            paths,
            vec![
                cwd.join("Report.XLSX"),
                cwd.join("Report.XLT"),
                cwd.join("Report.XLA"),
                cwd.join("Report.XLTX"),
                cwd.join("Report.XLTM"),
                cwd.join("Report.XLAM"),
                cwd.join("data.csv")
            ]
        );
    }

    #[test]
    fn handles_file_urls() {
        let dir = tempdir().unwrap();
        let file_path = dir.path().join("book.xlsm");
        let url = Url::from_file_path(&file_path).unwrap();

        let argv = vec!["formula-desktop".to_string(), url.to_string()];
        let paths = extract_open_file_paths_from_argv(&argv, Some(dir.path()));

        assert_eq!(paths, vec![file_path]);
    }

    #[test]
    fn handles_quoted_file_urls() {
        let dir = tempdir().unwrap();
        let file_path = dir.path().join("book.xlsm");
        let url = Url::from_file_path(&file_path).unwrap();

        let argv = vec![
            "formula-desktop".to_string(),
            format!("\"{}\"", url.to_string()),
        ];
        let paths = extract_open_file_paths_from_argv(&argv, Some(dir.path()));

        assert_eq!(paths, vec![file_path]);
    }

    #[test]
    fn extract_open_file_paths_from_argv_caps_by_count_dropping_oldest() {
        let dir = tempdir().unwrap();
        let cwd = dir.path();

        let mut argv = vec!["formula-desktop".to_string()];
        for idx in 0..(MAX_OPEN_FILE_PENDING_PATHS + 7) {
            argv.push(format!("p{idx}.xlsx"));
        }

        let paths = extract_open_file_paths_from_argv(&argv, Some(cwd));
        assert_eq!(paths.len(), MAX_OPEN_FILE_PENDING_PATHS);

        let expected: Vec<PathBuf> = (7..(MAX_OPEN_FILE_PENDING_PATHS + 7))
            .map(|idx| cwd.join(format!("p{idx}.xlsx")))
            .collect();
        assert_eq!(paths, expected);
    }

    #[test]
    fn accepts_workbook_with_unknown_extension_via_sniffing() {
        let dir = tempdir().unwrap();
        let cwd = dir.path();

        let fixture_path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../../../fixtures/xlsx/basic/basic.xlsx");
        let renamed = cwd.join("basic.bin");
        std::fs::copy(&fixture_path, &renamed).expect("copy fixture");

        let argv = vec!["formula-desktop".to_string(), "basic.bin".to_string()];
        let paths = extract_open_file_paths_from_argv(&argv, Some(cwd));
        assert_eq!(paths, vec![renamed]);
    }

    #[test]
    fn accepts_workbook_without_extension_via_sniffing() {
        let dir = tempdir().unwrap();
        let cwd = dir.path();

        let fixture_path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../../../fixtures/xlsx/basic/basic.xlsx");
        let renamed = cwd.join("basic");
        std::fs::copy(&fixture_path, &renamed).expect("copy fixture");

        let argv = vec!["formula-desktop".to_string(), "basic".to_string()];
        let paths = extract_open_file_paths_from_argv(&argv, Some(cwd));
        assert_eq!(paths, vec![renamed]);
    }

    #[cfg(feature = "parquet")]
    #[test]
    fn parquet_extension_is_supported_when_feature_enabled() {
        let dir = tempdir().unwrap();
        let cwd = dir.path();

        let argv = vec![
            "formula-desktop".to_string(),
            "data.parquet".to_string(),
            "other.csv".to_string(),
        ];

        let paths = extract_open_file_paths_from_argv(&argv, Some(cwd));
        assert_eq!(paths, vec![cwd.join("data.parquet"), cwd.join("other.csv")]);
    }

    #[cfg(feature = "parquet")]
    #[test]
    fn parquet_without_extension_is_supported_via_sniffing_when_feature_enabled() {
        let dir = tempdir().unwrap();
        let cwd = dir.path();

        let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../../packages/data-io/test/fixtures/simple.parquet");
        let renamed = cwd.join("simple");
        std::fs::copy(&fixture_path, &renamed).expect("copy fixture");

        let argv = vec!["formula-desktop".to_string(), "simple".to_string()];
        let paths = extract_open_file_paths_from_argv(&argv, Some(cwd));
        assert_eq!(paths, vec![renamed]);
    }

    #[test]
    fn normalize_open_file_request_paths_dedupes_drops_empty_and_keeps_most_recent() {
        let out = normalize_open_file_request_paths(vec![
            "a.xlsx".to_string(),
            "  ".to_string(),
            "b.csv".to_string(),
            "a.xlsx".to_string(),
        ]);

        // Keep the last occurrence of duplicates while preserving original order among kept items.
        assert_eq!(out, vec!["b.csv".to_string(), "a.xlsx".to_string()]);
    }

    #[test]
    fn normalize_open_file_request_paths_caps_by_count_dropping_oldest() {
        let paths: Vec<String> = (0..(MAX_OPEN_FILE_PENDING_PATHS + 7))
            .map(|idx| format!("p{idx}"))
            .collect();
        let out = normalize_open_file_request_paths(paths);

        assert_eq!(out.len(), MAX_OPEN_FILE_PENDING_PATHS);
        let expected: Vec<String> = (7..(MAX_OPEN_FILE_PENDING_PATHS + 7))
            .map(|idx| format!("p{idx}"))
            .collect();
        assert_eq!(out, expected);
    }

    #[test]
    fn normalize_open_file_request_paths_caps_by_total_bytes_dropping_oldest_deterministically() {
        // Use fixed-size strings so the expected trim point is deterministic.
        let entry_len = 8192;
        let payload = "x".repeat(entry_len - 4); // leave room for "{i:03}-"
        let paths: Vec<String> = (0..MAX_OPEN_FILE_PENDING_PATHS)
            .map(|i| format!("{i:03}-{payload}"))
            .collect();

        let out = normalize_open_file_request_paths(paths.clone());

        let expected_len = MAX_OPEN_FILE_PENDING_BYTES / entry_len;
        assert_eq!(out.len(), expected_len);

        let total_bytes: usize = out.iter().map(|p| p.len()).sum();
        assert!(
            total_bytes <= MAX_OPEN_FILE_PENDING_BYTES,
            "normalized bytes {total_bytes} exceeded cap {MAX_OPEN_FILE_PENDING_BYTES}"
        );

        let expected = paths[MAX_OPEN_FILE_PENDING_PATHS - expected_len..].to_vec();
        assert_eq!(out, expected);
    }
}
