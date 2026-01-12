use std::path::{Path, PathBuf};

use crate::file_io::looks_like_workbook;
use url::Url;

const SUPPORTED_EXTENSIONS: &[&str] = &[
    "xlsx",
    "xls",
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
/// - accepts: `xlsx`, `xls`, `xlsm`, `xltx`, `xltm`, `xlam`, `xlsb`, `csv` (case-insensitive)
///   - plus `parquet` when compiled with the `parquet` feature
/// - if the extension is missing/unsupported, performs a lightweight file signature sniff to
///   detect valid workbooks (OLE `.xls`, ZIP-based `.xlsx`/`.xlsm`/`.xlsb`, plus Parquet when the
///   `parquet` feature is enabled) so downloads and renamed files can still be opened via OS
///   open-file events
/// - handles `file://...` URLs (via [`Url::to_file_path`])
/// - resolves relative paths using `cwd` when provided (falls back to `std::env::current_dir()`)
/// - ignores args that look like flags (start with `-`)
pub fn extract_open_file_paths_from_argv(argv: &[String], cwd: Option<&Path>) -> Vec<PathBuf> {
    argv.iter()
        .filter_map(|arg| normalize_open_file_candidate(arg, cwd))
        .collect()
}

fn normalize_open_file_candidate(arg: &str, cwd: Option<&Path>) -> Option<PathBuf> {
    let arg = arg.trim();
    if arg.is_empty() {
        return None;
    }

    // Deep links (e.g. OAuth redirects) are delivered via argv on some platforms. Ignore them
    // here so we don't attempt to treat `formula://...` as a filesystem path.
    if arg.to_ascii_lowercase().starts_with("formula:") {
        return None;
    }

    // Finder launches can include args like `-psn_0_12345`. Tauri can also be passed flags.
    if arg.starts_with('-') {
        return None;
    }

    let mut path = if arg.to_ascii_lowercase().starts_with("file:") {
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

    let ext = path
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_ascii_lowercase());

    if let Some(ext) = ext {
        if SUPPORTED_EXTENSIONS.contains(&ext.as_str()) {
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
}
