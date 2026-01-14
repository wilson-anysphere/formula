use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};

use formula_engine::locale::get_locale;
use pretty_assertions::assert_eq;

fn casefold(s: &str) -> String {
    // Mirror `formula_engine::value::casefold` / `locale::registry::casefold_ident`.
    // Excel-style identifier matching is case-insensitive across Unicode (`ä` -> `Ä`, `ß` -> `SS`).
    if s.is_ascii() {
        s.to_ascii_uppercase()
    } else {
        s.chars().flat_map(|ch| ch.to_uppercase()).collect()
    }
}

fn parse_function_tsv(path: &Path) -> Vec<(String, String)> {
    let raw = fs::read_to_string(path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", path.display()));

    let mut rows = Vec::new();

    // Track the line that introduced each key so we can produce actionable diagnostics if the TSV
    // contains duplicate entries. Use Unicode-aware case folding to match runtime semantics.
    let mut canon_line: HashMap<String, (usize, String)> = HashMap::new();
    let mut loc_line: HashMap<String, (usize, String)> = HashMap::new();

    for (idx, raw_line) in raw.lines().enumerate() {
        let line_no = idx + 1;
        let trimmed = raw_line.trim();
        if trimmed.is_empty() || trimmed.starts_with('#') {
            continue;
        }

        // Parse the raw line (not the trimmed line) so trailing empty columns like
        // `SUM\tSUMME\t` are not silently accepted.
        let mut parts = raw_line.split('\t');
        let canon = parts.next().unwrap_or("");
        let loc = parts.next().unwrap_or_else(|| {
            panic!(
                "invalid function TSV line (expected Canonical<TAB>Localized) at {}:{}: {raw_line:?}",
                path.display(),
                line_no
            )
        });
        if parts.next().is_some() {
            panic!(
                "invalid function TSV line (too many columns) at {}:{}: {raw_line:?}",
                path.display(),
                line_no
            );
        }
        let canon = canon.trim();
        let loc = loc.trim();
        if canon.is_empty() || loc.is_empty() {
            panic!(
                "invalid function TSV line (empty entry) at {}:{}: {raw_line:?}",
                path.display(),
                line_no
            );
        }

        let canon_key = casefold(canon);
        let loc_key = casefold(loc);

        if let Some((prev_no, prev_line)) = canon_line.get(&canon_key) {
            panic!(
                "duplicate canonical function translation key {canon_key:?}\n  first: {}:{}: {prev_line:?}\n  second: {}:{}: {raw_line:?}",
                path.display(),
                prev_no,
                path.display(),
                line_no
            );
        }
        if let Some((prev_no, prev_line)) = loc_line.get(&loc_key) {
            panic!(
                "duplicate localized function translation key {loc_key:?}\n  first: {}:{}: {prev_line:?}\n  second: {}:{}: {raw_line:?}",
                path.display(),
                prev_no,
                path.display(),
                line_no
            );
        }

        canon_line.insert(canon_key, (line_no, raw_line.to_string()));
        loc_line.insert(loc_key, (line_no, raw_line.to_string()));
        rows.push((canon.to_string(), loc.to_string()));
    }

    rows
}

fn locale_function_tsv_paths() -> Vec<PathBuf> {
    let data_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("src/locale/data");
    let mut paths = Vec::new();

    let entries = fs::read_dir(&data_dir)
        .unwrap_or_else(|err| panic!("failed to list {}: {err}", data_dir.display()));
    for entry in entries {
        let entry = entry.unwrap_or_else(|err| {
            panic!(
                "failed to read directory entry under {}: {err}",
                data_dir.display()
            )
        });
        let path = entry.path();
        if !path.is_file() {
            continue;
        }
        let Some(filename) = path.file_name().and_then(|s| s.to_str()) else {
            continue;
        };
        if !filename.ends_with(".tsv") {
            continue;
        }
        if filename.ends_with(".errors.tsv") {
            continue;
        }
        if filename == "README.md" {
            continue;
        }
        paths.push(path);
    }

    paths.sort();
    paths
}

fn locale_id_from_function_tsv_path(path: &Path) -> String {
    let filename = path
        .file_name()
        .and_then(|s| s.to_str())
        .unwrap_or_else(|| panic!("expected UTF-8 path for {}", path.display()));
    filename
        .strip_suffix(".tsv")
        .unwrap_or_else(|| panic!("expected .tsv suffix for {}", path.display()))
        .to_string()
}

#[test]
fn locale_function_tsv_sync() {
    let paths = locale_function_tsv_paths();
    assert!(
        !paths.is_empty(),
        "no locale function TSVs found under src/locale/data"
    );

    for path in paths {
        let locale_id = locale_id_from_function_tsv_path(&path);
        let Some(locale) = get_locale(&locale_id) else {
            panic!(
                "found function TSV for locale {locale_id:?} at {}, but locale is not registered in formula_engine::locale::get_locale",
                path.display()
            );
        };

        let rows = parse_function_tsv(&path);
        assert!(
            !rows.is_empty(),
            "function TSV for {locale_id} ({}) has no data rows",
            path.display()
        );

        for (canon, loc) in rows {
            let runtime_loc = locale.localized_function_name(&canon);
            assert_eq!(
                runtime_loc, loc,
                "localized function name mismatch for locale {locale_id}: canonical={canon:?}"
            );

            let runtime_canon = locale.canonical_function_name(&loc);
            assert_eq!(
                runtime_canon, canon,
                "canonical function name mismatch for locale {locale_id}: localized={loc:?}"
            );
        }
    }
}
