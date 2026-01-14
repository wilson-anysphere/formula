use std::collections::{BTreeSet, HashMap};
use std::fs;
use std::path::{Path, PathBuf};

use formula_engine::locale::get_locale;
use formula_engine::ErrorKind;
use pretty_assertions::assert_eq;

fn expected_error_codes() -> BTreeSet<String> {
    [
        ErrorKind::Null,
        ErrorKind::Div0,
        ErrorKind::Value,
        ErrorKind::Ref,
        ErrorKind::Name,
        ErrorKind::Num,
        ErrorKind::NA,
        ErrorKind::GettingData,
        ErrorKind::Spill,
        ErrorKind::Calc,
        ErrorKind::Field,
        ErrorKind::Connect,
        ErrorKind::Blocked,
        ErrorKind::Unknown,
    ]
    .into_iter()
    .map(|kind| kind.as_code().to_string())
    .collect()
}

fn is_error_tsv_comment_line(line: &str) -> bool {
    let trimmed = line.trim();
    if trimmed.is_empty() {
        return true;
    }
    if trimmed == "#" {
        return true;
    }
    if trimmed.starts_with('#') {
        // Error literals start with `#`, so treat comments as `#` followed by whitespace (or a bare
        // `#`), matching the runtime loader + generator conventions.
        return trimmed
            .chars()
            .nth(1)
            .is_some_and(|c| c.is_whitespace());
    }
    false
}

fn parse_error_tsv(path: &Path) -> Vec<(String, String)> {
    let raw = fs::read_to_string(path)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", path.display()));

    let mut rows = Vec::new();

    for (idx, raw_line) in raw.lines().enumerate() {
        let line_no = idx + 1;
        if is_error_tsv_comment_line(raw_line) {
            continue;
        }

        let trimmed = raw_line.trim();

        // Match the Node generator's parsing behavior by splitting the *raw* line on tabs (so
        // leading/trailing tabs still count as empty columns).
        let mut parts = raw_line.split('\t');
        let canon = parts.next().unwrap_or("");
        let loc = parts.next().unwrap_or_else(|| {
            panic!(
                "invalid error TSV line (expected Canonical<TAB>Localized) at {}:{}: {trimmed:?}",
                path.display(),
                line_no
            )
        });
        if parts.next().is_some() {
            panic!(
                "invalid error TSV line (too many columns) at {}:{}: {trimmed:?}",
                path.display(),
                line_no
            );
        }

        let canon = canon.trim();
        let loc = loc.trim();
        if canon.is_empty() || loc.is_empty() {
            panic!(
                "invalid error TSV line (empty entry) at {}:{}: {trimmed:?}",
                path.display(),
                line_no
            );
        }
        if !canon.starts_with('#') {
            panic!(
                "invalid error TSV line (canonical must start with '#') at {}:{}: {trimmed:?}",
                path.display(),
                line_no
            );
        }
        if !loc.starts_with('#') {
            panic!(
                "invalid error TSV line (localized must start with '#') at {}:{}: {trimmed:?}",
                path.display(),
                line_no
            );
        }

        rows.push((canon.to_string(), loc.to_string()));
    }

    rows
}

fn locale_error_tsv_paths() -> Vec<PathBuf> {
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
        if filename.ends_with(".errors.tsv") {
            paths.push(path);
        }
    }

    paths.sort();
    paths
}

fn locale_id_from_error_tsv_path(path: &Path) -> String {
    let filename = path
        .file_name()
        .and_then(|s| s.to_str())
        .unwrap_or_else(|| panic!("expected UTF-8 path for {}", path.display()));
    filename
        .strip_suffix(".errors.tsv")
        .unwrap_or_else(|| panic!("expected .errors.tsv suffix for {}", path.display()))
        .to_string()
}

#[test]
fn locale_error_tsv_sync() {
    let paths = locale_error_tsv_paths();
    assert!(
        !paths.is_empty(),
        "no locale error TSVs found under src/locale/data"
    );

    let locale_ids: BTreeSet<String> = paths
        .iter()
        .map(|path| locale_id_from_error_tsv_path(path))
        .collect();

    for required in ["de-DE", "fr-FR", "es-ES"] {
        assert!(
            locale_ids.contains(required),
            "expected {required}.errors.tsv to exist under src/locale/data"
        );
    }

    let expected = expected_error_codes();

    for path in paths {
        let locale_id = locale_id_from_error_tsv_path(&path);
        let Some(locale) = get_locale(&locale_id) else {
            panic!(
                "found error TSV for locale {locale_id:?} at {}, but locale is not registered in formula_engine::locale::get_locale",
                path.display()
            );
        };

        let rows = parse_error_tsv(&path);
        let canon_set: BTreeSet<String> = rows.iter().map(|(canon, _)| canon.clone()).collect();

        assert_eq!(
            canon_set, expected,
            "error TSV canonical set mismatch for {locale_id} ({})",
            path.display()
        );

        // The error TSV can contain multiple localized spellings for the same canonical error
        // literal. In those cases, the runtime uses the first spelling as the preferred display
        // form and accepts all spellings for parsing.
        let mut preferred_by_canon: HashMap<String, String> = HashMap::new();
        for (canon, loc) in &rows {
            preferred_by_canon
                .entry(canon.clone())
                .or_insert_with(|| loc.clone());
        }

        for (canon, preferred_loc) in preferred_by_canon {
            let runtime_loc = locale.localized_error_literal(&canon).unwrap_or(canon.as_str());
            assert_eq!(
                runtime_loc, preferred_loc,
                "preferred localized error literal mismatch for locale {locale_id}: canonical={canon:?}"
            );
        }

        for (canon, loc) in rows {
            let runtime_canon = locale.canonical_error_literal(&loc).unwrap_or(loc.as_str());
            assert_eq!(
                runtime_canon, canon,
                "canonical error literal mismatch for locale {locale_id}: localized={loc:?}"
            );
        }
    }
}
