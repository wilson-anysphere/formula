use std::collections::BTreeSet;
use std::env;
use std::fs;
use std::path::{Path, PathBuf};

use formula_engine::{locale, ErrorKind};

fn usage() -> &'static str {
    r#"Export locale error literal mappings to TSV files (crates/formula-engine/src/locale/data/*.errors.tsv).

Usage:
  cargo run -p formula-engine --bin generate_locale_error_tsv -- [--check]

Options:
  --check   Exit non-zero if any TSV would change.
"#
}

fn normalize_newlines(s: &str) -> String {
    s.replace("\r\n", "\n")
}

fn locale_data_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("src")
        .join("locale")
        .join("data")
}

fn function_tsv_locale_ids(dir: &Path) -> Vec<String> {
    let mut out = Vec::new();
    for entry in fs::read_dir(dir).unwrap_or_else(|err| {
        panic!(
            "failed to list locale data directory {}: {err}",
            dir.display()
        )
    }) {
        let entry = entry.unwrap_or_else(|err| {
            panic!(
                "failed to read entry under locale data directory {}: {err}",
                dir.display()
            )
        });
        let path = entry.path();
        if path.extension().and_then(|s| s.to_str()) != Some("tsv") {
            continue;
        }
        let Some(file_name) = path.file_name().and_then(|s| s.to_str()) else {
            continue;
        };
        if file_name.ends_with(".errors.tsv") {
            continue;
        }
        let Some(stem) = file_name.strip_suffix(".tsv") else {
            continue;
        };
        out.push(stem.to_string());
    }
    out.sort();
    out
}

fn all_error_kinds() -> [ErrorKind; 14] {
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
}

fn render_error_tsv(locale: &locale::FormulaLocale) -> String {
    let mut out = String::new();
    out.push_str("# Canonical\tLocalized\n");
    out.push_str("# See `src/locale/data/README.md` for format + generators.\n\n");

    // Keep the TSV deterministically ordered for stable diffs.
    let mut kinds: Vec<ErrorKind> = all_error_kinds().into();
    kinds.sort_by(|a, b| a.as_code().cmp(b.as_code()));
    for kind in kinds {
        let canon = kind.as_code();
        let localized = locale.localized_error_literal(canon).unwrap_or(canon);
        out.push_str(canon);
        out.push('\t');
        out.push_str(localized);
        out.push('\n');
    }

    out
}

fn main() {
    let args: BTreeSet<String> = env::args().skip(1).collect();
    if args.contains("--help") || args.contains("-h") {
        print!("{}", usage());
        return;
    }
    let check = args.contains("--check");

    let data_dir = locale_data_dir();
    let locale_ids = function_tsv_locale_ids(&data_dir);
    if locale_ids.is_empty() {
        eprintln!(
            "no locale function TSVs found under {} (expected e.g. de-DE.tsv)",
            data_dir.display()
        );
        std::process::exit(1);
    }

    let mut dirty = false;
    for id in locale_ids {
        let Some(locale) = locale::get_locale(&id) else {
            panic!(
                "found {}.tsv in locale data directory, but locale {id:?} is not registered in src/locale/registry.rs",
                id
            );
        };

        let expected = render_error_tsv(locale);
        let path = data_dir.join(format!("{id}.errors.tsv"));

        let actual = fs::read_to_string(&path).unwrap_or_default();
        if normalize_newlines(&actual) != normalize_newlines(&expected) {
            dirty = true;
            if check {
                eprintln!(
                    "{} is out of date. Run:\n  cargo run -p formula-engine --bin generate_locale_error_tsv\n",
                    path.display()
                );
            } else {
                fs::write(&path, expected)
                    .unwrap_or_else(|err| panic!("failed to write {}: {err}", path.display()));
            }
        }
    }

    if check && dirty {
        std::process::exit(1);
    }
}
