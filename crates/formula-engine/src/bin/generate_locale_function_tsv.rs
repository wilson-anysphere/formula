use std::collections::{BTreeMap, BTreeSet};
use std::env;
use std::fs;
use std::path::{Path, PathBuf};

use serde::Deserialize;

#[derive(Debug, Deserialize)]
struct FunctionCatalog {
    functions: Vec<CatalogFunction>,
}

#[derive(Debug, Deserialize)]
struct CatalogFunction {
    name: String,
}

fn usage() -> &'static str {
    r#"Regenerate locale function TSVs (crates/formula-engine/src/locale/data/*.tsv).

Usage:
  cargo run -p formula-engine --bin generate_locale_function_tsv -- [--check]

Options:
  --check   Exit non-zero if any TSV would change.
"#
}

fn normalize_newlines(s: &str) -> String {
    s.replace("\r\n", "\n")
}

fn casefold_ident(ident: &str) -> String {
    if ident.is_ascii() {
        ident.to_ascii_uppercase()
    } else {
        ident.chars().flat_map(|ch| ch.to_uppercase()).collect()
    }
}

fn read_engine_catalog() -> Vec<String> {
    let manifest_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let catalog_path = manifest_dir
        .join("..")
        .join("..")
        .join("shared")
        .join("functionCatalog.json");
    let raw = fs::read_to_string(&catalog_path).unwrap_or_else(|err| {
        panic!(
            "failed to read function catalog at {}: {err}",
            catalog_path.display()
        )
    });

    let catalog: FunctionCatalog = serde_json::from_str(&raw).unwrap_or_else(|err| {
        panic!(
            "failed to parse function catalog JSON at {}: {err}",
            catalog_path.display()
        )
    });

    // functionCatalog.json is expected to be uppercase + sorted; re-normalize to be safe.
    let mut names: Vec<String> = catalog
        .functions
        .into_iter()
        .map(|f| f.name.to_ascii_uppercase())
        .collect();
    names.sort();
    names.dedup();
    names
}

fn locale_data_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("src")
        .join("locale")
        .join("data")
}

fn function_tsv_paths(dir: &Path) -> Vec<PathBuf> {
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
        out.push(path);
    }
    out.sort();
    out
}

fn parse_existing_tsv(src: &str) -> (Vec<String>, BTreeMap<String, String>) {
    let mut header_lines = Vec::new();
    let mut entries = BTreeMap::new();
    let mut in_header = true;

    for raw_line in src.lines() {
        let line = raw_line.trim_end();
        let trimmed = line.trim();

        if in_header && (trimmed.is_empty() || trimmed.starts_with('#')) {
            header_lines.push(line.to_string());
            continue;
        }
        in_header = false;

        if trimmed.is_empty() || trimmed.starts_with('#') {
            continue;
        }

        let (canon, loc) = trimmed.split_once('\t').unwrap_or_else(|| {
            panic!("invalid TSV line (expected Canonical<TAB>Localized): {trimmed:?}")
        });
        let canon = canon.trim();
        let loc = loc.trim();
        if canon.is_empty() || loc.is_empty() {
            panic!("invalid TSV line (empty entry): {trimmed:?}");
        }

        if entries.insert(canon.to_string(), loc.to_string()).is_some() {
            panic!("duplicate canonical key in TSV (case-sensitive): {canon}");
        }
    }

    (header_lines, entries)
}

fn render_tsv(
    header_lines: &[String],
    catalog: &[String],
    existing: &BTreeMap<String, String>,
) -> String {
    const README_HEADER: &str = "# See `src/locale/data/README.md` for format + generators.";

    // Always start with the standard header; preserve any additional comment lines that were
    // already present in the file for provenance.
    let mut out = String::new();
    out.push_str("# Canonical\tLocalized\n");

    for line in header_lines {
        let trimmed = line.trim_end();
        if trimmed.is_empty() {
            continue;
        }
        if trimmed == "# Canonical\tLocalized" {
            continue;
        }
        if trimmed == README_HEADER {
            continue;
        }
        if trimmed.starts_with('#') {
            out.push_str(trimmed);
            out.push('\n');
        }
    }

    out.push_str(README_HEADER);
    out.push_str("\n\n");

    for canonical in catalog {
        let localized = existing
            .get(canonical)
            .cloned()
            .unwrap_or_else(|| canonical.clone());
        let localized = casefold_ident(&localized);
        out.push_str(canonical);
        out.push('\t');
        out.push_str(&localized);
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

    let catalog = read_engine_catalog();
    let data_dir = locale_data_dir();
    let paths = function_tsv_paths(&data_dir);

    if paths.is_empty() {
        eprintln!(
            "no locale function TSVs found under {} (expected e.g. de-DE.tsv)",
            data_dir.display()
        );
        std::process::exit(1);
    }

    let mut dirty = false;
    for path in paths {
        let src = fs::read_to_string(&path)
            .unwrap_or_else(|err| panic!("failed to read {}: {err}", path.display()));
        let (header_lines, entries) = parse_existing_tsv(&src);

        let rendered = render_tsv(&header_lines, &catalog, &entries);
        if normalize_newlines(&src) != normalize_newlines(&rendered) {
            dirty = true;
            if check {
                eprintln!(
                    "{} is out of date. Run:\n  cargo run -p formula-engine --bin generate_locale_function_tsv\n",
                    path.display()
                );
            } else {
                fs::write(&path, rendered)
                    .unwrap_or_else(|err| panic!("failed to write {}: {err}", path.display()));
            }
        }
    }

    if check && dirty {
        std::process::exit(1);
    }
}
