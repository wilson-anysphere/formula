use std::collections::{BTreeMap, BTreeSet};

use formula_engine::functions::FunctionSpec;

struct ParsedLocaleTsv {
    canonical_keys: BTreeSet<String>,
    localized_keys: BTreeSet<String>,
}

fn inventory_function_names() -> BTreeSet<String> {
    let mut names = BTreeSet::new();
    for spec in inventory::iter::<FunctionSpec> {
        let name = spec.name.to_ascii_uppercase();
        if !names.insert(name.clone()) {
            panic!("duplicate function name registered in formula-engine inventory: {name}");
        }
    }
    names
}

fn parse_locale_tsv(locale_id: &str, raw_tsv: &str) -> ParsedLocaleTsv {
    let mut canonical_keys = BTreeSet::new();
    let mut localized_keys = BTreeSet::new();

    // Track first-seen line numbers so failures identify where duplicates came from.
    let mut canon_first_seen: BTreeMap<String, usize> = BTreeMap::new();
    let mut localized_first_seen: BTreeMap<String, (String, usize)> = BTreeMap::new();

    let mut duplicate_canon: Vec<String> = Vec::new();
    let mut duplicate_localized: Vec<String> = Vec::new();

    for (idx, line) in raw_tsv.lines().enumerate() {
        let line_no = idx + 1;
        let line = line.trim();
        if line.is_empty() || line.starts_with('#') {
            continue;
        }

        let (canon, loc) = line.split_once('\t').unwrap_or_else(|| {
            panic!(
                "invalid locale TSV entry in {locale_id} (expected `Canonical<TAB>Localized`) at line {line_no}: {line:?}"
            )
        });
        let canon = canon.trim();
        let loc = loc.trim();
        if canon.is_empty() || loc.is_empty() {
            panic!(
                "invalid locale TSV entry in {locale_id} (empty key/value) at line {line_no}: {line:?}"
            );
        }

        let canon = canon.to_ascii_uppercase();
        let loc = loc.to_ascii_uppercase();

        if let Some(first_line) = canon_first_seen.get(&canon).copied() {
            duplicate_canon.push(format!("{canon} (lines {first_line} and {line_no})"));
        } else {
            canon_first_seen.insert(canon.clone(), line_no);
        }

        if let Some((prior_canon, prior_line)) = localized_first_seen.get(&loc).cloned() {
            // Multiple canonical functions mapping to the same localized name makes the locale
            // canonicalization ambiguous (`loc_to_canon` collision).
            duplicate_localized.push(format!(
                "{loc} (canon {prior_canon} @ line {prior_line}, canon {canon} @ line {line_no})"
            ));
        } else {
            localized_first_seen.insert(loc.clone(), (canon.clone(), line_no));
        }

        canonical_keys.insert(canon);
        localized_keys.insert(loc);
    }

    if !duplicate_canon.is_empty() || !duplicate_localized.is_empty() {
        let mut report =
            format!("locale TSV for {locale_id} contains duplicate entries (case-insensitive)\n");

        if !duplicate_canon.is_empty() {
            report.push_str("\nDuplicate canonical keys:\n");
            for entry in duplicate_canon {
                report.push_str(&format!("  - {entry}\n"));
            }
        }

        if !duplicate_localized.is_empty() {
            report.push_str("\nDuplicate localized keys (collisions):\n");
            for entry in duplicate_localized {
                report.push_str(&format!("  - {entry}\n"));
            }
        }

        panic!("{report}");
    }

    ParsedLocaleTsv {
        canonical_keys,
        localized_keys,
    }
}

#[test]
fn locale_function_tsvs_are_complete_and_unique() {
    let inventory_names = inventory_function_names();

    let locale_tables = [
        ("de-DE", include_str!("../src/locale/data/de-DE.tsv")),
        ("fr-FR", include_str!("../src/locale/data/fr-FR.tsv")),
        ("es-ES", include_str!("../src/locale/data/es-ES.tsv")),
    ];

    let mut failures = String::new();

    for (locale_id, tsv) in locale_tables {
        let parsed = parse_locale_tsv(locale_id, tsv);

        let missing: Vec<String> = inventory_names
            .difference(&parsed.canonical_keys)
            .cloned()
            .collect();
        let extra: Vec<String> = parsed
            .canonical_keys
            .difference(&inventory_names)
            .cloned()
            .collect();

        if missing.is_empty() && extra.is_empty() {
            continue;
        }

        failures.push_str(&format!(
            "\nLocale TSV {locale_id} is out of sync with formula-engine function inventory\n"
        ));
        failures.push_str(&format!(
            "  TSV path: crates/formula-engine/src/locale/data/{locale_id}.tsv\n"
        ));

        if !missing.is_empty() {
            failures.push_str("\n  Missing canonical function keys:\n");
            for name in &missing {
                failures.push_str(&format!("    - {name}\n"));
            }
        }

        if !extra.is_empty() {
            failures.push_str("\n  Extra canonical function keys (not present in inventory):\n");
            for name in &extra {
                failures.push_str(&format!("    - {name}\n"));
            }
        }

        // Also assert localized identifiers are unique within a locale. This is enforced by
        // `parse_locale_tsv` above, but include counts here for easier debugging of coverage.
        failures.push_str(&format!(
            "\n  Parsed entries: canonical={} localized={}\n",
            parsed.canonical_keys.len(),
            parsed.localized_keys.len()
        ));
    }

    if !failures.is_empty() {
        let mut report = String::from(
            "Locale function TSVs are out of sync with the engine's function registry.\n",
        );
        report.push_str("\nWhen adding new built-in functions, update each locale TSV so it contains exactly one entry per canonical function name.\n");
        report.push_str(&failures);
        panic!("{report}");
    }
}
