use std::collections::{BTreeMap, BTreeSet};
use std::path::{Path, PathBuf};

use formula_engine::functions::FunctionSpec;
use formula_engine::locale::get_locale;
use formula_engine::ErrorKind;

fn casefold_key(s: &str) -> String {
    // Match the runtime case-folding used by the locale registry (`value::try_casefold`):
    // ASCII fast path + Unicode-aware uppercasing fallback.
    if s.is_ascii() {
        return s.to_ascii_uppercase();
    }
    s.chars().flat_map(|c| c.to_uppercase()).collect()
}

struct ParsedLocaleTsv {
    canonical_keys: BTreeSet<String>,
    localized_keys: BTreeSet<String>,
}

struct ParsedErrorTsv {
    canonical_keys: BTreeSet<String>,
    localized_keys: BTreeSet<String>,
    entries: BTreeMap<String, Vec<String>>,
}

fn inventory_function_names() -> BTreeSet<String> {
    let mut names = BTreeSet::new();
    for spec in inventory::iter::<FunctionSpec> {
        let name = casefold_key(spec.name);
        if !names.insert(name.clone()) {
            panic!("duplicate function name registered in formula-engine inventory: {name}");
        }
    }
    names
}

fn parse_locale_tsv(locale_id: &str, path: &Path, raw_tsv: &str) -> ParsedLocaleTsv {
    let mut canonical_keys = BTreeSet::new();
    let mut localized_keys = BTreeSet::new();

    // Track first-seen line numbers so failures identify where duplicates came from.
    let mut canon_first_seen: BTreeMap<String, usize> = BTreeMap::new();
    let mut localized_first_seen: BTreeMap<String, (String, usize)> = BTreeMap::new();

    let mut duplicate_canon: Vec<String> = Vec::new();
    let mut duplicate_localized: Vec<String> = Vec::new();

    for (idx, raw_line) in raw_tsv.lines().enumerate() {
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
                "invalid locale TSV entry in {locale_id} ({path}) (expected `Canonical<TAB>Localized`) at line {line_no}: {raw_line:?}",
                path = path.display()
            )
        });
        if parts.next().is_some() {
            panic!(
                "invalid locale TSV entry in {locale_id} ({path}) (too many columns) at line {line_no}: {raw_line:?}",
                path = path.display()
            );
        }
        let canon = canon.trim();
        let loc = loc.trim();
        if canon.is_empty() || loc.is_empty() {
            panic!(
                "invalid locale TSV entry in {locale_id} ({path}) (empty key/value) at line {line_no}: {raw_line:?}",
                path = path.display()
            );
        }

        // Locale translation tables are keyed using the same case folding as runtime parsing:
        // Unicode-aware uppercasing (`char::to_uppercase`), not ASCII-only uppercasing.
        //
        // This ensures collisions like `fü` vs `FÜ` are detected here the same way they'd be at
        // runtime when building `loc_to_canon`.
        let canon = casefold_key(canon);
        let loc = casefold_key(loc);

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
        let mut report = format!(
            "locale TSV for {locale_id} ({path}) contains duplicate entries (case-insensitive)\n",
            path = path.display()
        );

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
fn locale_function_tsv_completeness_detects_unicode_case_collisions_in_localized_keys() {
    let tsv = "FOO\tfü\nBAR\tFÜ\n";
    let path = Path::new("<in-memory>");
    let err = match std::panic::catch_unwind(|| parse_locale_tsv("xx-XX", path, tsv)) {
        Ok(_) => panic!("expected parse_locale_tsv to panic due to localized-key collision"),
        Err(err) => err,
    };
    let msg = err
        .downcast_ref::<String>()
        .map(String::as_str)
        .or_else(|| err.downcast_ref::<&str>().copied())
        .unwrap_or("<non-string panic>");
    assert!(
        msg.contains("Duplicate localized keys"),
        "expected duplicate localized keys report, got: {msg}"
    );
    assert!(
        msg.contains("line 1") && msg.contains("line 2"),
        "expected line numbers in report, got: {msg}"
    );
}

fn parse_error_tsv(
    locale_id: &str,
    path: &Path,
    raw_tsv: &str,
    require_sorted: bool,
) -> ParsedErrorTsv {
    let mut canonical_keys = BTreeSet::new();
    let mut localized_keys = BTreeSet::new();
    let mut entries: BTreeMap<String, Vec<String>> = BTreeMap::new();

    let mut localized_first_seen: BTreeMap<String, (String, usize)> = BTreeMap::new();

    let mut duplicate_localized: Vec<String> = Vec::new();

    let mut prev_canon: Option<String> = None;

    for (idx, raw_line) in raw_tsv.lines().enumerate() {
        let line_no = idx + 1;
        let trimmed = raw_line.trim();

        // Error literals themselves start with `#`, so comments are `#` followed by whitespace.
        let is_comment = trimmed == "#"
            || (trimmed.starts_with('#')
                && trimmed.chars().nth(1).is_some_and(|c| c.is_whitespace()));

        if trimmed.is_empty() || is_comment {
            continue;
        }

        let mut parts = raw_line.split('\t');
        let canon = parts.next().unwrap_or("");
        let loc = parts.next().unwrap_or_else(|| {
            panic!(
                "invalid error TSV entry in {locale_id} ({path}) (expected `Canonical<TAB>Localized`) at line {line_no}: {raw_line:?}",
                path = path.display()
            )
        });
        if parts.next().is_some() {
            panic!(
                "invalid error TSV entry in {locale_id} ({path}) (too many columns) at line {line_no}: {raw_line:?}",
                path = path.display()
            );
        }

        let canon = canon.trim();
        let loc = loc.trim();
        if canon.is_empty() || loc.is_empty() {
            panic!(
                "invalid error TSV entry in {locale_id} ({path}) (empty key/value) at line {line_no}: {raw_line:?}",
                path = path.display()
            );
        }
        if !canon.starts_with('#') || !loc.starts_with('#') {
            panic!(
                "invalid error TSV entry in {locale_id} ({path}) (expected error literals to start with '#') at line {line_no}: {raw_line:?}",
                path = path.display()
            );
        }

        if require_sorted {
            if let Some(prev) = prev_canon.as_ref() {
                if canon < prev.as_str() {
                    panic!(
                        "error TSV for {locale_id} ({path}) is not sorted by canonical key: line {line_no}: {canon:?} comes after {prev:?}",
                        path = path.display()
                    );
                }
            }
            prev_canon = Some(canon.to_string());
        }

        let canon_key = casefold_key(canon);
        let loc_key = casefold_key(loc);

        if let Some((prior_canon, prior_line)) = localized_first_seen.get(&loc_key).cloned() {
            duplicate_localized.push(format!(
                "{loc_key} (canon {prior_canon} @ line {prior_line}, canon {canon_key} @ line {line_no})"
            ));
        } else {
            localized_first_seen.insert(loc_key.clone(), (canon_key.clone(), line_no));
        }

        canonical_keys.insert(canon.to_string());
        localized_keys.insert(loc.to_string());
        entries
            .entry(canon.to_string())
            .or_default()
            .push(loc.to_string());
    }

    if !duplicate_localized.is_empty() {
        let mut report = format!(
            "locale error TSV for {locale_id} ({path}) contains duplicate entries (case-insensitive)\n",
            path = path.display()
        );

        if !duplicate_localized.is_empty() {
            report.push_str("\nDuplicate localized keys (collisions):\n");
            for entry in duplicate_localized {
                report.push_str(&format!("  - {entry}\n"));
            }
        }

        panic!("{report}");
    }

    ParsedErrorTsv {
        canonical_keys,
        localized_keys,
        entries,
    }
}

fn locale_data_dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("src/locale/data")
}

fn upstream_error_data_dir() -> PathBuf {
    locale_data_dir().join("upstream/errors")
}

fn discover_tsvs_in_dir(dir: &Path) -> BTreeMap<String, PathBuf> {
    let mut out = BTreeMap::new();
    let read_dir = std::fs::read_dir(dir).unwrap_or_else(|err| {
        panic!(
            "failed to read locale data directory {}: {err}",
            dir.display()
        )
    });

    for entry in read_dir {
        let entry = entry.unwrap_or_else(|err| {
            panic!(
                "failed to read entry in locale data directory {}: {err}",
                dir.display()
            )
        });
        let file_type = entry
            .file_type()
            .unwrap_or_else(|err| panic!("failed to stat entry {}: {err}", entry.path().display()));
        if !file_type.is_file() {
            continue;
        }

        let path = entry.path();
        let file_name = path
            .file_name()
            .and_then(|s| s.to_str())
            .unwrap_or_else(|| {
                panic!(
                    "locale data directory contains a non-utf8 file name: {}",
                    path.display()
                )
            });

        if !file_name.ends_with(".tsv") {
            continue;
        }
        let locale_id = file_name
            .strip_suffix(".tsv")
            .expect("already checked suffix");
        out.insert(locale_id.to_string(), path);
    }

    out
}

#[test]
fn locale_function_tsv_completeness_function_tsvs_are_complete_and_unique() {
    let inventory_names = inventory_function_names();

    let dir = locale_data_dir();
    let mut locale_tsvs = discover_tsvs_in_dir(&dir);
    // Exclude `*.errors.tsv` and anything that isn't a function translation table.
    locale_tsvs.retain(|locale_id, _path| !locale_id.ends_with(".errors"));

    for required in ["de-DE", "fr-FR", "es-ES"] {
        assert!(
            locale_tsvs.contains_key(required),
            "expected locale function TSV {required}.tsv to exist in {} (discovered: {:?})",
            dir.display(),
            locale_tsvs.keys().collect::<Vec<_>>()
        );
    }

    let mut failures = String::new();

    for (locale_id, path) in &locale_tsvs {
        let tsv = std::fs::read_to_string(path).unwrap_or_else(|err| {
            panic!(
                "failed to read locale function TSV {}: {err}",
                path.display()
            )
        });
        let parsed = parse_locale_tsv(locale_id, path, &tsv);

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
        failures.push_str(&format!("  TSV path: {}\n", path.display()));

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

#[test]
fn de_de_locale_function_tsv_is_not_mostly_identity_mappings() {
    // Regression guard: `de-DE.tsv` is generated from `sources/de-DE.json`. If the source mapping is
    // accidentally replaced with a tiny curated subset, the generator will silently fall back to
    // identity mappings for most functions (canonical == localized), breaking localized editing and
    // round-tripping.
    //
    // Allow some identity mappings since many functions are not localized in German Excel (e.g.
    // `ABS`, `COS`, etc) and some functions may be unavailable in a given Excel build.
    //
    // We use a dual threshold:
    // - A percentage-based threshold to catch "almost everything became English again".
    // - An absolute minimum so the test stays stable if the function catalog grows substantially
    //   (new functions may initially be identity mappings until locale data is refreshed).
    let tsv = include_str!("../src/locale/data/de-DE.tsv");
    let mut total = 0usize;
    let mut identity = 0usize;

    for line in tsv.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() || trimmed.starts_with('#') {
            continue;
        }

        // Parse the raw line (not the trimmed line) so trailing empty columns like
        // `SUM\tSUMME\t` are not silently accepted.
        let mut parts = line.split('\t');
        let canon = parts.next().unwrap_or("");
        let loc = parts.next().unwrap_or_else(|| {
            panic!("invalid TSV line in de-DE.tsv (expected `Canonical<TAB>Localized`): {line:?}")
        });
        if parts.next().is_some() {
            panic!("invalid TSV line in de-DE.tsv (too many columns): {line:?}");
        }
        total += 1;
        let canon = canon.trim();
        let loc = loc.trim();
        if canon.is_empty() || loc.is_empty() {
            panic!("invalid TSV line in de-DE.tsv (empty field): {line:?}");
        }
        if canon == loc {
            identity += 1;
        }
    }

    let non_identity = total - identity;
    let passes_ratio = non_identity * 100 >= total * 60;
    let passes_absolute = non_identity >= 300;
    assert!(
        passes_ratio || passes_absolute,
        "expected de-DE.tsv to contain many localized function spellings; got {non_identity}/{total} non-identity entries (identity={identity})"
    );
}

#[test]
fn es_es_locale_function_tsv_is_not_mostly_identity_mappings() {
    // Regression guard: `es-ES.tsv` is generated from `sources/es-ES.json`. If the source mapping
    // is accidentally replaced with a partial extract (e.g. from an older Excel build that treats
    // many functions as `_xludf.`), the generator will silently fall back to identity mappings
    // (canonical == localized) for many functions.
    //
    // Like `de-DE`, allow some identity mappings since many functions are not localized in Spanish
    // Excel (e.g. `ABS`, `COS`, etc).
    //
    // Thresholds are tuned to the currently committed `es-ES.tsv`:
    // - A percentage-based threshold to catch "almost everything became English again".
    // - An absolute minimum so the test stays stable if the function catalog grows substantially.
    let tsv = include_str!("../src/locale/data/es-ES.tsv");
    let mut total = 0usize;
    let mut identity = 0usize;

    for line in tsv.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() || trimmed.starts_with('#') {
            continue;
        }

        // Parse the raw line (not the trimmed line) so trailing empty columns like
        // `SUM\tSUMA\t` are not silently accepted.
        let mut parts = line.split('\t');
        let canon = parts.next().unwrap_or("");
        let loc = parts.next().unwrap_or_else(|| {
            panic!("invalid TSV line in es-ES.tsv (expected `Canonical<TAB>Localized`): {line:?}")
        });
        if parts.next().is_some() {
            panic!("invalid TSV line in es-ES.tsv (too many columns): {line:?}");
        }
        total += 1;
        let canon = canon.trim();
        let loc = loc.trim();
        if canon.is_empty() || loc.is_empty() {
            panic!("invalid TSV line in es-ES.tsv (empty field): {line:?}");
        }
        if canon == loc {
            identity += 1;
        }
    }

    let non_identity = total - identity;
    let passes_ratio = non_identity * 100 >= total * 65;
    let passes_absolute = non_identity >= 340;
    assert!(
        passes_ratio || passes_absolute,
        "expected es-ES.tsv to contain many localized function spellings; got {non_identity}/{total} non-identity entries (identity={identity})"
    );
}

#[test]
fn locale_function_tsv_completeness_error_tsvs_are_complete_and_unique() {
    let expected: BTreeSet<String> = [
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
    .map(|k| k.as_code().to_string())
    .collect();

    let upstream_dir = upstream_error_data_dir();
    let upstream_tsvs = discover_tsvs_in_dir(&upstream_dir);

    for required in ["de-DE", "fr-FR", "es-ES"] {
        assert!(
            upstream_tsvs.contains_key(required),
            "expected upstream locale error TSV {required}.tsv to exist in {} (discovered: {:?})",
            upstream_dir.display(),
            upstream_tsvs.keys().collect::<Vec<_>>()
        );
    }

    let mut failures = String::new();

    for (locale_id, upstream_path) in &upstream_tsvs {
        let generated_path = locale_data_dir().join(format!("{locale_id}.errors.tsv"));

        let upstream_tsv = std::fs::read_to_string(upstream_path).unwrap_or_else(|err| {
            panic!(
                "failed to read upstream locale error TSV {}: {err}",
                upstream_path.display()
            )
        });
        let generated_tsv = std::fs::read_to_string(&generated_path).unwrap_or_else(|err| {
            panic!(
                "failed to read generated locale error TSV {}: {err}",
                generated_path.display()
            )
        });

        let parsed = parse_error_tsv(
            locale_id,
            &generated_path,
            &generated_tsv,
            /*require_sorted*/ true,
        );
        let upstream = parse_error_tsv(
            locale_id,
            upstream_path,
            &upstream_tsv,
            /*require_sorted*/ false,
        );

        let missing: Vec<String> = expected
            .difference(&parsed.canonical_keys)
            .cloned()
            .collect();
        let extra: Vec<String> = parsed
            .canonical_keys
            .difference(&expected)
            .cloned()
            .collect();

        let upstream_missing: Vec<String> = expected
            .difference(&upstream.canonical_keys)
            .cloned()
            .collect();
        let upstream_extra: Vec<String> = upstream
            .canonical_keys
            .difference(&expected)
            .cloned()
            .collect();

        let mapping_matches_upstream = parsed.entries == upstream.entries;
        let mut runtime_mismatches: Vec<String> = Vec::new();
        let locale = match get_locale(locale_id) {
            Some(locale) => Some(locale),
            None => {
                runtime_mismatches.push(format!(
                    "upstream error TSV exists for locale {locale_id}, but locale is not registered in formula_engine::locale::get_locale"
                ));
                None
            }
        };

        if let Some(locale) = locale {
            for (canon, localized_list) in &parsed.entries {
                let preferred = localized_list.first().unwrap_or_else(|| {
                    panic!(
                        "expected {locale_id} error TSV entry for {canon:?} to contain at least one localized spelling"
                    )
                });
                if locale.localized_error_literal(canon) != Some(preferred.as_str()) {
                    runtime_mismatches.push(format!(
                        "canonical {canon:?}: preferred localized mismatch (tsv preferred={preferred:?}, runtime={:?})",
                        locale.localized_error_literal(canon)
                    ));
                }
                for localized in localized_list {
                    if locale.canonical_error_literal(localized) != Some(canon.as_str()) {
                        runtime_mismatches.push(format!(
                            "localized {localized:?}: canonical mismatch (expected {canon:?}, runtime={:?})",
                            locale.canonical_error_literal(localized)
                        ));
                    }
                }
            }
        }

        let runtime_ok = runtime_mismatches.is_empty();

        if missing.is_empty()
            && extra.is_empty()
            && upstream_missing.is_empty()
            && upstream_extra.is_empty()
            && mapping_matches_upstream
            && runtime_ok
        {
            continue;
        }

        failures.push_str(&format!(
            "\nLocale error TSV {locale_id} is out of sync with formula-engine ErrorKind\n"
        ));
        failures.push_str(&format!(
            "  Generated TSV path: {}\n",
            generated_path.display()
        ));
        failures.push_str(&format!(
            "  Upstream TSV path: {}\n",
            upstream_path.display()
        ));

        if !missing.is_empty() {
            failures.push_str("\n  Missing canonical error keys:\n");
            for code in &missing {
                failures.push_str(&format!("    - {code}\n"));
            }
        }

        if !extra.is_empty() {
            failures.push_str("\n  Extra canonical error keys (not present in ErrorKind):\n");
            for code in &extra {
                failures.push_str(&format!("    - {code}\n"));
            }
        }

        if !upstream_missing.is_empty() || !upstream_extra.is_empty() {
            failures.push_str(&format!(
                "\n  Upstream mapping is out of sync ({}):\n",
                upstream_path.display()
            ));
            if !upstream_missing.is_empty() {
                failures.push_str("    Missing canonical error keys:\n");
                for code in &upstream_missing {
                    failures.push_str(&format!("      - {code}\n"));
                }
            }
            if !upstream_extra.is_empty() {
                failures.push_str("    Extra canonical error keys:\n");
                for code in &upstream_extra {
                    failures.push_str(&format!("      - {code}\n"));
                }
            }
        }

        if !mapping_matches_upstream {
            failures
                .push_str("\n  Error TSV does not match upstream mapping source. Regenerate it.\n");
        }
        if !runtime_ok {
            failures.push_str("\n  Runtime locale error translation mismatch:\n");
            for entry in &runtime_mismatches {
                failures.push_str(&format!("    - {entry}\n"));
            }
        }

        failures.push_str(&format!(
            "\n  Parsed entries: canonical={} localized={}\n",
            parsed.canonical_keys.len(),
            parsed.localized_keys.len()
        ));
    }

    if !failures.is_empty() {
        let mut report =
            String::from("Locale error TSVs are out of sync with the engine's error registry.\n");
        report.push_str(
            "\nWhen adding new ErrorKind variants, update each locale's upstream error-literal mapping and regenerate TSVs.\n",
        );
        report.push_str("\nRegenerate them with:\n  node scripts/generate-locale-error-tsvs.mjs\n");
        report.push_str(&failures);
        panic!("{report}");
    }
}

fn assert_error_alias(entries: &BTreeMap<String, Vec<String>>, canon: &str, localized: &str) {
    let Some(list) = entries.get(canon) else {
        panic!("expected error TSV to contain canonical key {canon:?}");
    };
    assert!(
        list.iter().any(|value| value == localized),
        "expected error TSV mapping for {canon:?} to include alias {localized:?}; got {list:?}"
    );
}

fn assert_error_preferred(entries: &BTreeMap<String, Vec<String>>, canon: &str, preferred: &str) {
    let Some(list) = entries.get(canon) else {
        panic!("expected error TSV to contain canonical key {canon:?}");
    };
    let Some(first) = list.first() else {
        panic!("expected error TSV mapping for {canon:?} to have at least one localized spelling");
    };
    assert_eq!(
        first, preferred,
        "expected preferred localized spelling for {canon:?} to be {preferred:?}; got {list:?}"
    );
}

#[test]
fn locale_error_tsvs_preserve_known_alias_spellings() {
    // Regression guard: some locales deliberately include additional spellings ("aliases") for
    // Excel compatibility (e.g. Spanish inverted punctuation variants, French synonyms for
    // dynamic-array spill errors). These aliases are needed for robust localized parsing and are
    // easy to lose if upstream extraction is overwritten without preserving existing entries.
    let path = Path::new("<in-memory>");

    let de_de = parse_error_tsv(
        "de-DE",
        path,
        include_str!("../src/locale/data/de-DE.errors.tsv"),
        /*require_sorted*/ true,
    );
    assert_error_preferred(&de_de.entries, "#CALC!", "#KALK!");
    assert_error_preferred(&de_de.entries, "#GETTING_DATA", "#DATEN_ABRUFEN");
    assert_error_alias(&de_de.entries, "#N/A", "#NV");
    assert_error_alias(&de_de.entries, "#N/A", "#N/A!");
    assert_error_preferred(&de_de.entries, "#N/A", "#NV");
    assert_error_preferred(&de_de.entries, "#NUM!", "#ZAHL!");
    assert_error_preferred(&de_de.entries, "#REF!", "#BEZUG!");
    assert_error_preferred(&de_de.entries, "#SPILL!", "#ÜBERLAUF!");
    assert_error_preferred(&de_de.entries, "#VALUE!", "#WERT!");

    let de_de_locale = get_locale("de-DE").expect("expected de-DE locale to be registered");
    assert_eq!(
        de_de_locale.localized_error_literal("#CALC!"),
        Some("#KALK!")
    );
    assert_eq!(
        de_de_locale.localized_error_literal("#GETTING_DATA"),
        Some("#DATEN_ABRUFEN")
    );
    assert_eq!(de_de_locale.localized_error_literal("#N/A"), Some("#NV"));
    assert_eq!(de_de_locale.canonical_error_literal("#NV"), Some("#N/A"));
    assert_eq!(de_de_locale.canonical_error_literal("#N/A!"), Some("#N/A"));
    assert_eq!(
        de_de_locale.localized_error_literal("#NUM!"),
        Some("#ZAHL!")
    );
    assert_eq!(
        de_de_locale.localized_error_literal("#REF!"),
        Some("#BEZUG!")
    );
    assert_eq!(
        de_de_locale.localized_error_literal("#SPILL!"),
        Some("#ÜBERLAUF!")
    );
    assert_eq!(
        de_de_locale.localized_error_literal("#VALUE!"),
        Some("#WERT!")
    );

    let fr_fr = parse_error_tsv(
        "fr-FR",
        path,
        include_str!("../src/locale/data/fr-FR.errors.tsv"),
        /*require_sorted*/ true,
    );
    assert_error_alias(&fr_fr.entries, "#NULL!", "#NUL!");
    assert_error_alias(&fr_fr.entries, "#NULL!", "#NULL!");
    assert_error_alias(&fr_fr.entries, "#SPILL!", "#PROPAGATION!");
    assert_error_alias(&fr_fr.entries, "#SPILL!", "#DEVERSEMENT!");
    assert_error_preferred(&fr_fr.entries, "#BLOCKED!", "#BLOQUE!");
    assert_error_preferred(&fr_fr.entries, "#CALC!", "#CALCUL!");
    assert_error_preferred(&fr_fr.entries, "#CONNECT!", "#CONNEXION!");
    assert_error_preferred(&fr_fr.entries, "#FIELD!", "#CHAMP!");
    assert_error_preferred(&fr_fr.entries, "#GETTING_DATA", "#OBTENTION_DONNEES");
    assert_error_preferred(&fr_fr.entries, "#NAME?", "#NOM?");
    assert_error_preferred(&fr_fr.entries, "#NULL!", "#NUL!");
    assert_error_preferred(&fr_fr.entries, "#SPILL!", "#PROPAGATION!");
    assert_error_preferred(&fr_fr.entries, "#NUM!", "#NOMBRE!");
    assert_error_preferred(&fr_fr.entries, "#UNKNOWN!", "#INCONNU!");
    assert_error_preferred(&fr_fr.entries, "#VALUE!", "#VALEUR!");

    let fr_fr_locale = get_locale("fr-FR").expect("expected fr-FR locale to be registered");
    assert_eq!(
        fr_fr_locale.localized_error_literal("#BLOCKED!"),
        Some("#BLOQUE!")
    );
    assert_eq!(
        fr_fr_locale.localized_error_literal("#CALC!"),
        Some("#CALCUL!")
    );
    assert_eq!(
        fr_fr_locale.localized_error_literal("#CONNECT!"),
        Some("#CONNEXION!")
    );
    assert_eq!(
        fr_fr_locale.localized_error_literal("#FIELD!"),
        Some("#CHAMP!")
    );
    assert_eq!(
        fr_fr_locale.localized_error_literal("#GETTING_DATA"),
        Some("#OBTENTION_DONNEES")
    );
    assert_eq!(
        fr_fr_locale.localized_error_literal("#NAME?"),
        Some("#NOM?")
    );
    assert_eq!(
        fr_fr_locale.localized_error_literal("#NULL!"),
        Some("#NUL!")
    );
    assert_eq!(
        fr_fr_locale.canonical_error_literal("#NUL!"),
        Some("#NULL!")
    );
    assert_eq!(
        fr_fr_locale.localized_error_literal("#SPILL!"),
        Some("#PROPAGATION!")
    );
    assert_eq!(
        fr_fr_locale.canonical_error_literal("#PROPAGATION!"),
        Some("#SPILL!")
    );
    assert_eq!(
        fr_fr_locale.canonical_error_literal("#DEVERSEMENT!"),
        Some("#SPILL!")
    );
    assert_eq!(
        fr_fr_locale.localized_error_literal("#NUM!"),
        Some("#NOMBRE!")
    );
    assert_eq!(
        fr_fr_locale.localized_error_literal("#UNKNOWN!"),
        Some("#INCONNU!")
    );
    assert_eq!(
        fr_fr_locale.localized_error_literal("#VALUE!"),
        Some("#VALEUR!")
    );

    let es_es = parse_error_tsv(
        "es-ES",
        path,
        include_str!("../src/locale/data/es-ES.errors.tsv"),
        /*require_sorted*/ true,
    );
    assert_error_preferred(&es_es.entries, "#BLOCKED!", "#¡BLOQUEADO!");
    assert_error_alias(&es_es.entries, "#BLOCKED!", "#BLOQUEADO!");
    assert_error_preferred(&es_es.entries, "#CALC!", "#¡CALC!");
    assert_error_alias(&es_es.entries, "#CALC!", "#CALC!");
    assert_error_preferred(&es_es.entries, "#CONNECT!", "#¡CONECTAR!");
    assert_error_alias(&es_es.entries, "#CONNECT!", "#CONECTAR!");
    assert_error_preferred(&es_es.entries, "#DIV/0!", "#¡DIV/0!");
    assert_error_alias(&es_es.entries, "#DIV/0!", "#DIV/0!");
    assert_error_preferred(&es_es.entries, "#FIELD!", "#¡CAMPO!");
    assert_error_alias(&es_es.entries, "#FIELD!", "#CAMPO!");
    assert_error_preferred(&es_es.entries, "#GETTING_DATA", "#OBTENIENDO_DATOS");
    assert_error_preferred(&es_es.entries, "#NULL!", "#¡NULO!");
    assert_error_alias(&es_es.entries, "#NULL!", "#NULO!");
    assert_error_preferred(&es_es.entries, "#NUM!", "#¡NUM!");
    assert_error_alias(&es_es.entries, "#NUM!", "#NUM!");
    assert_error_preferred(&es_es.entries, "#REF!", "#¡REF!");
    assert_error_alias(&es_es.entries, "#REF!", "#REF!");
    assert_error_alias(&es_es.entries, "#VALUE!", "#¡VALOR!");
    assert_error_alias(&es_es.entries, "#VALUE!", "#VALOR!");
    assert_error_alias(&es_es.entries, "#NAME?", "#¿NOMBRE?");
    assert_error_alias(&es_es.entries, "#NAME?", "#NOMBRE?");
    assert_error_alias(&es_es.entries, "#SPILL!", "#¡DESBORDAMIENTO!");
    assert_error_alias(&es_es.entries, "#SPILL!", "#DESBORDAMIENTO!");
    assert_error_preferred(&es_es.entries, "#VALUE!", "#¡VALOR!");
    assert_error_preferred(&es_es.entries, "#NAME?", "#¿NOMBRE?");
    assert_error_preferred(&es_es.entries, "#SPILL!", "#¡DESBORDAMIENTO!");
    assert_error_preferred(&es_es.entries, "#UNKNOWN!", "#¡DESCONOCIDO!");
    assert_error_alias(&es_es.entries, "#UNKNOWN!", "#DESCONOCIDO!");

    let es_es_locale = get_locale("es-ES").expect("expected es-ES locale to be registered");
    assert_eq!(
        es_es_locale.localized_error_literal("#VALUE!"),
        Some("#¡VALOR!")
    );
    assert_eq!(
        es_es_locale.localized_error_literal("#GETTING_DATA"),
        Some("#OBTENIENDO_DATOS")
    );
    assert_eq!(
        es_es_locale.canonical_error_literal("#¡VALOR!"),
        Some("#VALUE!")
    );
    assert_eq!(
        es_es_locale.canonical_error_literal("#VALOR!"),
        Some("#VALUE!")
    );
    assert_eq!(
        es_es_locale.canonical_error_literal("#OBTENIENDO_DATOS"),
        Some("#GETTING_DATA")
    );
    assert_eq!(
        es_es_locale.localized_error_literal("#NAME?"),
        Some("#¿NOMBRE?")
    );
    assert_eq!(
        es_es_locale.canonical_error_literal("#¿NOMBRE?"),
        Some("#NAME?")
    );
    assert_eq!(
        es_es_locale.canonical_error_literal("#NOMBRE?"),
        Some("#NAME?")
    );
    assert_eq!(
        es_es_locale.localized_error_literal("#SPILL!"),
        Some("#¡DESBORDAMIENTO!")
    );
    assert_eq!(
        es_es_locale.canonical_error_literal("#¡DESBORDAMIENTO!"),
        Some("#SPILL!")
    );
    assert_eq!(
        es_es_locale.canonical_error_literal("#DESBORDAMIENTO!"),
        Some("#SPILL!")
    );
}
