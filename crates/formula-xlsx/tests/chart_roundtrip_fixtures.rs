use std::collections::{BTreeMap, BTreeSet};
use std::error::Error;
use std::path::{Path, PathBuf};

use formula_xlsx::load_from_bytes;
use tempfile::tempdir;

const CHART_PART_PREFIXES: &[&str] = &["xl/charts/", "xl/drawings/", "xl/media/"];

#[test]
fn chart_roundtrip_fixtures_preserve_chart_opc_parts() -> Result<(), Box<dyn Error>> {
    let fixtures_root = chart_fixtures_root();
    let mut fixtures = xlsx_diff::collect_fixture_paths(&fixtures_root)?;
    fixtures.retain(|path| path.extension().and_then(|ext| ext.to_str()) == Some("xlsx"));

    assert!(
        !fixtures.is_empty(),
        "no .xlsx fixtures found under {}",
        fixtures_root.display()
    );

    for fixture in fixtures {
        assert_fixture_roundtrip_preserves_charts(&fixture)?;
    }

    Ok(())
}

fn chart_fixtures_root() -> PathBuf {
    let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
    let preferred = manifest_dir.join("../../fixtures/charts/xlsx");
    if preferred.exists() {
        return preferred;
    }
    manifest_dir.join("../../fixtures/xlsx/charts")
}

fn assert_fixture_roundtrip_preserves_charts(fixture: &Path) -> Result<(), Box<dyn Error>> {
    let original_bytes = std::fs::read(fixture)?;
    let doc = load_from_bytes(&original_bytes)?;
    let roundtripped_bytes = doc.save_to_vec()?;

    let tmpdir = tempdir()?;
    let roundtripped_path = tmpdir.path().join("roundtripped.xlsx");
    std::fs::write(&roundtripped_path, &roundtripped_bytes)?;

    let report = xlsx_diff::diff_workbooks(fixture, &roundtripped_path)?;
    let expected = xlsx_diff::WorkbookArchive::open(fixture)?;
    let actual = xlsx_diff::WorkbookArchive::open(&roundtripped_path)?;

    let expected_parts: BTreeSet<String> = expected.part_names().into_iter().map(str::to_string).collect();
    let actual_parts: BTreeSet<String> = actual.part_names().into_iter().map(str::to_string).collect();

    let mut issues: BTreeMap<String, Vec<String>> = BTreeMap::new();

    for part in expected_parts.difference(&actual_parts) {
        if part.starts_with("xl/charts/") || part.starts_with("xl/drawings/") {
            issues
                .entry(part.to_string())
                .or_default()
                .push("missing_part".to_string());
        }
    }

    for part in expected_parts.iter() {
        if !part.ends_with(".rels") {
            continue;
        }

        let expected_bytes = expected
            .get(part)
            .ok_or_else(|| format!("expected archive missing part {part}"))?;
        match actual.get(part) {
            Some(actual_bytes) if actual_bytes == expected_bytes => {}
            Some(actual_bytes) => issues
                .entry(part.to_string())
                .or_default()
                .push(format!(
                    "rels_bytes_changed (expected {}, actual {}, first_diff {})",
                    expected_bytes.len(),
                    actual_bytes.len(),
                    first_diff_offset(expected_bytes, actual_bytes)
                        .map(|idx| idx.to_string())
                        .unwrap_or_else(|| "unknown".to_string())
                )),
            None => issues
                .entry(part.to_string())
                .or_default()
                .push("rels_missing".to_string()),
        }
    }

    for part in actual_parts.difference(&expected_parts) {
        if part.ends_with(".rels") {
            issues
                .entry(part.to_string())
                .or_default()
                .push("extra_rels_part".to_string());
        }
    }

    for part in expected_parts.iter() {
        if !is_chart_related_part(part) {
            continue;
        }

        let expected_bytes = expected
            .get(part)
            .ok_or_else(|| format!("expected archive missing part {part}"))?;
        match actual.get(part) {
            Some(actual_bytes) if actual_bytes == expected_bytes => {}
            Some(actual_bytes) => {
                let mut msg = format!(
                    "chart_part_bytes_changed (expected {}, actual {}, first_diff {})",
                    expected_bytes.len(),
                    actual_bytes.len(),
                    first_diff_offset(expected_bytes, actual_bytes)
                        .map(|idx| idx.to_string())
                        .unwrap_or_else(|| "unknown".to_string())
                );
                if part.starts_with("xl/charts/") && part.ends_with(".xml") {
                    if let (Ok(expected_xml), Ok(actual_xml)) = (
                        xlsx_diff::NormalizedXml::parse(part, expected_bytes),
                        xlsx_diff::NormalizedXml::parse(part, actual_bytes),
                    ) {
                        if expected_xml == actual_xml {
                            msg.push_str(" (normalized_xml_equal)");
                        }
                    }
                }
                issues.entry(part.to_string()).or_default().push(msg);
            }
            None => {
                issues
                    .entry(part.to_string())
                    .or_default()
                    .push("chart_part_missing".to_string());
            }
        }
    }

    if !issues.is_empty() {
        eprintln!(
            "Chart OPC parts changed during no-op round-trip for fixture {}",
            fixture.display()
        );
        print_grouped_report(&issues, &report);
        panic!("fixture {} did not preserve chart parts", fixture.display());
    }

    Ok(())
}

fn is_chart_related_part(part: &str) -> bool {
    CHART_PART_PREFIXES
        .iter()
        .any(|prefix| part.starts_with(prefix))
}

fn first_diff_offset(a: &[u8], b: &[u8]) -> Option<usize> {
    let shared = a.len().min(b.len());
    for idx in 0..shared {
        if a[idx] != b[idx] {
            return Some(idx);
        }
    }
    (a.len() != b.len()).then_some(shared)
}

fn print_grouped_report(issues: &BTreeMap<String, Vec<String>>, report: &xlsx_diff::DiffReport) {
    let mut diffs_by_part: BTreeMap<&str, Vec<&xlsx_diff::Difference>> = BTreeMap::new();
    for diff in &report.differences {
        if is_chart_related_part(&diff.part) || diff.part.ends_with(".rels") {
            diffs_by_part.entry(&diff.part).or_default().push(diff);
        }
    }

    let mut parts: BTreeSet<&str> = BTreeSet::new();
    parts.extend(issues.keys().map(String::as_str));
    parts.extend(diffs_by_part.keys().copied());

    for part in parts {
        eprintln!("- {part}");
        if let Some(entries) = issues.get(part) {
            for entry in entries {
                eprintln!("  issue: {entry}");
            }
        }
        if let Some(diffs) = diffs_by_part.get(part) {
            for diff in diffs {
                eprintln!(
                    "  diff: [{}] {}{}{}",
                    diff.severity,
                    diff.kind,
                    if diff.path.is_empty() { "" } else { " " },
                    diff.path
                );
                if diff.expected.is_some() || diff.actual.is_some() {
                    eprintln!(
                        "    expected: {}",
                        diff.expected.as_deref().unwrap_or("<none>")
                    );
                    eprintln!(
                        "    actual:   {}",
                        diff.actual.as_deref().unwrap_or("<none>")
                    );
                }
            }
        }
    }
}

