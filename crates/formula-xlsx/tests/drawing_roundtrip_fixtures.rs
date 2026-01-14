use std::collections::{BTreeMap, BTreeSet};
use std::error::Error;

use formula_model::{CellRef, CellValue};
use formula_xlsx::openxml::{parse_relationships, resolve_target};
use formula_xlsx::{load_from_bytes, XlsxPackage};

const IMAGE_FIXTURE: &[u8] = include_bytes!("../../../fixtures/xlsx/basic/image.xlsx");
const SMARTART_FIXTURE: &[u8] = include_bytes!("../../../fixtures/xlsx/basic/smartart.xlsx");
const CHART_FIXTURE: &[u8] = include_bytes!("../../../fixtures/charts/xlsx/bar.xlsx");

#[test]
fn drawing_roundtrip_fixtures_preserve_drawing_opc_parts() -> Result<(), Box<dyn Error>> {
    for (name, bytes) in [
        ("fixtures/xlsx/basic/image.xlsx", IMAGE_FIXTURE),
        ("fixtures/xlsx/basic/smartart.xlsx", SMARTART_FIXTURE),
        ("fixtures/charts/xlsx/bar.xlsx", CHART_FIXTURE),
    ] {
        assert_fixture_roundtrip_preserves_drawings(name, bytes)?;
    }

    Ok(())
}

#[test]
fn drawing_roundtrip_fixtures_preserve_drawing_opc_parts_on_unrelated_cell_edit(
) -> Result<(), Box<dyn Error>> {
    for (name, bytes) in [
        ("fixtures/xlsx/basic/image.xlsx", IMAGE_FIXTURE),
        ("fixtures/xlsx/basic/smartart.xlsx", SMARTART_FIXTURE),
        ("fixtures/charts/xlsx/bar.xlsx", CHART_FIXTURE),
    ] {
        assert_fixture_cell_edit_preserves_drawings(name, bytes)?;
    }

    Ok(())
}

fn assert_fixture_roundtrip_preserves_drawings(
    fixture_name: &str,
    original_bytes: &[u8],
) -> Result<(), Box<dyn Error>> {
    let doc = load_from_bytes(original_bytes)?;
    let roundtripped_bytes = doc.save_to_vec()?;

    let expected = XlsxPackage::from_bytes(original_bytes)?;
    let actual = XlsxPackage::from_bytes(&roundtripped_bytes)?;

    let expected_drawing_xml_parts = collect_parts(&expected, is_drawing_xml_part);
    let actual_drawing_xml_parts = collect_parts(&actual, is_drawing_xml_part);

    let expected_drawing_rels_parts = collect_parts(&expected, is_drawing_rels_part);
    let actual_drawing_rels_parts = collect_parts(&actual, is_drawing_rels_part);

    let expected_sheet_rels_parts = collect_parts(&expected, is_worksheet_rels_part);
    let actual_sheet_rels_parts = collect_parts(&actual, is_worksheet_rels_part);

    let expected_media_parts = collect_media_parts_referenced_by_drawings(&expected, &expected_drawing_rels_parts)?;

    let mut issues: BTreeMap<String, Vec<String>> = BTreeMap::new();

    report_part_set_differences(
        &expected_drawing_xml_parts,
        &actual_drawing_xml_parts,
        "drawing_xml",
        &mut issues,
    );
    report_part_set_differences(
        &expected_drawing_rels_parts,
        &actual_drawing_rels_parts,
        "drawing_rels",
        &mut issues,
    );
    report_part_set_differences(
        &expected_sheet_rels_parts,
        &actual_sheet_rels_parts,
        "worksheet_rels",
        &mut issues,
    );

    // Compare bytes for all drawing parts and their drawing-level rels.
    for part in expected_drawing_xml_parts
        .iter()
        .chain(expected_drawing_rels_parts.iter())
    {
        compare_part_bytes(&expected, &actual, part, &mut issues);
    }

    // Compare bytes for worksheet `.rels` parts as well; the writer should not need to touch them
    // when round-tripping drawings unchanged.
    for part in expected_sheet_rels_parts.iter() {
        compare_part_bytes(&expected, &actual, part, &mut issues);
    }

    // Compare bytes for drawing-referenced media parts.
    for part in expected_media_parts.iter() {
        compare_part_bytes(&expected, &actual, part, &mut issues);
    }

    if !issues.is_empty() {
        eprintln!(
            "Drawing OPC parts changed during no-op round-trip for fixture {fixture_name}"
        );
        for (part, entries) in issues {
            eprintln!("- {part}");
            for entry in entries {
                eprintln!("  issue: {entry}");
            }
        }
        panic!("fixture {fixture_name} did not preserve drawing parts");
    }

    Ok(())
}

fn assert_fixture_cell_edit_preserves_drawings(
    fixture_name: &str,
    original_bytes: &[u8],
) -> Result<(), Box<dyn Error>> {
    let mut doc = load_from_bytes(original_bytes)?;
    let sheet_id = doc
        .workbook
        .sheets
        .first()
        .map(|sheet| sheet.id)
        .expect("fixture should contain at least one sheet");
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .set_value(
            CellRef::from_a1("A1").expect("valid A1"),
            // Use a string so the writer must also patch shared string tables/relationships, which
            // should still not cause drawing parts to be rewritten.
            CellValue::String("hello".to_string()),
        );

    let roundtripped_bytes = doc.save_to_vec()?;

    let expected = XlsxPackage::from_bytes(original_bytes)?;
    let actual = XlsxPackage::from_bytes(&roundtripped_bytes)?;

    let expected_drawing_xml_parts = collect_parts(&expected, is_drawing_xml_part);
    let actual_drawing_xml_parts = collect_parts(&actual, is_drawing_xml_part);

    let expected_drawing_rels_parts = collect_parts(&expected, is_drawing_rels_part);
    let actual_drawing_rels_parts = collect_parts(&actual, is_drawing_rels_part);

    let expected_sheet_rels_parts = collect_parts(&expected, is_worksheet_rels_part);
    let actual_sheet_rels_parts = collect_parts(&actual, is_worksheet_rels_part);

    let expected_media_parts =
        collect_media_parts_referenced_by_drawings(&expected, &expected_drawing_rels_parts)?;

    let mut issues: BTreeMap<String, Vec<String>> = BTreeMap::new();

    report_part_set_differences(
        &expected_drawing_xml_parts,
        &actual_drawing_xml_parts,
        "drawing_xml",
        &mut issues,
    );
    report_part_set_differences(
        &expected_drawing_rels_parts,
        &actual_drawing_rels_parts,
        "drawing_rels",
        &mut issues,
    );
    report_part_set_differences(
        &expected_sheet_rels_parts,
        &actual_sheet_rels_parts,
        "worksheet_rels",
        &mut issues,
    );

    // Compare bytes for all drawing parts and their drawing-level rels.
    for part in expected_drawing_xml_parts
        .iter()
        .chain(expected_drawing_rels_parts.iter())
    {
        compare_part_bytes(&expected, &actual, part, &mut issues);
    }

    for part in expected_sheet_rels_parts.iter() {
        compare_part_bytes(&expected, &actual, part, &mut issues);
    }

    // Compare bytes for drawing-referenced media parts.
    for part in expected_media_parts.iter() {
        compare_part_bytes(&expected, &actual, part, &mut issues);
    }

    if !issues.is_empty() {
        eprintln!(
            "Drawing OPC parts changed during cell-edit round-trip for fixture {fixture_name}"
        );
        for (part, entries) in issues {
            eprintln!("- {part}");
            for entry in entries {
                eprintln!("  issue: {entry}");
            }
        }
        panic!("fixture {fixture_name} did not preserve drawing parts");
    }

    Ok(())
}

fn collect_parts(package: &XlsxPackage, predicate: fn(&str) -> bool) -> BTreeSet<String> {
    package
        .part_names()
        .map(canonical_part_name)
        .filter(|name| predicate(name))
        .map(str::to_string)
        .collect()
}

fn canonical_part_name(name: &str) -> &str {
    name.strip_prefix('/').unwrap_or(name)
}

fn is_drawing_xml_part(name: &str) -> bool {
    name.starts_with("xl/drawings/") && name.ends_with(".xml")
}

fn is_drawing_rels_part(name: &str) -> bool {
    name.starts_with("xl/drawings/_rels/") && name.ends_with(".rels")
}

fn is_worksheet_rels_part(name: &str) -> bool {
    name.starts_with("xl/worksheets/_rels/") && name.ends_with(".rels")
}

fn collect_media_parts_referenced_by_drawings(
    package: &XlsxPackage,
    drawing_rels_parts: &BTreeSet<String>,
) -> Result<BTreeSet<String>, Box<dyn Error>> {
    let mut out = BTreeSet::new();

    for rels_part in drawing_rels_parts {
        let Some(rels_bytes) = package.part(rels_part) else {
            continue;
        };

        let Some(source_part) = source_part_for_rels(rels_part) else {
            continue;
        };

        for rel in parse_relationships(rels_bytes)? {
            if rel
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
            {
                continue;
            }

            let resolved = resolve_target(&source_part, &rel.target);
            if resolved.starts_with("xl/media/") {
                out.insert(resolved);
            }
        }
    }

    Ok(out)
}

fn source_part_for_rels(rels_part: &str) -> Option<String> {
    let rels_part = canonical_part_name(rels_part);
    let suffix = ".rels";
    let stem = rels_part.strip_suffix(suffix)?;
    let prefix = "xl/drawings/_rels/";
    let file = stem.strip_prefix(prefix)?;
    Some(format!("xl/drawings/{file}"))
}

fn compare_part_bytes(
    expected: &XlsxPackage,
    actual: &XlsxPackage,
    part: &str,
    issues: &mut BTreeMap<String, Vec<String>>,
) {
    let part = canonical_part_name(part);
    let Some(expected_bytes) = expected.part(part) else {
        issues
            .entry(part.to_string())
            .or_default()
            .push("expected_missing_part".to_string());
        return;
    };

    match actual.part(part) {
        Some(actual_bytes) if actual_bytes == expected_bytes => {}
        Some(actual_bytes) => {
            issues.entry(part.to_string()).or_default().push(format!(
                "bytes_changed (expected {}, actual {}, first_diff {})",
                expected_bytes.len(),
                actual_bytes.len(),
                first_diff_offset(expected_bytes, actual_bytes)
                    .map(|idx| idx.to_string())
                    .unwrap_or_else(|| "unknown".to_string())
            ));
        }
        None => {
            issues
                .entry(part.to_string())
                .or_default()
                .push("missing_part".to_string());
        }
    }
}

fn report_part_set_differences(
    expected: &BTreeSet<String>,
    actual: &BTreeSet<String>,
    kind: &str,
    issues: &mut BTreeMap<String, Vec<String>>,
) {
    for part in expected.difference(actual) {
        issues
            .entry(part.clone())
            .or_default()
            .push(format!("{kind}_missing"));
    }
    for part in actual.difference(expected) {
        issues
            .entry(part.clone())
            .or_default()
            .push(format!("{kind}_extra"));
    }
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
