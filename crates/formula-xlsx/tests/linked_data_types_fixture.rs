use std::collections::BTreeSet;
use std::io::{Cursor, Read};
use std::path::Path;

use zip::ZipArchive;

fn normalize_opc_path(path: &str) -> String {
    let normalized = path.replace('\\', "/");
    let mut out: Vec<&str> = Vec::new();
    for segment in normalized.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                out.pop();
            }
            _ => out.push(segment),
        }
    }
    out.join("/")
}

fn resolve_relationship_target(base_part: &str, target: &str) -> String {
    let target = target.replace('\\', "/");
    if let Some(rest) = target.strip_prefix('/') {
        return normalize_opc_path(rest);
    }

    let base_dir = base_part
        .rsplit_once('/')
        .map(|(dir, _)| format!("{dir}/"))
        .unwrap_or_default();
    normalize_opc_path(&format!("{base_dir}{target}"))
}

fn zip_entry_names(zip_bytes: &[u8]) -> Result<Vec<String>, zip::result::ZipError> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor)?;
    let mut out = Vec::with_capacity(archive.len());
    for i in 0..archive.len() {
        let file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }
        out.push(file.name().to_string());
    }
    Ok(out)
}

fn zip_part_bytes(zip_bytes: &[u8], name: &str) -> Result<Vec<u8>, zip::result::ZipError> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor)?;
    let mut file = archive.by_name(name)?;
    let mut buf = Vec::with_capacity(file.size() as usize);
    file.read_to_end(&mut buf)?;
    Ok(buf)
}

#[test]
fn linked_data_types_fixture_roundtrips_richdata_parts() -> Result<(), Box<dyn std::error::Error>> {
    let fixture_path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/richdata/linked-data-types.xlsx");
    let fixture_bytes = std::fs::read(&fixture_path)?;

    let names = zip_entry_names(&fixture_bytes)?;
    let name_set: BTreeSet<&str> = names.iter().map(|s| s.as_str()).collect();

    assert!(
        name_set.contains("xl/metadata.xml"),
        "fixture missing xl/metadata.xml; zip entries: {names:?}"
    );
    assert!(
        name_set.contains("xl/_rels/metadata.xml.rels"),
        "fixture missing xl/_rels/metadata.xml.rels; zip entries: {names:?}"
    );

    let richdata_parts: Vec<String> = names
        .iter()
        .filter(|n| n.starts_with("xl/richData/"))
        .cloned()
        .collect();
    assert!(
        !richdata_parts.is_empty(),
        "fixture missing xl/richData/* parts; zip entries: {names:?}"
    );

    // Ensure metadata.xml.rels actually links to richData parts (don't hard-code names).
    let metadata_rels_bytes = zip_part_bytes(&fixture_bytes, "xl/_rels/metadata.xml.rels")?;
    let metadata_rels = std::str::from_utf8(&metadata_rels_bytes)?;
    let rels_doc = roxmltree::Document::parse(metadata_rels)?;
    let mut richdata_rel_targets: Vec<String> = Vec::new();
    for rel in rels_doc.descendants().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "Relationship"
            && n.attribute("Target").is_some()
    }) {
        let target = rel.attribute("Target").unwrap_or_default();
        let resolved = resolve_relationship_target("xl/metadata.xml", target);
        if resolved.starts_with("xl/richData/") {
            richdata_rel_targets.push(resolved);
        }
    }
    assert!(
        !richdata_rel_targets.is_empty(),
        "expected xl/_rels/metadata.xml.rels to reference at least one xl/richData/* target; rels: {metadata_rels}"
    );
    for target in &richdata_rel_targets {
        assert!(
            name_set.contains(target.as_str()),
            "metadata.xml.rels references missing target {target}; zip entries: {names:?}"
        );
    }

    // Parse the worksheet XML and ensure at least one cell has vm/cm attributes.
    let sheet_xml_bytes = zip_part_bytes(&fixture_bytes, "xl/worksheets/sheet1.xml")?;
    let sheet_xml = std::str::from_utf8(&sheet_xml_bytes)?;
    let parsed = roxmltree::Document::parse(sheet_xml)?;
    let has_vm_or_cm = parsed.descendants().any(|n| {
        n.is_element()
            && n.tag_name().name() == "c"
            && (n.attribute("vm").is_some() || n.attribute("cm").is_some())
    });
    assert!(
        has_vm_or_cm,
        "expected at least one <c> element with vm/cm attributes; sheet1.xml: {sheet_xml}"
    );

    // Try resolving vm -> rich value index via metadata.xml (best-effort).
    //
    // Note: Excel and other producers have been observed to use multiple metadata schemas
    // for rich values; this mapping can legitimately be empty. The core requirement for
    // this fixture/test is part + attribute preservation, not full semantic decoding.
    let metadata_xml_bytes = zip_part_bytes(&fixture_bytes, "xl/metadata.xml")?;
    let vm_map = formula_xlsx::parse_value_metadata_vm_to_rich_value_index_map(&metadata_xml_bytes)?;

    // Ensure any mapped rich value indices are in-range for xl/richData/richValue.xml.
    if !vm_map.is_empty() && name_set.contains("xl/richData/richValue.xml") {
        let rich_value_bytes = zip_part_bytes(&fixture_bytes, "xl/richData/richValue.xml")?;
        let rich_value_xml = std::str::from_utf8(&rich_value_bytes)?;
        let rich_value_doc = roxmltree::Document::parse(rich_value_xml)?;
        let rv_count = rich_value_doc
            .descendants()
            .filter(|n| n.is_element() && n.tag_name().name() == "rv")
            .count() as u32;
        assert!(rv_count > 0, "expected richValue.xml to contain <rv> records");
        for (vm, rv_idx) in &vm_map {
            assert!(
                *rv_idx < rv_count,
                "vm={vm} maps to rv index {rv_idx}, but richValue.xml has {rv_count} <rv> records"
            );
        }
    }

    // Round-trip through XlsxDocument and ensure richdata-related parts remain present.
    let doc = formula_xlsx::load_from_bytes(&fixture_bytes)?;
    let saved = doc.save_to_vec()?;

    let saved_names = zip_entry_names(&saved)?;
    let saved_set: BTreeSet<&str> = saved_names.iter().map(|s| s.as_str()).collect();

    let mut expected_parts: BTreeSet<String> = BTreeSet::new();
    expected_parts.insert("xl/metadata.xml".to_string());
    expected_parts.insert("xl/_rels/metadata.xml.rels".to_string());
    expected_parts.extend(richdata_parts.iter().cloned());

    for part in expected_parts {
        assert!(
            saved_set.contains(part.as_str()),
            "round-tripped workbook missing {part}; richData parts in fixture: {richdata_parts:?}; saved zip entries: {saved_names:?}"
        );
    }

    // Optional: diff the full workbooks and ensure there are no critical diffs.
    // This provides a stronger, real-world baseline than part-name existence checks alone.
    let tmpdir = tempfile::tempdir()?;
    let out_path = tmpdir.path().join("roundtripped.xlsx");
    std::fs::write(&out_path, &saved)?;

    let report = xlsx_diff::diff_workbooks(&fixture_path, &out_path)?;
    if report.has_at_least(xlsx_diff::Severity::Critical) {
        eprintln!(
            "Critical diffs detected for linked data types fixture {}",
            fixture_path.display()
        );
        for diff in report
            .differences
            .iter()
            .filter(|d| d.severity == xlsx_diff::Severity::Critical)
        {
            eprintln!("{diff}");
        }
        panic!("fixture did not round-trip cleanly via XlsxDocument");
    }

    Ok(())
}
