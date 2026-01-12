use std::collections::{BTreeSet, HashMap};
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
    // Relationship targets are URIs; internal targets may include a fragment (e.g. `foo.xml#bar`).
    // OPC part names do not include fragments, so strip them before resolving.
    let target = target.split_once('#').map(|(t, _)| t).unwrap_or(&target);
    if target.is_empty() {
        // A target of just `#fragment` refers to the source part itself.
        return normalize_opc_path(base_part.trim_start_matches('/'));
    }
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

fn parse_content_types_overrides(
    content_types_xml: &str,
) -> Result<HashMap<String, String>, roxmltree::Error> {
    let doc = roxmltree::Document::parse(content_types_xml)?;
    let mut out = HashMap::new();
    for node in doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Override")
    {
        let Some(part_name) = node.attribute("PartName") else {
            continue;
        };
        let Some(content_type) = node.attribute("ContentType") else {
            continue;
        };
        out.insert(part_name.to_string(), content_type.to_string());
    }
    Ok(out)
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

    // Validate `[Content_Types].xml` includes overrides for the rich-data parts and metadata.xml.
    // This is a useful baseline for a real Excel 365 fixture (Excel typically emits explicit
    // overrides for these custom rich value content types).
    let content_types_bytes = zip_part_bytes(&fixture_bytes, "[Content_Types].xml")?;
    let content_types_xml = std::str::from_utf8(&content_types_bytes)?;
    let overrides = parse_content_types_overrides(content_types_xml)?;
    assert_eq!(
        overrides
            .get("/xl/metadata.xml")
            .map(|s| s.as_str()),
        Some("application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"),
        "expected /xl/metadata.xml content type override; [Content_Types].xml: {content_types_xml}"
    );
    for part in &richdata_parts {
        // RichData may also include relationship parts like:
        // `xl/richData/_rels/richValueRel.xml.rels`
        // Those are covered by the `[Content_Types].xml` default `rels` mapping rather than an
        // explicit override, so only validate overrides for the richData XML payload parts.
        if part.starts_with("xl/richData/_rels/") || part.ends_with(".rels") {
            continue;
        }
        if !part.ends_with(".xml") {
            continue;
        }

        let ct_part_name = format!("/{part}");
        let Some(content_type) = overrides.get(&ct_part_name) else {
            panic!(
                "missing content type override for {ct_part_name}; [Content_Types].xml: {content_types_xml}"
            );
        };
        assert!(
            content_type.starts_with("application/vnd.ms-excel."),
            "expected {ct_part_name} content type to be an Excel richData type, got {content_type}"
        );
    }

    // Ensure metadata.xml.rels actually links to richData parts (don't hard-code names).
    let metadata_rels_bytes = zip_part_bytes(&fixture_bytes, "xl/_rels/metadata.xml.rels")?;
    let metadata_rels = std::str::from_utf8(&metadata_rels_bytes)?;
    let rels_doc = roxmltree::Document::parse(metadata_rels)?;
    let mut richdata_rel_targets: Vec<String> = Vec::new();
    let mut richdata_rel_type_uris: BTreeSet<String> = BTreeSet::new();
    for rel in rels_doc.descendants().filter(|n| {
        n.is_element()
            && n.tag_name().name() == "Relationship"
            && n.attribute("Target").is_some()
    }) {
        let target = rel.attribute("Target").unwrap_or_default();
        let resolved = resolve_relationship_target("xl/metadata.xml", target);
        if resolved.starts_with("xl/richData/") {
            richdata_rel_targets.push(resolved);
            if let Some(type_uri) = rel.attribute("Type") {
                richdata_rel_type_uris.insert(type_uri.to_string());
            }
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

    // Validate relationship type URIs (Excel-style Office 2017 rich value relationships).
    for expected in [
        "http://schemas.microsoft.com/office/2017/relationships/richValue",
        "http://schemas.microsoft.com/office/2017/relationships/richValueRel",
        "http://schemas.microsoft.com/office/2017/relationships/richValueTypes",
        "http://schemas.microsoft.com/office/2017/relationships/richValueStructure",
    ] {
        assert!(
            richdata_rel_type_uris.contains(expected),
            "expected metadata.xml.rels to contain relationship type URI {expected}; saw: {richdata_rel_type_uris:?}\nrels: {metadata_rels}"
        );
    }

    // Parse the worksheet XML and ensure at least one cell has vm/cm attributes.
    let sheet_xml_bytes = zip_part_bytes(&fixture_bytes, "xl/worksheets/sheet1.xml")?;
    let sheet_xml = std::str::from_utf8(&sheet_xml_bytes)?;
    let parsed = roxmltree::Document::parse(sheet_xml)?;
    let mut vm_values: Vec<u32> = Vec::new();
    let mut cm_values: Vec<u32> = Vec::new();
    for c in parsed
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "c")
    {
        if let Some(vm) = c.attribute("vm").and_then(|v| v.trim().parse::<u32>().ok()) {
            vm_values.push(vm);
        }
        if let Some(cm) = c.attribute("cm").and_then(|v| v.trim().parse::<u32>().ok()) {
            cm_values.push(cm);
        }
    }
    let has_vm_or_cm = !vm_values.is_empty() || !cm_values.is_empty();
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

    // Confirm whether worksheet vm/cm indices are 1-based or 0-based.
    //
    // Observed in real Excel 365 richValue fixtures in this repo: 1-based (`vm="1"` is first record).
    // Some synthetic fixtures and non-Excel producers can emit 0-based indices.
    if !vm_values.is_empty() {
        let min_vm = *vm_values.iter().min().unwrap();
        assert!(
            min_vm >= 1,
            "expected worksheet vm indices to be 1-based (no vm=0) for this fixture; vm values: {vm_values:?}\n(sheet1.xml: {sheet_xml})"
        );
        // If the metadata mapping is available, ensure every worksheet vm can be resolved directly.
        if !vm_map.is_empty() {
            for vm in &vm_values {
                assert!(
                    vm_map.contains_key(vm),
                    "worksheet vm={vm} not found in metadata vm map keys; keys: {:?}",
                    vm_map.keys().collect::<BTreeSet<_>>()
                );
            }
        }
    }
    if !cm_values.is_empty() {
        let min_cm = *cm_values.iter().min().unwrap();
        assert!(
            min_cm >= 1,
            "expected worksheet cm indices to be 1-based (no cm=0) for this fixture; cm values: {cm_values:?}\n(sheet1.xml: {sheet_xml})"
        );
        // If the worksheet uses `cm`, ensure metadata.xml declares a matching `cellMetadata` table.
        let metadata_xml = std::str::from_utf8(&metadata_xml_bytes)?;
        let metadata_doc = roxmltree::Document::parse(metadata_xml)?;
        let cell_metadata = metadata_doc
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "cellMetadata");
        assert!(
            cell_metadata.is_some(),
            "sheet uses cm attributes but metadata.xml is missing <cellMetadata>; metadata.xml: {metadata_xml}"
        );
        if let Some(node) = cell_metadata {
            if let Some(count) = node.attribute("count").and_then(|c| c.trim().parse::<u32>().ok())
            {
                let max_cm = *cm_values.iter().max().unwrap();
                assert!(
                    max_cm <= count,
                    "max cm={max_cm} exceeds metadata cellMetadata count={count}"
                );
            }
        }
    }

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
