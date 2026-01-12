use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Read, Write};
use std::path::Path;

use formula_model::SheetVisibility;
use formula_xlsx::load_from_bytes;
use pretty_assertions::assert_eq;
use quick_xml::events::Event;
use quick_xml::Reader;
use zip::write::FileOptions;
use zip::ZipArchive;
use zip::ZipWriter;

fn fixture_bytes() -> Vec<u8> {
    std::fs::read(fixture_path()).expect("fixture exists")
}

fn fixture_path() -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../fixtures/xlsx/basic/multi-sheet.xlsx")
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn workbook_sheet_entries(xml: &[u8]) -> Vec<(String, u32, String)> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).expect("read xml") {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"sheet" => {
                let mut name = None;
                let mut sheet_id = None;
                let mut rid = None;
                for attr in e.attributes().flatten() {
                    let key = attr.key.as_ref();
                    let val = attr.unescape_value().expect("attr").into_owned();
                    match key {
                        b"name" => name = Some(val),
                        b"sheetId" => sheet_id = val.parse::<u32>().ok(),
                        b"r:id" => rid = Some(val),
                        _ => {}
                    }
                }
                out.push((
                    name.expect("name"),
                    sheet_id.expect("sheetId"),
                    rid.expect("r:id"),
                ));
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn workbook_sheet_states(xml: &[u8]) -> BTreeMap<String, Option<String>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = BTreeMap::new();
    loop {
        match reader.read_event_into(&mut buf).expect("read xml") {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"sheet" => {
                let mut name = None;
                let mut state = None;
                for attr in e.attributes().flatten() {
                    let key = attr.key.as_ref();
                    let val = attr.unescape_value().expect("attr").into_owned();
                    match key {
                        b"name" => name = Some(val),
                        b"state" => state = Some(val),
                        _ => {}
                    }
                }
                out.insert(name.expect("name"), state);
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn workbook_relationship_targets(xml: &[u8]) -> BTreeMap<String, String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut map = BTreeMap::new();
    loop {
        match reader.read_event_into(&mut buf).expect("read xml") {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                let mut id = None;
                let mut target = None;
                for attr in e.attributes().flatten() {
                    let val = attr.unescape_value().expect("attr").into_owned();
                    match attr.key.as_ref() {
                        b"Id" => id = Some(val),
                        b"Target" => target = Some(val),
                        _ => {}
                    }
                }
                if let (Some(id), Some(target)) = (id, target) {
                    map.insert(id, target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    map
}

fn assert_sheets_have_backing_parts(xlsx_bytes: &[u8]) {
    let sheets = workbook_sheet_entries(&zip_part(xlsx_bytes, "xl/workbook.xml"));
    let rels = workbook_relationship_targets(&zip_part(xlsx_bytes, "xl/_rels/workbook.xml.rels"));

    let cursor = Cursor::new(xlsx_bytes);
    let archive = ZipArchive::new(cursor).expect("open zip");
    let names: BTreeSet<String> = archive
        .file_names()
        .map(|s| s.to_string())
        .collect();

    for (_, _, rid) in sheets {
        let target = rels.get(&rid).expect("sheet r:id exists in workbook rels");
        let part_name = if target.starts_with('/') {
            target.trim_start_matches('/').to_string()
        } else {
            format!("xl/{target}")
        };
        assert!(
            names.contains(&part_name),
            "worksheet part {part_name} missing for {rid}"
        );
    }
}

fn diff_parts(expected: &Path, actual_bytes: &[u8]) -> BTreeSet<String> {
    let tmpdir = tempfile::tempdir().expect("tmpdir");
    let out_path = tmpdir.path().join("out.xlsx");
    std::fs::write(&out_path, actual_bytes).expect("write output");

    let report = xlsx_diff::diff_workbooks(expected, &out_path).expect("diff workbooks");
    report
        .differences
        .iter()
        .map(|d| d.part.clone())
        .collect()
}

#[test]
fn rename_and_reorder_preserves_relationship_ids_and_parts() {
    let fixture = fixture_bytes();
    let fixture_path = fixture_path();

    let mut doc = load_from_bytes(&fixture).expect("load fixture");
    assert_eq!(doc.workbook.sheets.len(), 2);
    assert_eq!(doc.workbook.sheets[0].xlsx_sheet_id, Some(1));
    assert_eq!(doc.workbook.sheets[0].xlsx_rel_id.as_deref(), Some("rId1"));

    let sheet2_id = doc.workbook.sheets[1].id;
    doc.workbook
        .rename_sheet(sheet2_id, "Second")
        .expect("rename");
    assert!(doc.workbook.reorder_sheet(sheet2_id, 0));

    let saved = doc.save_to_vec().expect("save");

    assert_sheets_have_backing_parts(&saved);

    let entries = workbook_sheet_entries(&zip_part(&saved, "xl/workbook.xml"));
    assert_eq!(
        entries,
        vec![
            ("Second".to_string(), 2, "rId2".to_string()),
            ("Sheet1".to_string(), 1, "rId1".to_string()),
        ]
    );

    let parts = diff_parts(&fixture_path, &saved);
    assert_eq!(parts, BTreeSet::from(["xl/workbook.xml".to_string()]));
}

#[test]
fn add_sheet_creates_part_and_updates_rels_and_content_types() {
    let fixture = fixture_bytes();
    let fixture_path = fixture_path();

    let mut doc = load_from_bytes(&fixture).expect("load fixture");
    doc.workbook.add_sheet("Added").unwrap();

    let saved = doc.save_to_vec().expect("save");

    assert_sheets_have_backing_parts(&saved);

    let entries = workbook_sheet_entries(&zip_part(&saved, "xl/workbook.xml"));
    assert_eq!(entries.len(), 3);
    assert_eq!(entries[2].0, "Added");
    assert_eq!(entries[2].1, 3);
    assert_eq!(entries[2].2, "rId4");

    let rels = workbook_relationship_targets(&zip_part(&saved, "xl/_rels/workbook.xml.rels"));
    assert_eq!(
        rels.get("rId4").map(String::as_str),
        Some("worksheets/sheet3.xml")
    );

    let content_types = String::from_utf8(zip_part(&saved, "[Content_Types].xml")).expect("utf8");
    assert!(content_types.contains(r#"/xl/worksheets/sheet3.xml"#));

    let parts = diff_parts(&fixture_path, &saved);
    assert_eq!(
        parts,
        BTreeSet::from([
            "[Content_Types].xml".to_string(),
            "xl/_rels/workbook.xml.rels".to_string(),
            "xl/workbook.xml".to_string(),
            "xl/worksheets/sheet3.xml".to_string(),
        ])
    );
}

#[test]
fn delete_sheet_removes_part_and_relationship() {
    let fixture = fixture_bytes();
    let fixture_path = fixture_path();

    let mut doc = load_from_bytes(&fixture).expect("load fixture");
    let sheet1_id = doc.workbook.sheets[0].id;
    doc.workbook.sheets.retain(|s| s.id != sheet1_id);

    let saved = doc.save_to_vec().expect("save");

    assert_sheets_have_backing_parts(&saved);

    let entries = workbook_sheet_entries(&zip_part(&saved, "xl/workbook.xml"));
    assert_eq!(
        entries,
        vec![("Sheet2".to_string(), 2, "rId2".to_string())]
    );

    let rels = workbook_relationship_targets(&zip_part(&saved, "xl/_rels/workbook.xml.rels"));
    assert!(!rels.contains_key("rId1"));

    let cursor = Cursor::new(&saved);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    assert!(
        archive.by_name("xl/worksheets/sheet1.xml").is_err(),
        "expected deleted sheet part to be removed"
    );

    let content_types = String::from_utf8(zip_part(&saved, "[Content_Types].xml")).expect("utf8");
    assert!(!content_types.contains(r#"/xl/worksheets/sheet1.xml"#));

    let parts = diff_parts(&fixture_path, &saved);
    assert_eq!(
        parts,
        BTreeSet::from([
            "[Content_Types].xml".to_string(),
            "xl/_rels/workbook.xml.rels".to_string(),
            "xl/workbook.xml".to_string(),
            "xl/worksheets/sheet1.xml".to_string(),
        ])
    );
}

#[test]
fn add_sheet_preserves_macro_relationships_and_content_types() {
    const FIXTURE: &[u8] = include_bytes!("fixtures/rt_macro.xlsm");

    let mut doc = load_from_bytes(FIXTURE).expect("load fixture");
    doc.workbook.add_sheet("Added").unwrap();

    let saved = doc.save_to_vec().expect("save");

    assert_eq!(
        zip_part(FIXTURE, "xl/vbaProject.bin"),
        zip_part(&saved, "xl/vbaProject.bin"),
        "vbaProject.bin must be preserved byte-for-byte"
    );

    let rels = String::from_utf8(zip_part(&saved, "xl/_rels/workbook.xml.rels")).expect("utf8");
    assert!(
        rels.contains("relationships/vbaProject"),
        "workbook.xml.rels must retain the vbaProject relationship"
    );

    let content_types = String::from_utf8(zip_part(&saved, "[Content_Types].xml")).expect("utf8");
    assert!(
        content_types.contains("application/vnd.ms-office.vbaProject"),
        "[Content_Types].xml must retain vbaProject.bin override"
    );
    assert!(
        content_types.contains("macroEnabled.main+xml"),
        "[Content_Types].xml must retain macro-enabled workbook content type"
    );
}

#[test]
fn sheet_visibility_roundtrips_to_workbook_xml_state() {
    let fixture = fixture_bytes();
    let fixture_path = fixture_path();

    let mut doc = load_from_bytes(&fixture).expect("load fixture");
    let sheet2_id = doc.workbook.sheets[1].id;
    assert!(doc
        .workbook
        .set_sheet_visibility(sheet2_id, SheetVisibility::Hidden));

    let saved = doc.save_to_vec().expect("save");

    let states = workbook_sheet_states(&zip_part(&saved, "xl/workbook.xml"));
    assert_eq!(states.get("Sheet1").and_then(|s| s.as_deref()), None);
    assert_eq!(
        states.get("Sheet2").and_then(|s| s.as_deref()),
        Some("hidden")
    );

    let parts = diff_parts(&fixture_path, &saved);
    assert_eq!(parts, BTreeSet::from(["xl/workbook.xml".to_string()]));
}

#[test]
fn preserves_sheet_ids_when_internal_ids_change() {
    let fixture = fixture_bytes();

    let mut doc = load_from_bytes(&fixture).expect("load fixture");
    for sheet in &mut doc.workbook.sheets {
        // Simulate a reconstructed workbook model that kept XLSX identifiers but got new internal ids.
        sheet.id = sheet.id + 1000;
    }

    let saved = doc.save_to_vec().expect("save");
    let entries = workbook_sheet_entries(&zip_part(&saved, "xl/workbook.xml"));
    assert_eq!(
        entries,
        vec![
            ("Sheet1".to_string(), 1, "rId1".to_string()),
            ("Sheet2".to_string(), 2, "rId2".to_string()),
        ]
    );

    let rels = workbook_relationship_targets(&zip_part(&saved, "xl/_rels/workbook.xml.rels"));
    assert_eq!(
        rels.get("rId1").map(String::as_str),
        Some("worksheets/sheet1.xml")
    );
    assert_eq!(
        rels.get("rId2").map(String::as_str),
        Some("worksheets/sheet2.xml")
    );
    assert!(
        !rels.values().any(|target| target == "sharedStrings.xml"),
        "should not introduce sharedStrings relationship when original used inline strings"
    );

    let cursor = Cursor::new(&saved);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    assert!(
        archive.by_name("xl/sharedStrings.xml").is_err(),
        "should not create sharedStrings.xml when original used inline strings"
    );

    let content_types = String::from_utf8(zip_part(&saved, "[Content_Types].xml")).expect("utf8");
    assert!(!content_types.contains("/xl/sharedStrings.xml"));
}

#[test]
fn add_sheet_preserves_richdata_relationships_and_content_types() {
    let rich_value = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><richValue/>"#.to_vec();
    let rich_value_rel =
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><richValueRel/>"#.to_vec();
    let rich_value_types =
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><richValueTypes/>"#.to_vec();
    let rich_value_structure = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><richValueStructure/>"#.to_vec();

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://example.com/relationships/richData" Target="richData/richValueRel.xml"/>
</Relationships>
"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/richData/richValue.xml" ContentType="application/xml"/>
  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/xml"/>
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/xml"/>
  <Override PartName="/xl/richData/richValueStructure.xml" ContentType="application/xml"/>
</Types>
"#;

    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
</worksheet>
"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);
    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_bytes()),
        ("xl/workbook.xml", workbook_xml.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/worksheets/sheet1.xml", sheet1_xml.as_bytes()),
        ("xl/richData/richValue.xml", rich_value.as_slice()),
        ("xl/richData/richValueRel.xml", rich_value_rel.as_slice()),
        ("xl/richData/richValueTypes.xml", rich_value_types.as_slice()),
        ("xl/richData/richValueStructure.xml", rich_value_structure.as_slice()),
    ] {
        zip.start_file(name, options).expect("start file");
        zip.write_all(bytes).expect("write part");
    }
    let fixture = zip.finish().expect("finish zip").into_inner();

    let mut doc = load_from_bytes(&fixture).expect("load fixture");
    doc.workbook.add_sheet("Added").unwrap();

    let saved = doc.save_to_vec().expect("save");

    assert_eq!(
        zip_part(&saved, "xl/richData/richValue.xml"),
        rich_value,
        "richValue.xml must be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/richValueRel.xml"),
        rich_value_rel,
        "richValueRel.xml must be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/richValueTypes.xml"),
        rich_value_types,
        "richValueTypes.xml must be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/richValueStructure.xml"),
        rich_value_structure,
        "richValueStructure.xml must be preserved byte-for-byte"
    );

    let rels = workbook_relationship_targets(&zip_part(&saved, "xl/_rels/workbook.xml.rels"));
    assert_eq!(
        rels,
        BTreeMap::from([
            ("rId1".to_string(), "worksheets/sheet1.xml".to_string()),
            ("rId2".to_string(), "styles.xml".to_string()),
            (
                "rId3".to_string(),
                "richData/richValueRel.xml".to_string()
            ),
            ("rId4".to_string(), "worksheets/sheet2.xml".to_string()),
        ]),
        "workbook.xml.rels must keep RichData relationship and only add a new worksheet relationship"
    );

    let content_types_out = String::from_utf8(zip_part(&saved, "[Content_Types].xml")).expect("utf8");
    for part in [
        "/xl/richData/richValue.xml",
        "/xl/richData/richValueRel.xml",
        "/xl/richData/richValueTypes.xml",
        "/xl/richData/richValueStructure.xml",
    ] {
        assert!(
            content_types_out.contains(part),
            "[Content_Types].xml must retain RichData override for {part}"
        );
    }
}
