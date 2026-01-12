use std::io::{Cursor, Read, Write};

use formula_xlsx::load_from_bytes;
use pretty_assertions::assert_eq;
use quick_xml::events::Event;
use quick_xml::Reader;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

// Regression test: workbook sheet structure edits (add/delete) must not drop the
// linked-data-type infrastructure parts used by Excel rich data types
// (`xl/metadata.xml` + `xl/richData/*`) nor their workbook-level relationship and
// content type overrides.

fn build_fixture_xlsx() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>
"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.ms-excel.metadata+xml"/>
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/vnd.ms-excel.richValueTypes+xml"/>
  <Override PartName="/xl/richData/richValues.xml" ContentType="application/vnd.ms-excel.richValues+xml"/>
</Types>
"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"#;

    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="XLD" minSupportedVersion="0"/>
  </metadataTypes>
</metadata>
"#;

    let metadata_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2020/relationships/richValueTypes" Target="../richData/richValueTypes.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2020/relationships/richValues" Target="../richData/richValues.xml"/>
</Relationships>
"#;

    let rich_value_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypes xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rvType name="ExampleType"/>
</rvTypes>
"#;

    let rich_values_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv value="ExampleValue"/>
</rvData>
"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options)
        .expect("zip file");
    zip.write_all(content_types.as_bytes())
        .expect("zip write");

    zip.start_file("_rels/.rels", options).expect("zip file");
    zip.write_all(root_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/workbook.xml", options)
        .expect("zip file");
    zip.write_all(workbook_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("zip file");
    zip.write_all(workbook_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("zip file");
    zip.write_all(sheet_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/worksheets/sheet2.xml", options)
        .expect("zip file");
    zip.write_all(sheet_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/metadata.xml", options).expect("zip file");
    zip.write_all(metadata_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/_rels/metadata.xml.rels", options)
        .expect("zip file");
    zip.write_all(metadata_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/richData/richValueTypes.xml", options)
        .expect("zip file");
    zip.write_all(rich_value_types_xml.as_bytes())
        .expect("zip write");

    zip.start_file("xl/richData/richValues.xml", options)
        .expect("zip file");
    zip.write_all(rich_values_xml.as_bytes())
        .expect("zip write");

    zip.finish().expect("finish zip").into_inner()
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn workbook_sheets_with_rids(xml: &[u8]) -> Vec<(String, String)> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).expect("read xml") {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"sheet" => {
                let mut name = None;
                let mut rid = None;
                for attr in e.attributes().flatten() {
                    let val = attr.unescape_value().expect("attr").into_owned();
                    match attr.key.as_ref() {
                        b"name" => name = Some(val),
                        b"r:id" => rid = Some(val),
                        _ => {}
                    }
                }
                if let (Some(name), Some(rid)) = (name, rid) {
                    out.push((name, rid));
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn workbook_relationship_targets(xml: &[u8]) -> std::collections::BTreeMap<String, String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = std::collections::BTreeMap::new();
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
                    out.insert(id, target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn content_type_overrides(xml: &[u8]) -> std::collections::BTreeSet<String> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = std::collections::BTreeSet::new();
    loop {
        match reader.read_event_into(&mut buf).expect("read xml") {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"Override" => {
                for attr in e.attributes().flatten() {
                    if attr.key.as_ref() == b"PartName" {
                        out.insert(attr.unescape_value().expect("attr").into_owned());
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

#[test]
fn sheet_edits_preserve_richdata_parts_and_relationships() {
    let fixture = build_fixture_xlsx();

    let original_metadata = zip_part(&fixture, "xl/metadata.xml");
    let original_metadata_rels = zip_part(&fixture, "xl/_rels/metadata.xml.rels");
    let original_rich_value_types = zip_part(&fixture, "xl/richData/richValueTypes.xml");
    let original_rich_values = zip_part(&fixture, "xl/richData/richValues.xml");

    let mut doc = load_from_bytes(&fixture).expect("load fixture");
    assert_eq!(doc.workbook.sheets.len(), 2);

    let sheet2_id = doc.workbook.sheets[1].id;
    doc.workbook.delete_sheet(sheet2_id).expect("delete sheet2");
    doc.workbook.add_sheet("Added").expect("add sheet");

    let saved = doc.save_to_vec().expect("save");

    // Sanity: ensure we actually exercised the sheet-structure rewrite path.
    let workbook_xml = zip_part(&saved, "xl/workbook.xml");
    let workbook_sheet_rids = workbook_sheets_with_rids(&workbook_xml);
    assert!(
        !workbook_sheet_rids.iter().any(|(name, _)| name == "Sheet2"),
        "expected deleted sheet to be removed from xl/workbook.xml"
    );
    assert!(
        workbook_sheet_rids.iter().any(|(name, _)| name == "Added"),
        "expected newly-added sheet to be present in xl/workbook.xml"
    );
    let cursor = Cursor::new(&saved);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    assert!(
        archive.by_name("xl/worksheets/sheet2.xml").is_err(),
        "expected deleted sheet2 part to be removed from output package"
    );

    assert_eq!(
        zip_part(&saved, "xl/metadata.xml"),
        original_metadata,
        "xl/metadata.xml must be preserved byte-for-byte across sheet edits"
    );
    assert_eq!(
        zip_part(&saved, "xl/_rels/metadata.xml.rels"),
        original_metadata_rels,
        "xl/_rels/metadata.xml.rels must be preserved byte-for-byte across sheet edits"
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/richValueTypes.xml"),
        original_rich_value_types,
        "xl/richData/richValueTypes.xml must be preserved byte-for-byte across sheet edits"
    );
    assert_eq!(
        zip_part(&saved, "xl/richData/richValues.xml"),
        original_rich_values,
        "xl/richData/richValues.xml must be preserved byte-for-byte across sheet edits"
    );

    let workbook_rels =
        workbook_relationship_targets(&zip_part(&saved, "xl/_rels/workbook.xml.rels"));
    assert_eq!(
        workbook_rels.get("rId9").map(String::as_str),
        Some("metadata.xml"),
        "workbook.xml.rels must retain the metadata relationship (rId9 -> metadata.xml)"
    );

    let added_rid = workbook_sheet_rids
        .iter()
        .find(|(name, _)| name == "Added")
        .map(|(_, rid)| rid.as_str())
        .expect("Added sheet must have an r:id");
    let added_target = workbook_rels
        .get(added_rid)
        .expect("Added sheet r:id must exist in workbook.xml.rels");
    let added_part_name = if let Some(path) = added_target.strip_prefix('/') {
        path.to_string()
    } else {
        format!("xl/{added_target}")
    };
    assert!(
        archive.by_name(&added_part_name).is_ok(),
        "expected Added sheet backing part to exist in output package: {added_part_name}"
    );

    let content_types = content_type_overrides(&zip_part(&saved, "[Content_Types].xml"));
    assert!(
        !content_types.contains("/xl/worksheets/sheet2.xml"),
        "expected [Content_Types].xml to drop override for deleted /xl/worksheets/sheet2.xml"
    );
    assert!(
        content_types.contains("/xl/metadata.xml"),
        "[Content_Types].xml must retain override for /xl/metadata.xml"
    );
    assert!(
        content_types.contains("/xl/richData/richValueTypes.xml"),
        "[Content_Types].xml must retain override for /xl/richData/richValueTypes.xml"
    );
    assert!(
        content_types.contains("/xl/richData/richValues.xml"),
        "[Content_Types].xml must retain override for /xl/richData/richValues.xml"
    );

    let added_override_name = format!("/{added_part_name}");
    assert!(
        content_types.contains(&added_override_name),
        "expected [Content_Types].xml to include override for newly-added worksheet part: {added_override_name}"
    );
}
