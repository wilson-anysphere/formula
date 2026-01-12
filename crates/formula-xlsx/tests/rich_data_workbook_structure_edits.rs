use std::collections::BTreeSet;
use std::io::{Cursor, Read, Write};

use base64::Engine as _;
use formula_xlsx::load_from_bytes;
use quick_xml::events::Event;
use quick_xml::Reader;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn build_rich_data_fixture_xlsx() -> Vec<u8> {
    // This fixture is intentionally minimal, but includes:
    // - Workbook relationship -> xl/metadata.xml (with a high rId to stress ID allocation).
    // - Rich data parts + an image referenced via richValueRel.xml.rels.
    //
    // The test exercises workbook structure edits, which rewrite xl/workbook.xml and
    // xl/_rels/workbook.xml.rels.

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>"#;

    // Use rId99 for metadata to ensure we don't accidentally collide with newly created sheet rels.
    //
    // Also include an unrelated low-numbered relationship (styles: rId3) so buggy sheet-id allocation
    // strategies like "max sheet rId + 1" would collide (real XLSX files typically have these).
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
  <Override PartName="/xl/richData/richValue.xml" ContentType="application/vnd.ms-excel.richvalue+xml"/>
  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
</Types>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="0"/>
</metadata>"#;

    // Link metadata -> richData parts so they are part of the OPC graph.
    let metadata_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue" Target="richData/richValue.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel" Target="richData/richValueRel.xml"/>
</Relationships>"#;

    let rich_value_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rv:richValue xmlns:rv="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
</rv:richValue>"#;

    let rich_value_rel_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rv:richValueRel xmlns:rv="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
</rv:richValueRel>"#;

    let rich_value_rel_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#;

    // 1x1 transparent PNG.
    let image_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO5j/1cAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet2.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/metadata.xml.rels", options).unwrap();
    zip.write_all(metadata_rels.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValue.xml", options).unwrap();
    zip.write_all(rich_value_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValueRel.xml", options).unwrap();
    zip.write_all(rich_value_rel_xml.as_bytes()).unwrap();

    zip.start_file("xl/richData/_rels/richValueRel.xml.rels", options)
        .unwrap();
    zip.write_all(rich_value_rel_rels.as_bytes()).unwrap();

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(&image_bytes).unwrap();

    zip.finish().unwrap().into_inner()
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn zip_file_names(zip_bytes: &[u8]) -> BTreeSet<String> {
    let cursor = Cursor::new(zip_bytes);
    let archive = ZipArchive::new(cursor).expect("open zip");
    archive.file_names().map(|s| s.to_string()).collect()
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

#[derive(Debug)]
struct Relationship {
    id: String,
    target: String,
    type_uri: Option<String>,
}

fn workbook_relationships(xml: &[u8]) -> Vec<Relationship> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).expect("read xml") {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"Relationship" => {
                let mut id = None;
                let mut type_uri = None;
                let mut target = None;
                for attr in e.attributes().flatten() {
                    let val = attr.unescape_value().expect("attr").into_owned();
                    match attr.key.as_ref() {
                        b"Id" => id = Some(val),
                        b"Type" => type_uri = Some(val),
                        b"Target" => target = Some(val),
                        _ => {}
                    }
                }
                if let (Some(id), Some(target)) = (id, target) {
                    out.push(Relationship {
                        id,
                        target,
                        type_uri,
                    });
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
fn workbook_structure_edits_preserve_rich_data_parts_and_metadata_relationship() {
    let fixture = build_rich_data_fixture_xlsx();
    let mut doc = load_from_bytes(&fixture).expect("load fixture");

    assert_eq!(doc.workbook.sheets.len(), 2, "fixture must contain 2 sheets");

    // Rename, reorder, and add a new sheet.
    let sheet2_id = doc.workbook.sheets[1].id;
    doc.workbook
        .rename_sheet(sheet2_id, "Second")
        .expect("rename");
    assert!(
        doc.workbook.reorder_sheet(sheet2_id, 0),
        "reorder should succeed"
    );
    doc.workbook.add_sheet("Added").expect("add sheet");

    let saved = doc.save_to_vec().expect("save");

    // Assert richData + media parts still exist in the output package.
    let names = zip_file_names(&saved);
    for part in [
        "xl/metadata.xml",
        "xl/_rels/metadata.xml.rels",
        "xl/richData/richValue.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
        "xl/media/image1.png",
    ] {
        assert!(names.contains(part), "missing expected part {part}");
        // Parts should be preserved byte-for-byte (we don't understand them, so we should not
        // rewrite them).
        assert_eq!(
            zip_part(&fixture, part),
            zip_part(&saved, part),
            "part {part} must be preserved byte-for-byte"
        );
    }

    // Assert workbook.xml.rels still contains the metadata relationship with the same Id/Target.
    let rels = workbook_relationships(&zip_part(&saved, "xl/_rels/workbook.xml.rels"));
    let rel_ids: BTreeSet<&str> = rels.iter().map(|r| r.id.as_str()).collect();
    assert_eq!(
        rel_ids.len(),
        rels.len(),
        "workbook.xml.rels contains duplicate Relationship Ids"
    );
    let meta_rels: Vec<&Relationship> = rels.iter().filter(|r| r.id == "rId99").collect();
    assert_eq!(
        meta_rels.len(),
        1,
        "expected exactly one workbook relationship with Id=rId99"
    );
    assert_eq!(meta_rels[0].target, "metadata.xml");
    assert_eq!(
        meta_rels[0].type_uri.as_deref(),
        Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata"),
        "metadata relationship Type must be preserved"
    );

    // Ensure we didn't drop or overwrite unrelated workbook relationships when inserting sheets.
    // This specifically guards against rId allocation strategies like `max(sheet_rId) + 1` which
    // would collide with common low-numbered relationships (e.g. styles).
    let style_rels: Vec<&Relationship> = rels.iter().filter(|r| r.id == "rId3").collect();
    assert_eq!(
        style_rels.len(),
        1,
        "expected exactly one workbook relationship with Id=rId3 (styles)"
    );
    assert_eq!(style_rels[0].target, "styles.xml");
    assert_eq!(
        style_rels[0].type_uri.as_deref(),
        Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"),
        "styles relationship Type must be preserved"
    );

    // And ensure it did not collide with a sheet relationship id.
    let sheet_entries = workbook_sheet_entries(&zip_part(&saved, "xl/workbook.xml"));
    assert!(
        !sheet_entries.iter().any(|(_, _, rid)| rid == "rId99"),
        "a sheet unexpectedly reused metadata relationship id rId99"
    );
    assert!(
        !sheet_entries.iter().any(|(_, _, rid)| rid == "rId3"),
        "a sheet unexpectedly reused styles relationship id rId3"
    );

    // Workbook structure edits also patch `[Content_Types].xml` for sheet insertions/removals; ensure
    // we didn't accidentally drop the richData-related overrides.
    let content_types =
        String::from_utf8(zip_part(&saved, "[Content_Types].xml")).expect("[Content_Types].xml utf8");
    assert!(
        content_types.contains(r#"PartName="/xl/metadata.xml""#),
        "[Content_Types].xml missing override for /xl/metadata.xml"
    );
    assert!(
        content_types.contains(r#"PartName="/xl/richData/richValue.xml""#),
        "[Content_Types].xml missing override for /xl/richData/richValue.xml"
    );
    assert!(
        content_types.contains(r#"PartName="/xl/richData/richValueRel.xml""#),
        "[Content_Types].xml missing override for /xl/richData/richValueRel.xml"
    );
}
