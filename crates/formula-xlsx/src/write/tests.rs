use super::*;

use quick_xml::events::Event;
use quick_xml::Reader;
use std::collections::BTreeMap;
use std::io::{Read, Write};
use zip::write::{FileOptions, ZipWriter};
use zip::ZipArchive;

fn worksheet_formula_texts_from_xlsx(bytes: &[u8], part_name: &str) -> Vec<String> {
    let cursor = std::io::Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(part_name).expect("worksheet part missing");
    let mut xml = String::new();
    file.read_to_string(&mut xml).expect("read worksheet xml");

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut formulas = Vec::new();

    let mut in_f = false;
    let mut current = String::new();
    loop {
        match reader.read_event_into(&mut buf).expect("xml parse") {
            Event::Start(e) if e.name().as_ref() == b"f" => {
                in_f = true;
                current.clear();
            }
            Event::Empty(e) if e.name().as_ref() == b"f" => {
                formulas.push(String::new());
            }
            Event::Text(t) if in_f => {
                current.push_str(&t.unescape().expect("unescape").into_owned());
            }
            Event::End(e) if e.name().as_ref() == b"f" => {
                in_f = false;
                formulas.push(current.clone());
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    formulas
}

fn build_minimal_xlsx_with_sheet1(sheet1_xml: &str) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let cursor = std::io::Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);

    zip.start_file("xl/workbook.xml", options)
        .expect("start xl/workbook.xml");
    zip.write_all(workbook_xml.as_bytes())
        .expect("write xl/workbook.xml");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("start xl/_rels/workbook.xml.rels");
    zip.write_all(workbook_rels.as_bytes())
        .expect("write xl/_rels/workbook.xml.rels");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("start xl/worksheets/sheet1.xml");
    zip.write_all(sheet1_xml.as_bytes())
        .expect("write xl/worksheets/sheet1.xml");

    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn writes_spreadsheetml_formula_text_without_leading_equals() {
    let mut workbook = formula_model::Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1".to_string()).unwrap();
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
    let a1 = formula_model::CellRef::from_a1("A1").unwrap();
    sheet.set_formula(a1, Some("1+1".to_string()));

    let mut doc = crate::XlsxDocument::new(workbook);

    // Simulate stale/incorrect `FormulaMeta` coming from a caller: the `<f>` text must
    // never contain a leading '='.
    doc.meta.cell_meta.insert(
        (sheet_id, a1),
        crate::CellMeta {
            formula: Some(crate::FormulaMeta {
                file_text: "=1+1".to_string(),
                ..Default::default()
            }),
            ..Default::default()
        },
    );

    let bytes = write_to_vec(&doc).expect("write doc");
    let formulas = worksheet_formula_texts_from_xlsx(&bytes, "xl/worksheets/sheet1.xml");
    for f in formulas.into_iter().filter(|f| !f.is_empty()) {
        assert!(
            !f.trim_start().starts_with('='),
            "SpreadsheetML <f> text must not start with '=' (got {f:?})"
        );
    }
}

#[test]
fn sheetdata_patch_emits_vm_cm_for_inserted_cells() {
    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#;
    let input = build_minimal_xlsx_with_sheet1(sheet1_xml);

    let mut doc = crate::load_from_bytes(&input).expect("load minimal xlsx");

    let sheet_id = doc
        .workbook
        .sheets
        .first()
        .map(|s| s.id)
        .expect("sheet exists");

    let b1 = formula_model::CellRef::from_a1("B1").unwrap();
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        // Rich-data cells (entities/images) are typically represented in worksheets as a cached
        // `#VALUE!` placeholder plus a `vm="..."` pointer into `xl/metadata.xml`.
        .set_value(
            b1,
            formula_model::CellValue::Error(formula_model::ErrorValue::Value),
        );

    doc.meta.cell_meta.insert(
        (sheet_id, b1),
        crate::CellMeta {
            vm: Some("1".to_string()),
            cm: Some("2".to_string()),
            ..Default::default()
        },
    );

    let out = write_to_vec(&doc).expect("write patched xlsx");

    let cursor = std::io::Cursor::new(out);
    let mut archive = ZipArchive::new(cursor).expect("open output zip");
    let mut file = archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet part missing");
    let mut xml = String::new();
    file.read_to_string(&mut xml).expect("read worksheet xml");

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut found = false;
    let mut found_vm: Option<String> = None;
    let mut found_cm: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf).expect("xml parse") {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"c" => {
                let mut r: Option<String> = None;
                let mut vm: Option<String> = None;
                let mut cm: Option<String> = None;

                for attr in e.attributes() {
                    let attr = attr.expect("attr");
                    let v = attr.unescape_value().expect("attr value").into_owned();
                    match attr.key.as_ref() {
                        b"r" => r = Some(v),
                        b"vm" => vm = Some(v),
                        b"cm" => cm = Some(v),
                        _ => {}
                    }
                }

                if r.as_deref() == Some("B1") {
                    found = true;
                    found_vm = vm;
                    found_cm = cm;
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    assert!(found, "expected to find <c r=\"B1\"> in patched worksheet");
    assert_eq!(found_vm.as_deref(), Some("1"), "missing/incorrect vm= on B1");
    assert_eq!(found_cm.as_deref(), Some("2"), "missing/incorrect cm= on B1");
}

#[test]
fn writes_cell_vm_cm_attributes_from_cell_meta_when_rendering_sheet_data() {
    let mut workbook = formula_model::Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1".to_string()).unwrap();
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
    let a1 = formula_model::CellRef::from_a1("A1").unwrap();
    sheet.set_value(a1, formula_model::CellValue::Number(1.0));

    let mut doc = crate::XlsxDocument::new(workbook);
    doc.meta.cell_meta.insert(
        (sheet_id, a1),
        crate::CellMeta {
            vm: Some("1".to_string()),
            cm: Some("2".to_string()),
            ..Default::default()
        },
    );

    let bytes = write_to_vec(&doc).expect("write doc");
    let cursor = std::io::Cursor::new(&bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet part missing");
    let mut xml = String::new();
    file.read_to_string(&mut xml).expect("read worksheet xml");

    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut found = false;

    loop {
        match reader.read_event_into(&mut buf).expect("xml parse") {
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"c" => {
                let mut r = None;
                let mut vm = None;
                let mut cm = None;
                for attr in e.attributes() {
                    let attr = attr.expect("attr");
                    match attr.key.as_ref() {
                        b"r" => r = Some(attr.unescape_value().expect("unescape").into_owned()),
                        b"vm" => vm = Some(attr.unescape_value().expect("unescape").into_owned()),
                        b"cm" => cm = Some(attr.unescape_value().expect("unescape").into_owned()),
                        _ => {}
                    }
                }

                if r.as_deref() == Some("A1") {
                    assert_eq!(vm.as_deref(), Some("1"));
                    assert_eq!(cm.as_deref(), Some("2"));
                    found = true;
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    assert!(found, "expected to find <c r=\"A1\"> in worksheet xml");
}

#[test]
fn ensure_content_types_default_inserts_png() {
    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    let minimal = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
        r#"<Default Extension="xml" ContentType="application/xml"/>"#,
        r#"</Types>"#
    );
    parts.insert("[Content_Types].xml".to_string(), minimal.as_bytes().to_vec());

    ensure_content_types_default(&mut parts, "png", "image/png").expect("insert png default");

    let xml = std::str::from_utf8(parts.get("[Content_Types].xml").unwrap()).unwrap();
    let entry = r#"<Default Extension="png" ContentType="image/png"/>"#;
    assert!(xml.contains(entry));
    assert_eq!(xml.matches(r#"Extension="png""#).count(), 1);

    let idx_entry = xml.find(entry).unwrap();
    let idx_close = xml.rfind("</Types>").unwrap();
    assert!(idx_entry < idx_close);
}

#[test]
fn ensure_content_types_default_idempotent() {
    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    let minimal = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
        r#"<Default Extension="xml" ContentType="application/xml"/>"#,
        r#"</Types>"#
    );
    parts.insert("[Content_Types].xml".to_string(), minimal.as_bytes().to_vec());

    ensure_content_types_default(&mut parts, "png", "image/png").expect("first insert");
    let once = parts.get("[Content_Types].xml").cloned().unwrap();
    ensure_content_types_default(&mut parts, "png", "image/png").expect("second insert");
    let twice = parts.get("[Content_Types].xml").cloned().unwrap();
    assert_eq!(once, twice);
}

#[test]
fn ensure_content_types_default_noops_when_content_types_part_missing() {
    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    ensure_content_types_default(&mut parts, "png", "image/png").expect("no-op");
    assert!(
        !parts.contains_key("[Content_Types].xml"),
        "helper must not synthesize [Content_Types].xml when missing"
    );
}

#[test]
fn ensure_content_types_default_does_not_false_positive_on_extension_substrings() {
    let ct_xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        r#"<Default Extension="xpng" ContentType="application/x-xpng"/>"#,
        r#"</Types>"#
    );

    let mut parts = BTreeMap::new();
    parts.insert("[Content_Types].xml".to_string(), ct_xml.as_bytes().to_vec());

    ensure_content_types_default(&mut parts, "png", "image/png").expect("insert png default");

    let updated = std::str::from_utf8(parts.get("[Content_Types].xml").unwrap()).unwrap();
    assert!(
        updated.contains(r#"<Default Extension="png" ContentType="image/png"/>"#),
        "expected png default entry to be inserted when only xpng exists"
    );
    assert_eq!(updated.matches(r#"Extension="png""#).count(), 1);
}

#[test]
fn ensure_content_types_default_preserves_prefix_only_content_types() {
    let ct_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
</ct:Types>"#;

    let mut parts = BTreeMap::new();
    parts.insert("[Content_Types].xml".to_string(), ct_xml.as_bytes().to_vec());

    ensure_content_types_default(&mut parts, "png", "image/png").expect("patch content types");

    let updated = std::str::from_utf8(parts.get("[Content_Types].xml").expect("ct part"))
        .expect("utf8 ct xml");

    assert!(
        updated.contains(r#"<ct:Default Extension="png" ContentType="image/png"/>"#),
        "expected inserted ct:Default; got:\n{updated}"
    );
    assert!(
        !updated.contains(r#"<Default Extension="png""#),
        "must not introduce namespace-less <Default> elements; got:\n{updated}"
    );

    for name in default_element_names(updated) {
        assert!(
            name.starts_with(b"ct:"),
            "expected only prefixed Default elements; saw {:?} in:\n{updated}",
            String::from_utf8_lossy(&name)
        );
    }
}

fn override_element_names(xml: &str) -> Vec<Vec<u8>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).expect("xml parse") {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"Override" => {
                out.push(e.name().as_ref().to_vec());
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

fn default_element_names(xml: &str) -> Vec<Vec<u8>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).expect("xml parse") {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"Default" => {
                out.push(e.name().as_ref().to_vec());
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    out
}

#[test]
fn ensure_content_types_override_preserves_prefix_only_content_types() {
    let ct_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</ct:Types>"#;

    let mut parts = BTreeMap::new();
    parts.insert("[Content_Types].xml".to_string(), ct_xml.as_bytes().to_vec());

    ensure_content_types_override(
        &mut parts,
        "/xl/styles.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
    )
    .expect("patch content types");

    let updated = std::str::from_utf8(parts.get("[Content_Types].xml").expect("ct part"))
        .expect("utf8 ct xml");

    assert!(
        updated.contains(r#"<ct:Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#),
        "expected inserted ct:Override; got:\n{updated}"
    );
    assert!(
        !updated.contains("<Override"),
        "must not introduce namespace-less <Override> elements; got:\n{updated}"
    );

    for name in override_element_names(updated) {
        assert!(
            name.starts_with(b"ct:"),
            "expected only prefixed Override elements; saw {:?} in:\n{updated}",
            String::from_utf8_lossy(&name)
        );
    }
}

#[test]
fn patch_content_types_for_sheet_edits_preserves_prefix_only_content_types() {
    let ct_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</ct:Types>"#;

    let mut parts = BTreeMap::new();
    parts.insert("[Content_Types].xml".to_string(), ct_xml.as_bytes().to_vec());

    let added = vec![SheetMeta {
        worksheet_id: 1,
        sheet_id: 1,
        relationship_id: "rId1".to_string(),
        state: None,
        path: "xl/worksheets/sheet2.xml".to_string(),
    }];

    patch_content_types_for_sheet_edits(&mut parts, &[], &added).expect("patch content types");

    let updated = std::str::from_utf8(parts.get("[Content_Types].xml").expect("ct part"))
        .expect("utf8 ct xml");

    assert!(
        updated.contains(r#"<ct:Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#),
        "expected inserted ct:Override for worksheet; got:\n{updated}"
    );
    assert!(
        !updated.contains("<Override"),
        "must not introduce namespace-less <Override> elements; got:\n{updated}"
    );

    for name in override_element_names(updated) {
        assert!(
            name.starts_with(b"ct:"),
            "expected only prefixed Override elements; saw {:?} in:\n{updated}",
            String::from_utf8_lossy(&name)
        );
    }
}
