use super::*;

use quick_xml::events::Event;
use quick_xml::Reader;
use roxmltree::Document;
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

fn zip_part_to_string(bytes: &[u8], part_name: &str) -> String {
    let cursor = std::io::Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(part_name).expect("part missing");
    let mut xml = String::new();
    file.read_to_string(&mut xml).expect("read part");
    xml
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
fn sheet_protection_patching_preserves_modern_hashing_attrs_and_children() {
    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
  <x:sheetProtection sheet="1" formatCells="0" objects="true" algorithmName="SHA-512" hashValue="aGFzaA==" saltValue="c2FsdA==" spinCount="100000" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><extLst><ext uri="{01234567-89AB-CDEF-0123-456789ABCDEF}"><x14ac:protection foo="bar"/></ext></extLst></x:sheetProtection>
</worksheet>"#;
    let input = build_minimal_xlsx_with_sheet1(sheet1_xml);

    let mut doc = crate::load_from_bytes(&input).expect("load minimal xlsx");
    let sheet_id = doc.workbook.sheets[0].id;
    {
        let sheet = doc.workbook.sheet_mut(sheet_id).expect("sheet exists");
        assert!(
            sheet.sheet_protection.enabled,
            "expected protection enabled"
        );
        assert!(
            sheet.sheet_protection.password_hash.is_none(),
            "fixture should not use legacy password hash"
        );
        // Flip one allow-list flag without touching the legacy password.
        sheet.sheet_protection.format_cells = true;
    }

    let out = write_to_vec(&doc).expect("write patched xlsx");
    let xml = zip_part_to_string(&out, "xl/worksheets/sheet1.xml");

    // The original worksheet used a prefixed sheetProtection tag; the patcher must preserve the
    // qualified name (including the prefix) exactly.
    assert!(
        xml.contains("<x:sheetProtection"),
        "expected sheetProtection tag prefix to be preserved:\n{xml}"
    );

    // Modern hashing attributes must be preserved.
    assert!(
        xml.contains(r#"algorithmName="SHA-512""#),
        "missing algorithmName attr after patch:\n{xml}"
    );
    assert!(
        xml.contains(r#"hashValue="aGFzaA==""#),
        "missing hashValue attr after patch:\n{xml}"
    );
    assert!(
        xml.contains(r#"saltValue="c2FsdA==""#),
        "missing saltValue attr after patch:\n{xml}"
    );
    assert!(
        xml.contains(r#"spinCount="100000""#),
        "missing spinCount attr after patch:\n{xml}"
    );

    // Nested extension list must be preserved byte-for-byte.
    let extlst = r#"<extLst><ext uri="{01234567-89AB-CDEF-0123-456789ABCDEF}"><x14ac:protection foo="bar"/></ext></extLst>"#;
    assert!(
        xml.contains(extlst),
        "sheetProtection children were not preserved:\n{xml}"
    );

    // Unchanged allow-list flags should preserve original value formatting.
    assert!(
        xml.contains(r#"objects="true""#),
        "expected objects=\"true\" to be preserved:\n{xml}"
    );

    // Edited allow-list flags must update.
    assert!(
        xml.contains(r#"formatCells="1""#),
        "expected formatCells to flip to 1:\n{xml}"
    );
}

#[test]
fn sheet_protection_disabled_removes_element() {
    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
  <sheetProtection sheet="1" formatCells="0"/>
</worksheet>"#;
    let input = build_minimal_xlsx_with_sheet1(sheet1_xml);

    let mut doc = crate::load_from_bytes(&input).expect("load minimal xlsx");
    let sheet_id = doc.workbook.sheets[0].id;
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .sheet_protection
        .enabled = false;

    let out = write_to_vec(&doc).expect("write patched xlsx");
    let xml = zip_part_to_string(&out, "xl/worksheets/sheet1.xml");
    assert!(
        !xml.contains("sheetProtection"),
        "expected sheetProtection element to be removed when disabled:\n{xml}"
    );
}

#[test]
fn clearing_conditional_formatting_strips_cf_blocks_but_preserves_other_extlst_entries() {
    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
  xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
  <conditionalFormatting sqref="A1">
    <cfRule type="expression" priority="1"><formula>TRUE</formula></cfRule>
  </conditionalFormatting>
  <extLst>
    <ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
      <x14:conditionalFormattings>
        <x14:conditionalFormatting>
          <x14:cfRule type="dataBar" priority="1" id="{00000000-0000-0000-0000-000000000001}">
            <x14:dataBar>
              <x14:cfvo type="min"/>
              <x14:cfvo type="max"/>
              <x14:negativeFillColor rgb="FFFF0000"/>
              <x14:axisColor rgb="FF00FF00"/>
            </x14:dataBar>
          </x14:cfRule>
          <xm:sqref>A1</xm:sqref>
        </x14:conditionalFormatting>
      </x14:conditionalFormattings>
    </ext>
    <ext uri="{01234567-89AB-CDEF-0123-456789ABCDEF}">
      <otherExt foo="bar"/>
    </ext>
  </extLst>
</worksheet>"#;
    let input = build_minimal_xlsx_with_sheet1(sheet1_xml);

    let mut doc = crate::load_from_bytes(&input).expect("load minimal xlsx");
    let sheet_id = doc.workbook.sheets[0].id;
    {
        let sheet = doc.workbook.sheet_mut(sheet_id).expect("sheet exists");
        assert!(
            !sheet.conditional_formatting_rules.is_empty(),
            "expected conditional formatting to be parsed from sheet1.xml"
        );
        sheet.clear_conditional_formatting();
    }

    let out = write_to_vec(&doc).expect("write patched xlsx");
    let xml = zip_part_to_string(&out, "xl/worksheets/sheet1.xml");
    assert!(
        !xml.contains("conditionalFormatting"),
        "expected conditional formatting blocks removed:\n{xml}"
    );
    assert!(
        !xml.contains("78C0D931-6437-407d-A8EE-F0AAD7539E65")
            && !xml.contains("78C0D931-6437-407D-A8EE-F0AAD7539E65"),
        "expected x14 conditional formatting extLst entry removed:\n{xml}"
    );
    assert!(
        xml.contains("{01234567-89AB-CDEF-0123-456789ABCDEF}"),
        "expected unrelated extLst entry to be preserved:\n{xml}"
    );
    assert!(
        xml.contains("otherExt") && xml.contains(r#"foo="bar""#),
        "expected unrelated extLst payload to be preserved:\n{xml}"
    );
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
    assert_eq!(
        found_vm.as_deref(),
        Some("1"),
        "missing/incorrect vm= on B1"
    );
    assert_eq!(
        found_cm.as_deref(),
        Some("2"),
        "missing/incorrect cm= on B1"
    );
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
fn writer_drops_vm_when_raw_value_no_longer_matches_cell_value() {
    let mut workbook = formula_model::Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1".to_string()).unwrap();
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
    let a1 = formula_model::CellRef::from_a1("A1").unwrap();
    sheet.set_value(a1, formula_model::CellValue::Number(1.0));

    // Attach a vm/cm pointer plus the raw `<v>` payload captured on read. When the in-memory cell
    // value changes, the writer should drop `vm` to avoid leaving a dangling metadata pointer.
    let mut doc = crate::XlsxDocument::new(workbook);
    doc.meta.cell_meta.insert(
        (sheet_id, a1),
        crate::CellMeta {
            vm: Some("1".to_string()),
            cm: Some("2".to_string()),
            raw_value: Some("1".to_string()),
            ..Default::default()
        },
    );

    // Change the cell value without updating the corresponding metadata records.
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .set_value(a1, formula_model::CellValue::Number(2.0));

    let bytes = write_to_vec(&doc).expect("write doc");
    let cursor = std::io::Cursor::new(&bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("worksheet part missing");
    let mut xml = String::new();
    file.read_to_string(&mut xml).expect("read worksheet xml");

    let doc = roxmltree::Document::parse(&xml).expect("parse sheet xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");

    assert_eq!(cell.attribute("vm"), None, "vm should be dropped: {xml}");
    assert_eq!(
        cell.attribute("cm"),
        Some("2"),
        "cm should be preserved: {xml}"
    );
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
    parts.insert(
        "[Content_Types].xml".to_string(),
        minimal.as_bytes().to_vec(),
    );

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
    parts.insert(
        "[Content_Types].xml".to_string(),
        minimal.as_bytes().to_vec(),
    );

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
    parts.insert(
        "[Content_Types].xml".to_string(),
        ct_xml.as_bytes().to_vec(),
    );

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
    parts.insert(
        "[Content_Types].xml".to_string(),
        ct_xml.as_bytes().to_vec(),
    );

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

#[test]
fn ensure_content_types_default_expands_self_closing_prefix_only_root() {
    let ct_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"/>"#;

    let mut parts = BTreeMap::new();
    parts.insert(
        "[Content_Types].xml".to_string(),
        ct_xml.as_bytes().to_vec(),
    );

    ensure_content_types_default(&mut parts, "png", "image/png").expect("patch content types");

    let updated = std::str::from_utf8(parts.get("[Content_Types].xml").expect("ct part"))
        .expect("utf8 ct xml");
    assert_parses_xml(updated);
    assert!(
        updated.contains("</ct:Types>"),
        "expected self-closing ct:Types root to be expanded; got:\n{updated}"
    );
    assert!(
        updated.contains(r#"<ct:Default Extension="png" ContentType="image/png"/>"#),
        "expected inserted ct:Default; got:\n{updated}"
    );
    assert!(
        !updated.contains(r#"<Default Extension="png""#),
        "must not introduce namespace-less <Default> elements; got:\n{updated}"
    );
}

#[test]
fn ensure_content_types_default_treats_existing_extension_case_insensitively() {
    let ct_xml = concat!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        r#"<Default Extension="PNG" ContentType="image/png"/>"#,
        r#"</Types>"#
    );

    let mut parts = BTreeMap::new();
    parts.insert(
        "[Content_Types].xml".to_string(),
        ct_xml.as_bytes().to_vec(),
    );

    let before = parts.get("[Content_Types].xml").cloned().unwrap();
    ensure_content_types_default(&mut parts, "png", "image/png").expect("no-op");
    let after = parts.get("[Content_Types].xml").cloned().unwrap();
    assert_eq!(
        before, after,
        "expected helper to detect existing PNG extension case-insensitively"
    );
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

fn relationship_element_names(xml: &str) -> Vec<Vec<u8>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).expect("xml parse") {
            Event::Start(e) | Event::Empty(e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
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

fn assert_parses_xml(xml: &str) {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => panic!("xml parse error: {err}\nxml:\n{xml}"),
        }
        buf.clear();
    }
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
    parts.insert(
        "[Content_Types].xml".to_string(),
        ct_xml.as_bytes().to_vec(),
    );

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
fn ensure_content_types_override_expands_self_closing_prefix_only_root() {
    let ct_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"/>"#;

    let mut parts = BTreeMap::new();
    parts.insert(
        "[Content_Types].xml".to_string(),
        ct_xml.as_bytes().to_vec(),
    );

    ensure_content_types_override(
        &mut parts,
        "/xl/styles.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
    )
    .expect("patch content types");

    let updated = std::str::from_utf8(parts.get("[Content_Types].xml").expect("ct part"))
        .expect("utf8 ct xml");
    assert_parses_xml(updated);
    assert!(
        updated.contains("</ct:Types>"),
        "expected self-closing ct:Types root to be expanded; got:\n{updated}"
    );
    assert!(
        updated.contains(r#"<ct:Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#),
        "expected inserted ct:Override; got:\n{updated}"
    );
    assert!(
        !updated.contains("<Override"),
        "must not introduce namespace-less <Override> elements; got:\n{updated}"
    );
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
    parts.insert(
        "[Content_Types].xml".to_string(),
        ct_xml.as_bytes().to_vec(),
    );

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

#[test]
fn relationship_target_by_type_handles_prefixed_relationship_elements() {
    let rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
  <pr:Relationship Id="rId1" Type="{REL_TYPE_STYLES}" Target="styles.xml"/>
</pr:Relationships>"#
    );

    let target = relationship_target_by_type(rels.as_bytes(), REL_TYPE_STYLES).expect("parse rels");
    assert_eq!(target.as_deref(), Some("styles.xml"));
}

#[test]
fn relationship_targets_by_type_ignores_external_relationships() {
    let rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="{rel_type}" Target="https://example.com/image.png" TargetMode="External"/>
  <Relationship Id="rId2" Type="{rel_type}" Target="../media/image1.png"/>
</Relationships>"#,
        rel_type = crate::drawings::REL_TYPE_IMAGE
    );

    let targets =
        relationship_targets_by_type(rels.as_bytes(), crate::drawings::REL_TYPE_IMAGE).expect("parse");
    assert_eq!(targets, vec!["../media/image1.png".to_string()]);
}

#[test]
fn ensure_workbook_rels_has_relationship_inserts_prefixed_relationship() {
    let rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
  <pr:Relationship Id="rId1" Type="{REL_TYPE_STYLES}" Target="styles.xml"/>
</pr:Relationships>"#
    );

    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    parts.insert(WORKBOOK_RELS_PART.to_string(), rels.into_bytes());

    ensure_workbook_rels_has_relationship(&mut parts, REL_TYPE_SHARED_STRINGS, "sharedStrings.xml")
        .expect("patch workbook rels");

    let out = String::from_utf8(
        parts
            .get(WORKBOOK_RELS_PART)
            .expect("workbook rels present")
            .clone(),
    )
    .expect("utf8");
    roxmltree::Document::parse(&out).expect("output rels should be valid xml");
    assert!(
        out.contains(REL_TYPE_SHARED_STRINGS),
        "expected sharedStrings relationship to be inserted"
    );
    assert!(
        !out.contains("<Relationship"),
        "prefix-only rels must not contain namespace-less <Relationship> tags, got:\n{out}"
    );
}

#[test]
fn ensure_workbook_rels_has_relationship_expands_self_closing_prefix_only_root() {
    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    parts.insert(WORKBOOK_RELS_PART.to_string(), rels.as_bytes().to_vec());

    ensure_workbook_rels_has_relationship(&mut parts, REL_TYPE_STYLES, "styles.xml")
        .expect("patch workbook rels");

    let out = std::str::from_utf8(
        parts
            .get(WORKBOOK_RELS_PART)
            .expect("workbook rels present")
            .as_slice(),
    )
    .expect("utf8");
    assert_parses_xml(out);
    assert!(
        out.contains("</pr:Relationships>"),
        "expected self-closing pr:Relationships root to be expanded; got:\n{out}"
    );
    assert!(
        out.contains("<pr:Relationship"),
        "expected inserted pr:Relationship; got:\n{out}"
    );
    assert!(
        !out.contains("<Relationship"),
        "prefix-only rels must not contain namespace-less <Relationship> tags, got:\n{out}"
    );
}

#[test]
fn ensure_workbook_rels_has_relationship_prefers_root_prefix_when_default_xmlns_is_unrelated() {
    let rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns="urn:unused" xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    parts.insert(WORKBOOK_RELS_PART.to_string(), rels.as_bytes().to_vec());

    ensure_workbook_rels_has_relationship(&mut parts, REL_TYPE_STYLES, "styles.xml")
        .expect("patch workbook rels");

    let out = std::str::from_utf8(
        parts
            .get(WORKBOOK_RELS_PART)
            .expect("workbook rels present")
            .as_slice(),
    )
    .expect("utf8");
    assert_parses_xml(out);
    assert!(
        out.contains("<pr:Relationship"),
        "expected inserted pr:Relationship; got:\n{out}"
    );
    assert!(
        !out.contains("<Relationship"),
        "must not introduce namespace-less <Relationship> tags, got:\n{out}"
    );
}

#[test]
fn patch_workbook_rels_for_sheet_edits_handles_prefixed_relationships() {
    let rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
  <pr:Relationship Id="rId1" Type="{WORKSHEET_REL_TYPE}" Target="worksheets/sheet1.xml"/>
  <pr:Relationship Id="rId2" Type="{REL_TYPE_STYLES}" Target="styles.xml"/>
  <pr:Relationship Id="rId3" Type="{REL_TYPE_SHARED_STRINGS}" Target="sharedStrings.xml"/>
</pr:Relationships>"#
    );

    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    parts.insert(WORKBOOK_RELS_PART.to_string(), rels.into_bytes());

    let removed = vec![SheetMeta {
        worksheet_id: 1,
        sheet_id: 1,
        relationship_id: "rId1".to_string(),
        state: None,
        path: "xl/worksheets/sheet1.xml".to_string(),
    }];
    let added = vec![SheetMeta {
        worksheet_id: 2,
        sheet_id: 2,
        relationship_id: "rId4".to_string(),
        state: None,
        path: "xl/worksheets/sheet2.xml".to_string(),
    }];

    patch_workbook_rels_for_sheet_edits(&mut parts, &removed, &added).expect("patch rels");

    let out = String::from_utf8(
        parts
            .get(WORKBOOK_RELS_PART)
            .expect("workbook rels present")
            .clone(),
    )
    .expect("utf8");
    roxmltree::Document::parse(&out).expect("output rels should be valid xml");
    assert!(
        !out.contains("<Relationship"),
        "prefix-only rels must not contain namespace-less <Relationship> tags, got:\n{out}"
    );

    let doc = roxmltree::Document::parse(&out).expect("parse output");
    let ids: Vec<&str> = doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
        .filter_map(|n| n.attribute("Id"))
        .collect();
    assert!(
        !ids.contains(&"rId1"),
        "expected removed sheet relationship to be dropped"
    );
    assert!(
        ids.contains(&"rId4"),
        "expected added sheet relationship to be inserted"
    );
}

#[test]
fn patch_workbook_xml_infers_relationships_prefix_from_sheets_namespace() {
    let mut workbook = formula_model::Workbook::new();
    workbook.add_sheet("Sheet1".to_string()).unwrap();
    let doc = crate::XlsxDocument::new(workbook);

    let workbook_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="{spreadsheetml}">
  <x:sheets xmlns:rel="{rels}">
    <x:sheet name="Sheet1" sheetId="1" rel:id="rId1"/>
  </x:sheets>
</x:workbook>"#,
        spreadsheetml = SPREADSHEETML_NS,
        rels = crate::xml::OFFICE_RELATIONSHIPS_NS
    );

    let patched = patch_workbook_xml(&doc, workbook_xml.as_bytes(), &doc.meta.sheets)
        .expect("patch workbook xml");
    let patched = std::str::from_utf8(&patched).expect("patched xml is utf8");

    Document::parse(patched).expect("patched workbook.xml must parse as XML");

    assert!(
        patched.contains("rel:id="),
        "expected patched workbook.xml to use rel:id; got:\n{patched}"
    );
    assert!(
        !patched.contains(" r:id="),
        "must not introduce undeclared r:id attributes; got:\n{patched}"
    );
}

#[test]
fn patch_workbook_xml_expands_self_closing_prefixed_root() {
    let mut workbook = formula_model::Workbook::new();
    workbook.add_sheet("Sheet1".to_string()).unwrap();
    let doc = crate::XlsxDocument::new(workbook);

    let workbook_xml = format!(
        r#"<x:workbook xmlns:x="{spreadsheetml}" xmlns:rel="{rels}"/>"#,
        spreadsheetml = SPREADSHEETML_NS,
        rels = crate::xml::OFFICE_RELATIONSHIPS_NS
    );

    let patched = patch_workbook_xml(&doc, workbook_xml.as_bytes(), &doc.meta.sheets)
        .expect("patch workbook xml");
    let patched = std::str::from_utf8(&patched).expect("patched xml is utf8");

    Document::parse(patched).expect("patched workbook.xml must parse as XML");

    assert!(
        patched.contains("<x:workbook") && patched.contains("</x:workbook>"),
        "expected expanded <x:workbook> root; got:\n{patched}"
    );
    assert!(
        patched.contains("<x:sheets") && patched.contains("<x:sheet"),
        "expected inserted x:sheets/x:sheet children; got:\n{patched}"
    );
    assert!(
        patched.contains("rel:id="),
        "expected patched workbook.xml to use rel:id; got:\n{patched}"
    );
    assert!(
        !patched.contains(" r:id="),
        "must not introduce undeclared r:id attributes; got:\n{patched}"
    );
}

#[test]
fn patch_workbook_xml_expands_self_closing_root_and_declares_r_namespace_when_missing() {
    let mut workbook = formula_model::Workbook::new();
    workbook.add_sheet("Sheet1".to_string()).unwrap();
    let doc = crate::XlsxDocument::new(workbook);

    let workbook_xml = format!(
        r#"<x:workbook xmlns:x="{spreadsheetml}"/>"#,
        spreadsheetml = SPREADSHEETML_NS
    );

    let patched = patch_workbook_xml(&doc, workbook_xml.as_bytes(), &doc.meta.sheets)
        .expect("patch workbook xml");
    let patched = std::str::from_utf8(&patched).expect("patched xml is utf8");

    Document::parse(patched).expect("patched workbook.xml must parse as XML");

    assert!(
        patched.contains(&format!(
            r#"xmlns:r="{}""#,
            crate::xml::OFFICE_RELATIONSHIPS_NS
        )),
        "expected patched workbook.xml to declare xmlns:r; got:\n{patched}"
    );
    assert!(
        patched.contains(" r:id="),
        "expected patched workbook.xml to use r:id; got:\n{patched}"
    );
    assert!(
        !patched.contains(" id=\""),
        "must not fall back to namespace-less id= attributes; got:\n{patched}"
    );
}

#[test]
fn patch_workbook_xml_expands_self_closing_default_namespace_root() {
    let mut workbook = formula_model::Workbook::new();
    workbook.add_sheet("Sheet1".to_string()).unwrap();
    let doc = crate::XlsxDocument::new(workbook);

    let workbook_xml = format!(
        r#"<workbook xmlns="{spreadsheetml}" xmlns:r="{rels}"/>"#,
        spreadsheetml = SPREADSHEETML_NS,
        rels = crate::xml::OFFICE_RELATIONSHIPS_NS
    );

    let patched = patch_workbook_xml(&doc, workbook_xml.as_bytes(), &doc.meta.sheets)
        .expect("patch workbook xml");
    let patched = std::str::from_utf8(&patched).expect("patched xml is utf8");

    Document::parse(patched).expect("patched workbook.xml must parse as XML");

    assert!(
        patched.contains("<workbook") && patched.contains("</workbook>"),
        "expected expanded <workbook> root; got:\n{patched}"
    );
    assert!(
        patched.contains("<sheets") && patched.contains("<sheet"),
        "expected inserted sheets/sheet children; got:\n{patched}"
    );
    assert!(
        patched.contains(" r:id="),
        "expected inserted sheets to use r:id; got:\n{patched}"
    );
}

#[test]
fn sheet_structure_patchers_expand_self_closing_prefix_only_roots() {
    let workbook_rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
    let content_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"/>"#;

    let mut parts = BTreeMap::new();
    parts.insert(
        WORKBOOK_RELS_PART.to_string(),
        workbook_rels_xml.as_bytes().to_vec(),
    );
    parts.insert(
        "[Content_Types].xml".to_string(),
        content_types_xml.as_bytes().to_vec(),
    );

    let added = vec![SheetMeta {
        worksheet_id: 1,
        sheet_id: 1,
        relationship_id: "rId1".to_string(),
        state: None,
        path: "xl/worksheets/sheet2.xml".to_string(),
    }];

    patch_workbook_rels_for_sheet_edits(&mut parts, &[], &added).expect("patch workbook rels");
    patch_content_types_for_sheet_edits(&mut parts, &[], &added).expect("patch content types");

    let updated_rels =
        std::str::from_utf8(parts.get(WORKBOOK_RELS_PART).expect("rels part")).expect("utf8 rels");
    assert_parses_xml(updated_rels);
    assert!(
        updated_rels.contains(r#"<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">"#)
            && updated_rels.contains("</pr:Relationships>"),
        "expected expanded Relationships root; got:\n{updated_rels}"
    );
    assert!(
        updated_rels.contains("<pr:Relationship"),
        "expected prefixed pr:Relationship; got:\n{updated_rels}"
    );
    assert!(
        !updated_rels.contains("<Relationship"),
        "must not introduce namespace-less <Relationship> elements; got:\n{updated_rels}"
    );
    let rel_names = relationship_element_names(updated_rels);
    assert_eq!(rel_names.len(), 1, "expected one inserted Relationship");
    for name in rel_names {
        assert!(
            name.starts_with(b"pr:"),
            "expected only prefixed Relationship elements; saw {:?} in:\n{updated_rels}",
            String::from_utf8_lossy(&name)
        );
    }

    let updated_ct =
        std::str::from_utf8(parts.get("[Content_Types].xml").expect("ct part")).expect("utf8 ct");
    assert_parses_xml(updated_ct);
    assert!(
        updated_ct.contains(
            r#"<ct:Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
        ),
        "expected inserted ct:Override for worksheet; got:\n{updated_ct}"
    );
    assert!(
        !updated_ct.contains("<Override"),
        "must not introduce namespace-less <Override> elements; got:\n{updated_ct}"
    );
    let override_names = override_element_names(updated_ct);
    assert_eq!(override_names.len(), 1, "expected one inserted Override");
    for name in override_names {
        assert!(
            name.starts_with(b"ct:"),
            "expected only prefixed Override elements; saw {:?} in:\n{updated_ct}",
            String::from_utf8_lossy(&name)
        );
    }
}

#[test]
fn ensure_workbook_rels_has_relationship_handles_relationship_scoped_prefix_declaration() {
    // Prefix `pr` is declared on the Relationship element, not on the root. Writers must not
    // insert a new `pr:Relationship` sibling without ensuring the prefix is in scope.
    let rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <pr:Relationship xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships" Id="rId1" Type="{REL_TYPE_STYLES}" Target="styles.xml"/>
</Relationships>"#
    );

    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    parts.insert(WORKBOOK_RELS_PART.to_string(), rels.into_bytes());

    ensure_workbook_rels_has_relationship(&mut parts, REL_TYPE_SHARED_STRINGS, "sharedStrings.xml")
        .expect("patch workbook rels");

    let out = String::from_utf8(
        parts
            .get(WORKBOOK_RELS_PART)
            .expect("workbook rels present")
            .clone(),
    )
    .expect("utf8");
    assert!(
        out.contains(REL_TYPE_SHARED_STRINGS),
        "expected sharedStrings relationship to be inserted"
    );
    Document::parse(&out).expect("output rels should be valid xml");
}

#[test]
fn patch_workbook_rels_for_sheet_edits_handles_relationship_scoped_prefix_declaration() {
    // Prefix `pr` is declared on the Relationship element, not on the root. Writers must not
    // insert a new `pr:Relationship` sibling without ensuring the prefix is in scope.
    let rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <pr:Relationship xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships" Id="rId1" Type="{WORKSHEET_REL_TYPE}" Target="worksheets/sheet1.xml"/>
</Relationships>"#
    );

    let mut parts: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    parts.insert(WORKBOOK_RELS_PART.to_string(), rels.into_bytes());

    let removed = vec![SheetMeta {
        worksheet_id: 1,
        sheet_id: 1,
        relationship_id: "rId1".to_string(),
        state: None,
        path: "xl/worksheets/sheet1.xml".to_string(),
    }];
    let added = vec![SheetMeta {
        worksheet_id: 2,
        sheet_id: 2,
        relationship_id: "rId2".to_string(),
        state: None,
        path: "xl/worksheets/sheet2.xml".to_string(),
    }];

    patch_workbook_rels_for_sheet_edits(&mut parts, &removed, &added).expect("patch rels");

    let out = String::from_utf8(
        parts
            .get(WORKBOOK_RELS_PART)
            .expect("workbook rels present")
            .clone(),
    )
    .expect("utf8");
    Document::parse(&out).expect("output rels should be valid xml");
    assert!(
        out.contains(r#"Id="rId2""#),
        "expected added sheet relationship to be inserted"
    );
    assert!(
        !out.contains(r#"Id="rId1""#),
        "expected removed sheet relationship to be dropped"
    );
}
