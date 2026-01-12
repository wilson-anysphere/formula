use std::io::{Cursor, Read, Write};

use formula_model::{WorkbookWindow, WorkbookWindowState};
use formula_xlsx::{load_from_bytes, read_workbook_model_from_bytes};
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

fn build_fixture_xlsx() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews>
    <workbookView activeTab="0"/>
  </bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
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
</Types>
"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", root_rels.as_bytes()),
        ("xl/workbook.xml", workbook_xml.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/worksheets/sheet1.xml", sheet_xml.as_bytes()),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn build_fixture_xlsx_without_book_views() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
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
</Types>
"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", root_rels.as_bytes()),
        ("xl/workbook.xml", workbook_xml.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/worksheets/sheet1.xml", sheet_xml.as_bytes()),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn build_fixture_xlsx_empty_book_views() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews/>
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
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
</Types>
"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", root_rels.as_bytes()),
        ("xl/workbook.xml", workbook_xml.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/worksheets/sheet1.xml", sheet_xml.as_bytes()),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

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

fn build_two_sheet_fixture_with_book_views() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews>
    <workbookView activeTab="0" firstSheet="0"/>
  </bookViews>
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
</Types>
"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", root_rels.as_bytes()),
        ("xl/workbook.xml", workbook_xml.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/worksheets/sheet1.xml", sheet_xml.as_bytes()),
        ("xl/worksheets/sheet2.xml", sheet_xml.as_bytes()),
    ] {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn patch_workbook_xml_inserts_workbook_view_window_metadata() {
    let fixture = build_fixture_xlsx();
    let mut doc = load_from_bytes(&fixture).expect("load fixture");

    doc.workbook.view.window = Some(WorkbookWindow {
        x: Some(10),
        y: Some(20),
        width: Some(800),
        height: Some(600),
        state: Some(WorkbookWindowState::Maximized),
    });

    let saved = doc.save_to_vec().expect("save");
    let workbook_xml = zip_part(&saved, "xl/workbook.xml");
    let workbook_xml_str = std::str::from_utf8(&workbook_xml).expect("workbook.xml utf-8");
    let parsed = roxmltree::Document::parse(workbook_xml_str).expect("parse workbook.xml");

    let workbook_view = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "workbookView")
        .expect("expected <workbookView>");

    assert_eq!(workbook_view.attribute("xWindow"), Some("10"));
    assert_eq!(workbook_view.attribute("yWindow"), Some("20"));
    assert_eq!(workbook_view.attribute("windowWidth"), Some("800"));
    assert_eq!(workbook_view.attribute("windowHeight"), Some("600"));
    assert_eq!(workbook_view.attribute("windowState"), Some("maximized"));

    // Reader should round-trip the patched window state too.
    let loaded = load_from_bytes(&saved).expect("load saved");
    assert_eq!(
        loaded.workbook.view.window,
        Some(WorkbookWindow {
            x: Some(10),
            y: Some(20),
            width: Some(800),
            height: Some(600),
            state: Some(WorkbookWindowState::Maximized),
        })
    );

    let model = read_workbook_model_from_bytes(&saved).expect("read workbook model");
    assert_eq!(
        model.view.window,
        Some(WorkbookWindow {
            x: Some(10),
            y: Some(20),
            width: Some(800),
            height: Some(600),
            state: Some(WorkbookWindowState::Maximized),
        })
    );
}

#[test]
fn patch_workbook_xml_inserts_book_views_when_missing() {
    let fixture = build_fixture_xlsx_without_book_views();
    let mut doc = load_from_bytes(&fixture).expect("load fixture");

    doc.workbook.view.window = Some(WorkbookWindow {
        x: Some(1),
        y: Some(2),
        width: Some(3),
        height: Some(4),
        state: Some(WorkbookWindowState::Minimized),
    });

    let saved = doc.save_to_vec().expect("save");
    let workbook_xml = zip_part(&saved, "xl/workbook.xml");
    let workbook_xml_str = std::str::from_utf8(&workbook_xml).expect("workbook.xml utf-8");
    let parsed = roxmltree::Document::parse(workbook_xml_str).expect("parse workbook.xml");

    let workbook_view = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "workbookView")
        .expect("expected <workbookView>");

    assert_eq!(workbook_view.attribute("xWindow"), Some("1"));
    assert_eq!(workbook_view.attribute("yWindow"), Some("2"));
    assert_eq!(workbook_view.attribute("windowWidth"), Some("3"));
    assert_eq!(workbook_view.attribute("windowHeight"), Some("4"));
    assert_eq!(workbook_view.attribute("windowState"), Some("minimized"));
}

#[test]
fn patch_workbook_xml_updates_active_tab_from_model() {
    let fixture = build_two_sheet_fixture_with_book_views();
    let mut doc = load_from_bytes(&fixture).expect("load fixture");

    let sheet2_id = doc.workbook.sheet_by_name("Sheet2").expect("Sheet2").id;
    assert!(doc.workbook.set_active_sheet(sheet2_id));

    let saved = doc.save_to_vec().expect("save");
    let workbook_xml = zip_part(&saved, "xl/workbook.xml");
    let workbook_xml_str = std::str::from_utf8(&workbook_xml).expect("workbook.xml utf-8");
    let parsed = roxmltree::Document::parse(workbook_xml_str).expect("parse workbook.xml");

    let workbook_view = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "workbookView")
        .expect("expected <workbookView>");
    assert_eq!(workbook_view.attribute("activeTab"), Some("1"));
}

#[test]
fn patch_workbook_xml_inserts_workbook_view_when_book_views_empty() {
    let fixture = build_fixture_xlsx_empty_book_views();
    let mut doc = load_from_bytes(&fixture).expect("load fixture");

    doc.workbook.view.window = Some(WorkbookWindow {
        x: Some(1),
        y: Some(2),
        width: Some(3),
        height: Some(4),
        state: Some(WorkbookWindowState::Minimized),
    });

    let saved = doc.save_to_vec().expect("save");
    let workbook_xml = zip_part(&saved, "xl/workbook.xml");
    let workbook_xml_str = std::str::from_utf8(&workbook_xml).expect("workbook.xml utf-8");
    let parsed = roxmltree::Document::parse(workbook_xml_str).expect("parse workbook.xml");

    let workbook_view = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "workbookView")
        .expect("expected <workbookView>");
    assert_eq!(workbook_view.attribute("xWindow"), Some("1"));
    assert_eq!(workbook_view.attribute("yWindow"), Some("2"));
    assert_eq!(workbook_view.attribute("windowWidth"), Some("3"));
    assert_eq!(workbook_view.attribute("windowHeight"), Some("4"));
    assert_eq!(workbook_view.attribute("windowState"), Some("minimized"));
}
