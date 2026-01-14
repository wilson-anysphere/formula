use std::io::{Cursor, Read, Write};

use formula_model::Workbook;
use formula_xlsx::{load_from_bytes, read_workbook_model_from_bytes, write_workbook_to_writer};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

const X14AC_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

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

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    fn add_file(
        zip: &mut ZipWriter<Cursor<Vec<u8>>>,
        options: FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(&mut zip, options, "xl/workbook.xml", workbook_xml.as_bytes());
    add_file(
        &mut zip,
        options,
        "xl/_rels/workbook.xml.rels",
        workbook_rels.as_bytes(),
    );
    add_file(
        &mut zip,
        options,
        "xl/worksheets/sheet1.xml",
        sheet1_xml.as_bytes(),
    );

    zip.finish().unwrap().into_inner()
}

fn zip_part(zip_bytes: &[u8], name: &str) -> String {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = String::new();
    file.read_to_string(&mut buf).expect("read part");
    buf
}

#[test]
fn reads_sheet_format_pr_defaults_into_model() {
    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14ac="{X14AC_NS}">
  <sheetFormatPr defaultRowHeight="20" defaultColWidth="9.5" baseColWidth="8" x14ac:dyDescent="0.25"/>
  <sheetData/>
</worksheet>"#
    );
    let bytes = build_minimal_xlsx_with_sheet1(&sheet_xml);

    let workbook = read_workbook_model_from_bytes(&bytes).expect("fast reader");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].default_row_height, Some(20.0));
    assert_eq!(workbook.sheets[0].default_col_width, Some(9.5));

    let doc = load_from_bytes(&bytes).expect("load_from_bytes");
    assert_eq!(doc.workbook.sheets.len(), 1);
    assert_eq!(doc.workbook.sheets[0].default_row_height, Some(20.0));
    assert_eq!(doc.workbook.sheets[0].default_col_width, Some(9.5));
}

#[test]
fn semantic_export_emits_sheet_format_pr_defaults() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1".to_string()).unwrap();
    let sheet = workbook.sheet_mut(sheet_id).unwrap();
    sheet.default_row_height = Some(20.0);
    sheet.default_col_width = Some(9.5);

    let mut cursor = Cursor::new(Vec::new());
    write_workbook_to_writer(&workbook, &mut cursor).expect("write workbook");
    let bytes = cursor.into_inner();

    let sheet_xml = zip_part(&bytes, "xl/worksheets/sheet1.xml");
    let doc = roxmltree::Document::parse(&sheet_xml).expect("parse sheet xml");
    let sheet_format = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "sheetFormatPr")
        .expect("expected sheetFormatPr element");

    let row_ht: f32 = sheet_format
        .attribute("defaultRowHeight")
        .expect("defaultRowHeight")
        .parse()
        .expect("parse defaultRowHeight");
    let col_w: f32 = sheet_format
        .attribute("defaultColWidth")
        .expect("defaultColWidth")
        .parse()
        .expect("parse defaultColWidth");

    assert_eq!(row_ht, 20.0);
    assert_eq!(col_w, 9.5);
}

#[test]
fn roundtrip_patching_preserves_sheet_format_pr_and_unknown_attrs() {
    let sheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14ac="{X14AC_NS}">
  <dimension ref="A1"/>
  <sheetFormatPr defaultRowHeight="20" defaultColWidth="9.5" baseColWidth="8" x14ac:dyDescent="0.25"/>
  <sheetData/>
</worksheet>"#
    );
    let bytes = build_minimal_xlsx_with_sheet1(&sheet_xml);
    let mut doc = load_from_bytes(&bytes).expect("load xlsx");
    let sheet_id = doc.workbook.sheets[0].id;

    // Force a real patch to sheetFormatPr: update only the defaultRowHeight.
    let sheet = doc.workbook.sheet_mut(sheet_id).unwrap();
    sheet.default_row_height = Some(21.0);

    let saved = doc.save_to_vec().expect("save");
    let out_xml = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let parsed = roxmltree::Document::parse(&out_xml).expect("parse sheet xml");
    let sheet_format = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "sheetFormatPr")
        .expect("sheetFormatPr element");

    let row_ht: f32 = sheet_format
        .attribute("defaultRowHeight")
        .expect("defaultRowHeight")
        .parse()
        .expect("parse row ht");
    let col_w: f32 = sheet_format
        .attribute("defaultColWidth")
        .expect("defaultColWidth")
        .parse()
        .expect("parse col width");
    assert_eq!(row_ht, 21.0);
    assert_eq!(col_w, 9.5);
    assert_eq!(
        sheet_format.attribute("baseColWidth"),
        Some("8"),
        "expected baseColWidth to be preserved"
    );
    let dy_descent = sheet_format
        .attribute((X14AC_NS, "dyDescent"))
        .or_else(|| sheet_format.attribute("x14ac:dyDescent"));
    assert_eq!(
        dy_descent,
        Some("0.25"),
        "expected unknown x14ac:dyDescent to be preserved"
    );

    // If the model does not specify defaults, the existing sheetFormatPr should be preserved.
    let mut doc = load_from_bytes(&bytes).expect("reload xlsx");
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).unwrap();
    sheet.default_row_height = None;
    sheet.default_col_width = None;

    let saved = doc.save_to_vec().expect("save without defaults");
    let out_xml = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let parsed = roxmltree::Document::parse(&out_xml).expect("parse sheet xml");
    let sheet_format = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "sheetFormatPr")
        .expect("sheetFormatPr element");

    let row_ht: f32 = sheet_format
        .attribute("defaultRowHeight")
        .expect("defaultRowHeight")
        .parse()
        .expect("parse row ht");
    let col_w: f32 = sheet_format
        .attribute("defaultColWidth")
        .expect("defaultColWidth")
        .parse()
        .expect("parse col width");
    assert_eq!(row_ht, 20.0);
    assert_eq!(col_w, 9.5);
    assert_eq!(
        sheet_format.attribute("baseColWidth"),
        Some("8"),
        "expected baseColWidth to be preserved"
    );
    let dy_descent = sheet_format
        .attribute((X14AC_NS, "dyDescent"))
        .or_else(|| sheet_format.attribute("x14ac:dyDescent"));
    assert_eq!(
        dy_descent,
        Some("0.25"),
        "expected unknown x14ac:dyDescent to be preserved"
    );
}

