use std::io::{Cursor, Read, Write};

use formula_xlsx::load_from_bytes;
use zip::ZipArchive;

const STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="0.00"/>
  </numFmts>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"#;

// Like STYLES_XML, but with the custom format at xf index 0.
const STYLES_XML_XF0_CUSTOM: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="0.00"/>
  </numFmts>
  <cellXfs count="2">
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>"#;

fn build_minimal_xlsx(sheet_xml: &str, styles_xml: &str) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options =
        zip::write::FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/styles.xml", options).unwrap();
    zip.write_all(styles_xml.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

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

#[test]
fn save_to_vec_preserves_col_style_when_outline_removed() -> Result<(), Box<dyn std::error::Error>> {
    // Column B carries both a default style and an outlineLevel. Clearing the outline should not
    // drop the style attributes when the writer rewrites the <cols> section.
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="2" max="2" style="1" customFormat="1" outlineLevel="1"/>
  </cols>
  <sheetData/>
</worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml, STYLES_XML);
    let mut doc = load_from_bytes(&bytes)?;

    let sheet_id = doc.workbook.sheets[0].id;
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .ungroup_cols(2, 2);

    let saved = doc.save_to_vec()?;
    let xml_bytes = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let xml = std::str::from_utf8(&xml_bytes)?;
    let parsed = roxmltree::Document::parse(xml)?;

    let col_b = parsed
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "col"
                && n.attribute("min") == Some("2")
                && n.attribute("max") == Some("2")
        })
        .expect("col B exists");
    assert_eq!(col_b.attribute("outlineLevel"), None);
    assert_eq!(col_b.attribute("style"), Some("1"));
    assert_eq!(col_b.attribute("customFormat"), Some("1"));

    Ok(())
}

#[test]
fn noop_save_preserves_cols_when_style_xf_is_zero() -> Result<(), Box<dyn std::error::Error>> {
    // Some producers place a custom xf at index 0. Ensure we treat `style="0"` as a real style
    // reference so no-op saves don't rewrite <cols> (which would reorder attributes).
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="2" max="2" customFormat="1" style="0"/>
  </cols>
  <sheetData/>
</worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml, STYLES_XML_XF0_CUSTOM);
    let doc = load_from_bytes(&bytes)?;
    let saved = doc.save_to_vec()?;

    let orig_sheet = zip_part(&bytes, "xl/worksheets/sheet1.xml");
    let saved_sheet = zip_part(&saved, "xl/worksheets/sheet1.xml");
    assert_eq!(
        saved_sheet, orig_sheet,
        "expected no-op save to preserve sheet XML bytes"
    );

    Ok(())
}

#[test]
fn noop_save_preserves_cols_when_style_is_zero_default() -> Result<(), Box<dyn std::error::Error>> {
    // Excel's default xf is typically stored at index 0. Some producers still emit
    // `customFormat="1" style="0"` in `<col>` elements even though it has no semantic effect.
    //
    // Ensure we treat that as equivalent to "no style override" so no-op saves preserve the
    // original sheet XML bytes.
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="2" max="2" customFormat="1" style="0"/>
  </cols>
  <sheetData/>
</worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml, STYLES_XML);
    let doc = load_from_bytes(&bytes)?;
    let saved = doc.save_to_vec()?;

    let orig_sheet = zip_part(&bytes, "xl/worksheets/sheet1.xml");
    let saved_sheet = zip_part(&saved, "xl/worksheets/sheet1.xml");
    assert_eq!(
        saved_sheet, orig_sheet,
        "expected no-op save to preserve sheet XML bytes"
    );

    Ok(())
}
