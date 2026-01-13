use std::io::{Cursor, Write};

use formula_model::ArgbColor;
use formula_xlsx::{load_from_bytes, read_workbook_model_from_bytes};
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_minimal_xlsx_with_custom_theme() -> Vec<u8> {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let theme_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom Theme">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="111111"/></a:dk1>
      <a:lt1><a:srgbClr val="EEEEEE"/></a:lt1>
      <a:dk2><a:srgbClr val="222222"/></a:dk2>
      <a:lt2><a:srgbClr val="DDDDDD"/></a:lt2>
      <a:accent1><a:srgbClr val="010203"/></a:accent1>
      <a:accent2><a:srgbClr val="040506"/></a:accent2>
      <a:accent3><a:srgbClr val="070809"/></a:accent3>
      <a:accent4><a:srgbClr val="0A0B0C"/></a:accent4>
      <a:accent5><a:srgbClr val="0D0E0F"/></a:accent5>
      <a:accent6><a:srgbClr val="101112"/></a:accent6>
      <a:hlink><a:srgbClr val="131415"/></a:hlink>
      <a:folHlink><a:srgbClr val="161718"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options =
        FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    fn add_file(
        zip: &mut ZipWriter<Cursor<Vec<u8>>>,
        options: FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(&mut zip, options, "xl/workbook.xml", workbook_xml);
    add_file(
        &mut zip,
        options,
        "xl/_rels/workbook.xml.rels",
        workbook_rels,
    );
    add_file(&mut zip, options, "xl/worksheets/sheet1.xml", worksheet_xml);
    add_file(&mut zip, options, "xl/theme/theme1.xml", theme_xml);

    zip.finish().unwrap().into_inner()
}

fn assert_custom_palette(theme: &formula_model::ThemePalette) {
    assert_eq!(theme.dk1, ArgbColor(0xFF111111));
    assert_eq!(theme.lt1, ArgbColor(0xFFEEEEEE));
    assert_eq!(theme.dk2, ArgbColor(0xFF222222));
    assert_eq!(theme.lt2, ArgbColor(0xFFDDDDDD));
    assert_eq!(theme.accent1, ArgbColor(0xFF010203));
    assert_eq!(theme.accent2, ArgbColor(0xFF040506));
    assert_eq!(theme.accent3, ArgbColor(0xFF070809));
    assert_eq!(theme.accent4, ArgbColor(0xFF0A0B0C));
    assert_eq!(theme.accent5, ArgbColor(0xFF0D0E0F));
    assert_eq!(theme.accent6, ArgbColor(0xFF101112));
    assert_eq!(theme.hlink, ArgbColor(0xFF131415));
    assert_eq!(theme.fol_hlink, ArgbColor(0xFF161718));
}

#[test]
fn load_from_bytes_populates_workbook_theme_palette() {
    let bytes = build_minimal_xlsx_with_custom_theme();
    let doc = load_from_bytes(&bytes).expect("load xlsx document");
    assert_custom_palette(&doc.workbook.theme);
}

#[test]
fn read_workbook_model_from_bytes_populates_workbook_theme_palette() {
    let bytes = build_minimal_xlsx_with_custom_theme();
    let workbook = read_workbook_model_from_bytes(&bytes).expect("read workbook model");
    assert_custom_palette(&workbook.theme);
}

