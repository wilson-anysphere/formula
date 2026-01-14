use std::io::Write;

use formula_model::ArgbColor;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn build_xlsx(include_theme_relationship: bool) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = if include_theme_relationship {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"#
    } else {
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#
    };

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    // Non-default accent1 (0xFF112233).
    let theme_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom Theme">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="112233"/></a:accent1>
    </a:clrScheme>
  </a:themeElements>
</a:theme>"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let cursor = std::io::Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

    writer
        .start_file("[Content_Types].xml", options)
        .expect("start content types");
    writer
        .write_all(content_types.as_bytes())
        .expect("write content types");

    writer
        .start_file("_rels/.rels", options)
        .expect("start root rels");
    writer
        .write_all(root_rels.as_bytes())
        .expect("write root rels");

    writer
        .start_file("xl/workbook.xml", options)
        .expect("start workbook.xml");
    writer
        .write_all(workbook_xml.as_bytes())
        .expect("write workbook.xml");

    writer
        .start_file("xl/_rels/workbook.xml.rels", options)
        .expect("start workbook rels");
    writer
        .write_all(workbook_rels.as_bytes())
        .expect("write workbook rels");

    writer
        .start_file("xl/worksheets/sheet1.xml", options)
        .expect("start sheet");
    writer.write_all(sheet_xml.as_bytes()).expect("write sheet");

    writer
        .start_file("xl/theme/theme1.xml", options)
        .expect("start theme");
    writer.write_all(theme_xml.as_bytes()).expect("write theme");

    writer.finish().expect("finish zip").into_inner()
}

#[test]
fn reads_theme_palette_from_workbook_relationship() {
    let bytes = build_xlsx(true);

    let workbook =
        formula_xlsx::read_workbook_model_from_bytes(&bytes).expect("read workbook model");
    assert_eq!(workbook.theme.accent1, ArgbColor(0xFF112233));

    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load workbook");
    assert_eq!(doc.workbook.theme.accent1, ArgbColor(0xFF112233));
}

#[test]
fn reads_theme_palette_without_workbook_relationship() {
    let bytes = build_xlsx(false);

    let workbook =
        formula_xlsx::read_workbook_model_from_bytes(&bytes).expect("read workbook model");
    assert_eq!(workbook.theme.accent1, ArgbColor(0xFF112233));

    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load workbook");
    assert_eq!(doc.workbook.theme.accent1, ArgbColor(0xFF112233));
}
