use std::io::{Cursor, Write};

use formula_model::ArgbColor;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

enum ThemeCase {
    /// `workbook.xml.rels` points at a non-default theme part. The reader must
    /// respect the relationship (and not blindly fall back to `xl/theme/theme1.xml`).
    RelationshipToCustomPart,
    /// No theme relationship; readers should fall back to `xl/theme/theme1.xml`.
    NoThemeRelationship,
}

fn theme_xml(accent1: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom Theme">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="{accent1}"/></a:accent1>
    </a:clrScheme>
  </a:themeElements>
</a:theme>"#,
    )
}

fn build_xlsx(case: ThemeCase) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = match case {
        ThemeCase::RelationshipToCustomPart => {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/customTheme.xml"/>
</Relationships>"#
        }
        ThemeCase::NoThemeRelationship => {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#
        }
    };

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    // Desired accent1 (0xFF112233).
    let expected_theme_xml = theme_xml("112233");
    // Different accent1 so we can prove the relationship wins when present.
    let fallback_theme_xml = theme_xml("445566");

    let content_types = match case {
        ThemeCase::RelationshipToCustomPart => {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/theme/customTheme.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"#
        }
        ThemeCase::NoThemeRelationship => {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"#
        }
    };

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
    let theme1_xml = match case {
        ThemeCase::RelationshipToCustomPart => fallback_theme_xml.as_bytes(),
        ThemeCase::NoThemeRelationship => expected_theme_xml.as_bytes(),
    };
    writer.write_all(theme1_xml).expect("write theme");

    if matches!(case, ThemeCase::RelationshipToCustomPart) {
        writer
            .start_file("xl/theme/customTheme.xml", options)
            .expect("start custom theme");
        writer
            .write_all(expected_theme_xml.as_bytes())
            .expect("write custom theme");
    }

    writer.finish().expect("finish zip").into_inner()
}

#[test]
fn reads_theme_palette_from_workbook_relationship() {
    let bytes = build_xlsx(ThemeCase::RelationshipToCustomPart);

    let workbook =
        formula_xlsx::read_workbook_model_from_bytes(&bytes).expect("read workbook model");
    assert_eq!(workbook.theme.accent1, ArgbColor(0xFF112233));

    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load workbook");
    assert_eq!(doc.workbook.theme.accent1, ArgbColor(0xFF112233));

    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("parse package");
    let pkg_theme = pkg
        .theme_palette()
        .expect("theme palette")
        .expect("theme palette present");
    assert_eq!(pkg_theme.accent1, 0xFF112233);

    let theme = formula_xlsx::theme_palette_from_reader(Cursor::new(bytes))
        .expect("read theme palette")
        .expect("theme palette present");
    assert_eq!(theme.accent1, 0xFF112233);
}

#[test]
fn reads_theme_palette_without_workbook_relationship() {
    let bytes = build_xlsx(ThemeCase::NoThemeRelationship);

    let workbook =
        formula_xlsx::read_workbook_model_from_bytes(&bytes).expect("read workbook model");
    assert_eq!(workbook.theme.accent1, ArgbColor(0xFF112233));

    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load workbook");
    assert_eq!(doc.workbook.theme.accent1, ArgbColor(0xFF112233));

    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("parse package");
    let pkg_theme = pkg
        .theme_palette()
        .expect("theme palette")
        .expect("theme palette present");
    assert_eq!(pkg_theme.accent1, 0xFF112233);

    let theme = formula_xlsx::theme_palette_from_reader(Cursor::new(bytes))
        .expect("read theme palette")
        .expect("theme palette present");
    assert_eq!(theme.accent1, 0xFF112233);
}
