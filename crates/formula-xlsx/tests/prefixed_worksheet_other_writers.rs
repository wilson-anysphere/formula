use std::io::{Cursor, Write};

use formula_model::{CellRef, Hyperlink, HyperlinkTarget, Range, TabColor};
use formula_xlsx::XlsxPackage;

fn build_minimal_xlsx(sheet_xml: &str) -> Vec<u8> {
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
</Relationships>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn tab_color_insertion_preserves_worksheet_prefix() -> Result<(), Box<dyn std::error::Error>> {
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetData/>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let color = TabColor::rgb("FF00FF00");
    pkg.set_worksheet_tab_color("xl/worksheets/sheet1.xml", Some(&color))?;

    let updated = std::str::from_utf8(
        pkg.part("xl/worksheets/sheet1.xml")
            .expect("worksheet part exists"),
    )?;

    roxmltree::Document::parse(updated)?;
    assert!(
        updated.contains("<x:sheetPr") && updated.contains("</x:sheetPr>"),
        "expected inserted <x:sheetPr> with matching end tag, got:\n{updated}"
    );
    assert!(
        updated.contains("<x:tabColor"),
        "expected inserted <x:tabColor>, got:\n{updated}"
    );
    assert!(
        !updated.contains("<sheetPr"),
        "should not introduce an unprefixed <sheetPr> element"
    );
    assert!(
        !updated.contains("<tabColor"),
        "should not introduce an unprefixed <tabColor> element"
    );

    Ok(())
}

#[test]
fn tab_color_replacement_preserves_worksheet_prefix() -> Result<(), Box<dyn std::error::Error>> {
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetPr>
    <x:tabColor rgb="FFFF0000"/>
  </x:sheetPr>
  <x:sheetData/>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml);
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;

    let color = TabColor::rgb("FF00FF00");
    pkg.set_worksheet_tab_color("xl/worksheets/sheet1.xml", Some(&color))?;

    let updated = std::str::from_utf8(
        pkg.part("xl/worksheets/sheet1.xml")
            .expect("worksheet part exists"),
    )?;

    roxmltree::Document::parse(updated)?;
    assert!(
        updated.contains("<x:tabColor rgb=\"FF00FF00\""),
        "expected prefixed tabColor to update, got:\n{updated}"
    );
    assert!(
        !updated.contains("<tabColor"),
        "should not introduce an unprefixed <tabColor> element"
    );

    Ok(())
}

#[test]
fn hyperlinks_insertion_preserves_worksheet_prefix() -> Result<(), Box<dyn std::error::Error>> {
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetData/>
</x:worksheet>"#;

    let links = vec![Hyperlink {
        range: Range::from_a1("A1")?,
        target: HyperlinkTarget::Internal {
            sheet: "Sheet1".to_string(),
            cell: CellRef::new(0, 0),
        },
        display: Some("Go".to_string()),
        tooltip: None,
        rel_id: None,
    }];

    let updated = formula_xlsx::hyperlinks::update_worksheet_xml(sheet_xml, &links)?;

    roxmltree::Document::parse(&updated)?;
    assert!(
        updated.contains("<x:hyperlinks") && updated.contains("</x:hyperlinks>"),
        "expected inserted <x:hyperlinks> block, got:\n{updated}"
    );
    assert!(
        updated.contains("<x:hyperlink"),
        "expected inserted <x:hyperlink>, got:\n{updated}"
    );
    assert!(
        !updated.contains("<hyperlinks"),
        "should not introduce an unprefixed <hyperlinks> element"
    );
    assert!(
        !updated.contains("<hyperlink"),
        "should not introduce an unprefixed <hyperlink> element"
    );

    Ok(())
}

