use std::io::{Cursor, Write};

use formula_model::rich_text::{RichText, RichTextRun, RichTextRunStyle};
use formula_model::{CellRef, CellValue};
use formula_xlsx::{CellPatch, WorkbookCellPatches, XlsxPackage};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn build_fixture_xlsx(worksheet_xml: &str) -> Vec<u8> {
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

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

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

    zip.finish().unwrap().into_inner()
}

#[test]
fn in_memory_patch_preserves_unknown_inline_rich_rpr_tags_when_patch_value_omits_empty_runs(
) -> Result<(), Box<dyn std::error::Error>> {
    // Inline rich text cell with an unsupported `<strike/>` tag in `<rPr>`.
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <r>
            <rPr><b/><strike/></rPr>
            <t>Bold</t>
          </r>
          <r>
            <t>Plain</t>
          </r>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

    let bytes = build_fixture_xlsx(worksheet_xml);

    // Patch with a RichText value that represents only the styled runs (omits the unstyled run).
    // This should still be semantically equal to the existing inline rich text payload.
    let rich = RichText {
        text: "BoldPlain".to_string(),
        runs: vec![RichTextRun {
            start: 0,
            end: 4,
            style: RichTextRunStyle {
                bold: Some(true),
                ..Default::default()
            },
        }],
        phonetic: None,
    };

    let mut pkg = XlsxPackage::from_bytes(&bytes)?;
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value_with_style(CellValue::RichText(rich), 1),
    );
    pkg.apply_cell_patches(&patches)?;

    let out_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    assert!(
        out_xml.contains(r#"s="1""#),
        "expected patched cell to contain s=\"1\" style attribute:\n{out_xml}"
    );
    assert!(
        out_xml.contains("<strike"),
        "expected patched worksheet XML to preserve unknown <strike> run property:\n{out_xml}"
    );

    Ok(())
}
