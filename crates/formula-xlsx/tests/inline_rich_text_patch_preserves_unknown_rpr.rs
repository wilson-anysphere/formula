use std::io::{Cursor, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{load_from_bytes, CellPatch, WorkbookCellPatches, XlsxPackage};
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
fn in_memory_patch_preserves_unknown_inline_rich_rpr_tags_on_style_only_patch(
) -> Result<(), Box<dyn std::error::Error>> {
    // Include an unsupported `<strike/>` tag in the run properties to ensure we don't rewrite the
    // inline rich text payload when only updating the cell's style index.
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

    // Load to obtain the model cell value (now parsed as RichText because of the `<b/>` run style).
    let doc = load_from_bytes(&bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).expect("sheet exists");
    let cell_ref = CellRef::from_a1("A1")?;
    let value = sheet.value(cell_ref).clone();
    match &value {
        CellValue::RichText(rich) => assert_eq!(rich.text, "BoldPlain"),
        other => panic!("expected RichText value, got {other:?}"),
    }

    // Apply a style-only patch using the RichText value. The patcher should treat the existing
    // inline rich text as semantically equal and preserve the original `<is>` subtree, including
    // the unsupported `<strike/>` tag.
    let mut pkg = XlsxPackage::from_bytes(&bytes)?;
    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        cell_ref,
        CellPatch::set_value_with_style(value, 1),
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

