use std::collections::BTreeMap;
use std::io::{Cursor, Write};

use base64::Engine;
use formula_model::{CellRef, CellValue};
use zip::write::FileOptions;
use zip::ZipWriter;

use formula_xlsx::{CellPatch, WorkbookCellPatches, XlsxPackage};

fn build_fixture_xlsx() -> Vec<u8> {
    // 1x1 transparent PNG.
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    let parts: BTreeMap<String, Vec<u8>> = [
        (
            "[Content_Types].xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
</Types>
"#
            .to_vec(),
        ),
        (
            "_rels/.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#
            .to_vec(),
        ),
        (
            "xl/workbook.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#
            .to_vec(),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"#
            .to_vec(),
        ),
        (
            "xl/worksheets/sheet1.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>
"#
            .to_vec(),
        ),
        (
            "xl/cellimages.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<etc:cellImages
 xmlns:etc="http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages"
 xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <etc:cellImage>
    <xdr:pic>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
      </xdr:blipFill>
    </xdr:pic>
  </etc:cellImage>
</etc:cellImages>
"#
            .to_vec(),
        ),
        (
            "xl/_rels/cellimages.xml.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
"#
            .to_vec(),
        ),
        ("xl/media/image1.png".to_string(), png_bytes),
    ]
    .into_iter()
    .collect();

    let cursor = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in parts {
        writer.start_file(name, options).unwrap();
        writer.write_all(&bytes).unwrap();
    }

    let cursor = writer.finish().unwrap();
    cursor.into_inner()
}

#[test]
fn apply_cell_patches_preserves_cellimages_catalog_and_media_bytes() {
    let bytes = build_fixture_xlsx();
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("load package");

    let cellimages_xml = pkg.part("xl/cellimages.xml").unwrap().to_vec();
    let cellimages_rels = pkg.part("xl/_rels/cellimages.xml.rels").unwrap().to_vec();
    let image_bytes = pkg.part("xl/media/image1.png").unwrap().to_vec();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value(CellValue::Number(2.0)),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let updated_sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap())
        .expect("sheet xml is utf8");
    assert!(
        updated_sheet_xml.contains("<v>2</v>"),
        "expected patched sheet to contain updated value, got: {updated_sheet_xml}"
    );

    assert_eq!(
        pkg.part("xl/cellimages.xml").unwrap(),
        cellimages_xml.as_slice(),
        "expected xl/cellimages.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        pkg.part("xl/_rels/cellimages.xml.rels").unwrap(),
        cellimages_rels.as_slice(),
        "expected xl/_rels/cellimages.xml.rels to be preserved byte-for-byte"
    );
    assert_eq!(
        pkg.part("xl/media/image1.png").unwrap(),
        image_bytes.as_slice(),
        "expected xl/media/image1.png to be preserved byte-for-byte"
    );
}

