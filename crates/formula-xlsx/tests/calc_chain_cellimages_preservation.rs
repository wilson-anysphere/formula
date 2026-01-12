use std::collections::BTreeMap;
use std::io::{Cursor, Write};

use base64::Engine as _;
use formula_model::{CellRef, CellValue};
use formula_xlsx::{CellPatch, WorkbookCellPatches, XlsxPackage};
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_fixture_xlsx() -> Vec<u8> {
    // 1x1 transparent PNG.
    let png_bytes = base64::engine::general_purpose::STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/58HAQUBAO3+2NoAAAAASUVORK5CYII=")
        .expect("valid base64 png");

    // Keep the cellimages content type string stable so the test can verify we preserve it.
    let cellimages_content_type = "application/vnd.ms-excel.cellimages+xml";

    let parts: BTreeMap<String, Vec<u8>> = [
        (
            "[Content_Types].xml".to_string(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="{cellimages_content_type}"/>
</Types>
"#
            )
            .into_bytes(),
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/10/relationships/cellimages" Target="cellimages.xml"/>
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
            "xl/calcChain.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
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
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
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

    writer.finish().unwrap().into_inner()
}

#[test]
fn formula_edits_drop_calc_chain_without_stripping_cellimages_metadata() {
    let bytes = build_fixture_xlsx();
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("load package");

    // Capture cellimages parts for byte-for-byte preservation checks.
    let cellimages_xml = pkg.part("xl/cellimages.xml").unwrap().to_vec();
    let cellimages_rels = pkg.part("xl/_rels/cellimages.xml.rels").unwrap().to_vec();
    let image_bytes = pkg.part("xl/media/image1.png").unwrap().to_vec();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value_with_formula(CellValue::Number(2.0), "=SUM(1,1)"),
    );
    pkg.apply_cell_patches(&patches)
        .expect("apply_cell_patches should succeed");

    // Assert: calc chain part is removed.
    assert!(
        pkg.part("xl/calcChain.xml").is_none(),
        "expected formula edit to remove xl/calcChain.xml"
    );

    // Assert: workbook rels no longer reference calcChain.xml, but still reference cellimages.xml.
    let workbook_rels_xml = std::str::from_utf8(pkg.part("xl/_rels/workbook.xml.rels").unwrap())
        .expect("workbook.xml.rels is utf-8");
    let rels_doc = roxmltree::Document::parse(workbook_rels_xml).expect("parse workbook.xml.rels");
    let rel_targets: Vec<&str> = rels_doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
        .filter_map(|n| n.attribute("Target"))
        .collect();

    assert!(
        !rel_targets.iter().any(|t| t.ends_with("calcChain.xml")),
        "expected workbook.xml.rels to drop calcChain relationship, but found targets: {rel_targets:?}\nworkbook.xml.rels:\n{workbook_rels_xml}"
    );
    assert!(
        rel_targets.iter().any(|t| t.ends_with("cellimages.xml")),
        "expected workbook.xml.rels to preserve cellimages relationship, but found targets: {rel_targets:?}\nworkbook.xml.rels:\n{workbook_rels_xml}"
    );

    // Assert: [Content_Types].xml no longer references calcChain.xml but preserves cellimages.xml override.
    let content_types_xml = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap())
        .expect("[Content_Types].xml is utf-8");
    let ct_doc = roxmltree::Document::parse(content_types_xml).expect("parse [Content_Types].xml");

    let override_part_names: Vec<&str> = ct_doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Override")
        .filter_map(|n| n.attribute("PartName"))
        .collect();
    assert!(
        !override_part_names
            .iter()
            .any(|part| part.ends_with("calcChain.xml")),
        "expected [Content_Types].xml to drop calcChain override, but found PartNames: {override_part_names:?}\n[Content_Types].xml:\n{content_types_xml}"
    );

    let cellimages_override = ct_doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Override"
                && n.attribute("PartName") == Some("/xl/cellimages.xml")
        })
        .expect("expected [Content_Types].xml to preserve /xl/cellimages.xml override");
    assert_eq!(
        cellimages_override.attribute("ContentType"),
        Some("application/vnd.ms-excel.cellimages+xml"),
        "expected [Content_Types].xml to preserve cellimages ContentType string\n[Content_Types].xml:\n{content_types_xml}"
    );

    // Assert: cellimages parts are preserved byte-for-byte.
    assert_eq!(
        pkg.part("xl/cellimages.xml").unwrap(),
        cellimages_xml.as_slice(),
        "expected xl/cellimages.xml to be preserved byte-for-byte during calcChain cleanup"
    );
    assert_eq!(
        pkg.part("xl/_rels/cellimages.xml.rels").unwrap(),
        cellimages_rels.as_slice(),
        "expected xl/_rels/cellimages.xml.rels to be preserved byte-for-byte during calcChain cleanup"
    );
    assert_eq!(
        pkg.part("xl/media/image1.png").unwrap(),
        image_bytes.as_slice(),
        "expected xl/media/image1.png to be preserved byte-for-byte during calcChain cleanup"
    );
}

