use std::io::{Cursor, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{CellPatch, PackageCellPatch, WorkbookCellPatches, XlsxPackage};

fn build_cellimages_calcchain_fixture() -> (Vec<u8>, Vec<u8>, Vec<u8>, Vec<u8>) {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"></Override>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
    Target="calcChain.xml"></Relationship>
  <Relationship Id="rId3"
    Type="http://schemas.microsoft.com/office/2020/relationships/cellimages"
    Target="cellimages.xml"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1">
        <f>1+1</f>
        <v>2</v>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

    let calc_chain_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    // Keep these parts byte-for-byte stable (the regression this test guards against).
    let cellimages_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/9/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:embed="rId1"/>
</cellImages>"#
        .to_vec();

    let cellimages_rels = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="media/image1.png"/>
</Relationships>"#
        .to_vec();

    let image1_png = b"not-a-real-png".to_vec();

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    fn add_file(
        zip: &mut zip::ZipWriter<Cursor<Vec<u8>>>,
        options: zip::write::FileOptions<()>,
        name: &str,
        bytes: &[u8],
    ) {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    add_file(&mut zip, options, "[Content_Types].xml", content_types.as_bytes());
    add_file(&mut zip, options, "_rels/.rels", root_rels.as_bytes());

    add_file(&mut zip, options, "xl/workbook.xml", workbook_xml.as_bytes());
    add_file(
        &mut zip,
        options,
        "xl/_rels/workbook.xml.rels",
        workbook_rels.as_bytes(),
    );
    add_file(&mut zip, options, "xl/worksheets/sheet1.xml", worksheet_xml.as_bytes());

    add_file(&mut zip, options, "xl/calcChain.xml", calc_chain_xml);

    add_file(&mut zip, options, "xl/cellimages.xml", &cellimages_xml);
    add_file(
        &mut zip,
        options,
        "xl/_rels/cellimages.xml.rels",
        &cellimages_rels,
    );
    add_file(&mut zip, options, "xl/media/image1.png", &image1_png);

    (zip.finish().unwrap().into_inner(), cellimages_xml, cellimages_rels, image1_png)
}

fn workbook_relationships(rels_xml: &str) -> Vec<(String, String)> {
    let doc = roxmltree::Document::parse(rels_xml).expect("parse workbook.xml.rels");
    doc.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Relationship")
        .map(|n| {
            (
                n.attribute("Type").unwrap_or_default().to_string(),
                n.attribute("Target").unwrap_or_default().to_string(),
            )
        })
        .collect()
}

fn content_type_overrides(ct_xml: &str) -> Vec<String> {
    let doc = roxmltree::Document::parse(ct_xml).expect("parse [Content_Types].xml");
    doc.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Override")
        .filter_map(|n| n.attribute("PartName").map(ToString::to_string))
        .collect()
}

fn content_type_for_part(ct_xml: &str, part_name: &str) -> Option<String> {
    let doc = roxmltree::Document::parse(ct_xml).expect("parse [Content_Types].xml");
    doc.descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Override"
                && n.attribute("PartName") == Some(part_name)
        })
        .and_then(|n| n.attribute("ContentType").map(ToString::to_string))
}

#[test]
fn formula_patches_drop_calc_chain_but_preserve_cellimages_parts() -> Result<(), Box<dyn std::error::Error>>
{
    let (fixture, cellimages_xml, cellimages_rels, image1_png) =
        build_cellimages_calcchain_fixture();

    let mut pkg = XlsxPackage::from_bytes(&fixture)?;
    assert!(
        pkg.part("xl/calcChain.xml").is_some(),
        "fixture must include xl/calcChain.xml"
    );
    assert!(
        pkg.part("xl/cellimages.xml").is_some(),
        "fixture must include xl/cellimages.xml"
    );

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        // Update the formula so `RecalcPolicy` drops calcChain.
        CellPatch::set_value_with_formula(CellValue::Number(3.0), " =1+2"),
    );
    pkg.apply_cell_patches(&patches)?;

    let out_bytes = pkg.write_to_bytes()?;
    let out_pkg = XlsxPackage::from_bytes(&out_bytes)?;

    // calcChain.xml should be removed entirely.
    assert!(out_pkg.part("xl/calcChain.xml").is_none());

    // workbook.xml.rels should remove the calcChain relationship but preserve the cellimages one.
    let workbook_rels_xml =
        std::str::from_utf8(out_pkg.part("xl/_rels/workbook.xml.rels").unwrap())?;
    let rels = workbook_relationships(workbook_rels_xml);
    assert!(
        !rels.iter()
            .any(|(ty, target)| ty.contains("relationships/calcChain") || target.ends_with("calcChain.xml")),
        "expected workbook.xml.rels to remove calcChain relationship (got {workbook_rels_xml:?})"
    );
    assert!(
        rels.iter()
            .any(|(_, target)| target.ends_with("cellimages.xml")),
        "expected workbook.xml.rels to preserve cellimages relationship (got {workbook_rels_xml:?})"
    );

    // [Content_Types].xml should remove the calcChain override but preserve the cellimages one.
    let ct_xml = std::str::from_utf8(out_pkg.part("[Content_Types].xml").unwrap())?;
    let overrides = content_type_overrides(ct_xml);
    assert!(
        !overrides.iter().any(|p| p.ends_with("calcChain.xml")),
        "expected [Content_Types].xml to remove calcChain override (got {ct_xml:?})"
    );
    assert!(
        overrides.iter().any(|p| p == "/xl/cellimages.xml"),
        "expected [Content_Types].xml to preserve cellimages override (got {ct_xml:?})"
    );
    assert_eq!(
        content_type_for_part(ct_xml, "/xl/cellimages.xml").as_deref(),
        Some("application/vnd.ms-excel.cellimages+xml"),
        "expected [Content_Types].xml to preserve cellimages ContentType string (got {ct_xml:?})"
    );

    // In-cell image parts must be preserved byte-for-byte.
    assert_eq!(out_pkg.part("xl/cellimages.xml").unwrap(), cellimages_xml);
    assert_eq!(
        out_pkg.part("xl/_rels/cellimages.xml.rels").unwrap(),
        cellimages_rels
    );
    assert_eq!(out_pkg.part("xl/media/image1.png").unwrap(), image1_png);

    Ok(())
}

#[test]
fn streaming_formula_patches_drop_calc_chain_but_preserve_cellimages_parts(
) -> Result<(), Box<dyn std::error::Error>> {
    let (fixture, cellimages_xml, cellimages_rels, image1_png) =
        build_cellimages_calcchain_fixture();

    let pkg = XlsxPackage::from_bytes(&fixture)?;
    let patch = PackageCellPatch::for_sheet_name(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellValue::Number(3.0),
        Some("=1+2".to_string()),
    );
    let out_bytes = pkg.apply_cell_patches_to_bytes(&[patch])?;
    let out_pkg = XlsxPackage::from_bytes(&out_bytes)?;

    // calcChain.xml should be removed entirely.
    assert!(out_pkg.part("xl/calcChain.xml").is_none());

    // workbook.xml.rels should remove the calcChain relationship but preserve the cellimages one.
    let workbook_rels_xml =
        std::str::from_utf8(out_pkg.part("xl/_rels/workbook.xml.rels").unwrap())?;
    let rels = workbook_relationships(workbook_rels_xml);
    assert!(
        !rels.iter().any(|(ty, target)| {
            ty.contains("relationships/calcChain") || target.ends_with("calcChain.xml")
        }),
        "expected workbook.xml.rels to remove calcChain relationship (got {workbook_rels_xml:?})"
    );
    assert!(
        rels.iter()
            .any(|(_, target)| target.ends_with("cellimages.xml")),
        "expected workbook.xml.rels to preserve cellimages relationship (got {workbook_rels_xml:?})"
    );

    // [Content_Types].xml should remove the calcChain override but preserve the cellimages one.
    let ct_xml = std::str::from_utf8(out_pkg.part("[Content_Types].xml").unwrap())?;
    let overrides = content_type_overrides(ct_xml);
    assert!(
        !overrides.iter().any(|p| p.ends_with("calcChain.xml")),
        "expected [Content_Types].xml to remove calcChain override (got {ct_xml:?})"
    );
    assert!(
        overrides.iter().any(|p| p == "/xl/cellimages.xml"),
        "expected [Content_Types].xml to preserve cellimages override (got {ct_xml:?})"
    );
    assert_eq!(
        content_type_for_part(ct_xml, "/xl/cellimages.xml").as_deref(),
        Some("application/vnd.ms-excel.cellimages+xml"),
        "expected [Content_Types].xml to preserve cellimages ContentType string (got {ct_xml:?})"
    );

    // In-cell image parts must be preserved byte-for-byte.
    assert_eq!(out_pkg.part("xl/cellimages.xml").unwrap(), cellimages_xml);
    assert_eq!(
        out_pkg.part("xl/_rels/cellimages.xml.rels").unwrap(),
        cellimages_rels
    );
    assert_eq!(out_pkg.part("xl/media/image1.png").unwrap(), image1_png);

    Ok(())
}
