use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{PackageCellPatch, XlsxPackage};
use zip::result::ZipError;

fn build_synthetic_richdata_calcchain_fixture() -> (Vec<u8>, Vec<(&'static str, Vec<u8>)>) {
    let content_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
  <Override PartName="/xl/richData/richValue.xml" ContentType="application/xml"/>
  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/xml"/>
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/xml"/>
  <Override PartName="/xl/richData/richValueStructure.xml" ContentType="application/xml"/>
</Types>"#;

    let root_rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rIdSheet1"/>
  </sheets>
</workbook>"#;

    let workbook_rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdSheet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rIdRichData" Type="http://example.com/relationships/richData" Target="richData/richValue.xml"/>
  <Relationship Id="rIdCalcChain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>"#;

    // Intentionally minimal: a single cell A1 and a tight dimension.
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    // Valid-but-minimal style sheet so the package resembles a "real" workbook (not required for
    // the test assertions, but helps ensure the patcher can round-trip core parts).
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>"#;

    let calc_chain_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    // RichData parts: ensure these survive byte-for-byte when recalc policy drops calcChain.
    let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns="http://example.com/richData">RICH-VALUE</richValue>"#
        .to_vec();
    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://example.com/richData">RICH-VALUE-REL</richValueRel>"#
        .to_vec();
    let rich_value_types_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueTypes xmlns="http://example.com/richData">RICH-VALUE-TYPES</richValueTypes>"#
        .to_vec();
    let rich_value_structure_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueStructure xmlns="http://example.com/richData">RICH-VALUE-STRUCTURE</richValueStructure>"#
        .to_vec();

    let rich_parts = vec![
        ("xl/richData/richValue.xml", rich_value_xml.clone()),
        ("xl/richData/richValueRel.xml", rich_value_rel_xml.clone()),
        ("xl/richData/richValueTypes.xml", rich_value_types_xml.clone()),
        (
            "xl/richData/richValueStructure.xml",
            rich_value_structure_xml.clone(),
        ),
    ];

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types_xml.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels_xml.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels_xml.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/styles.xml", options).unwrap();
    zip.write_all(styles_xml.as_bytes()).unwrap();

    zip.start_file("xl/calcChain.xml", options).unwrap();
    zip.write_all(calc_chain_xml.as_bytes()).unwrap();

    for (name, bytes) in &rich_parts {
        zip.start_file(name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    (zip.finish().unwrap().into_inner(), rich_parts)
}

#[test]
fn apply_cell_patches_formula_edit_preserves_richdata_parts_and_metadata(
) -> Result<(), Box<dyn std::error::Error>> {
    let (bytes, rich_parts) = build_synthetic_richdata_calcchain_fixture();
    let pkg = XlsxPackage::from_bytes(&bytes)?;

    // Introduce a formula to trigger the default recalc policy (drop calcChain + rewrite workbook
    // rels and content types).
    let patch = PackageCellPatch::for_worksheet_part(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::Number(2.0),
        Some("=1+1".to_string()),
    );
    let out_bytes = pkg.apply_cell_patches_to_bytes(&[patch])?;

    let mut archive = zip::ZipArchive::new(Cursor::new(&out_bytes))?;

    // Existing behavior: calcChain is removed on formula edits.
    assert!(
        matches!(archive.by_name("xl/calcChain.xml").err(), Some(ZipError::FileNotFound)),
        "expected xl/calcChain.xml to be dropped after formula edits"
    );

    // Regression coverage: RichData parts must remain byte-for-byte.
    for (name, expected) in rich_parts {
        let mut buf = Vec::new();
        archive.by_name(name)?.read_to_end(&mut buf)?;
        assert_eq!(buf, expected, "expected {name} to be preserved byte-for-byte");
    }

    // Regression coverage: workbook.xml.rels must keep RichData relationships and only drop the
    // calcChain relationship.
    let mut workbook_rels = String::new();
    archive
        .by_name("xl/_rels/workbook.xml.rels")?
        .read_to_string(&mut workbook_rels)?;
    let rels_doc = roxmltree::Document::parse(&workbook_rels)?;
    let rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships";
    let relationships: Vec<(String, String)> = rels_doc
        .descendants()
        .filter(|n| n.is_element() && n.has_tag_name((rels_ns, "Relationship")))
        .map(|n| {
            (
                n.attribute("Type").unwrap_or_default().to_string(),
                n.attribute("Target").unwrap_or_default().to_string(),
            )
        })
        .collect();

    assert!(
        relationships
            .iter()
            .any(|(ty, target)| ty == "http://example.com/relationships/richData"
                && target == "richData/richValue.xml"),
        "expected workbook.xml.rels to preserve the RichData relationship"
    );
    assert!(
        !relationships
            .iter()
            .any(|(ty, target)| ty.ends_with("/calcChain") || target.ends_with("calcChain.xml")),
        "expected workbook.xml.rels to drop only the calcChain relationship"
    );
    assert_eq!(
        relationships.len(),
        3,
        "expected workbook.xml.rels to contain 3 relationships after dropping calcChain"
    );

    // Regression coverage: [Content_Types].xml must keep RichData overrides and only drop the
    // calcChain override.
    let mut content_types = String::new();
    archive
        .by_name("[Content_Types].xml")?
        .read_to_string(&mut content_types)?;
    let ct_doc = roxmltree::Document::parse(&content_types)?;
    let ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types";
    let overrides: Vec<String> = ct_doc
        .descendants()
        .filter(|n| n.is_element() && n.has_tag_name((ct_ns, "Override")))
        .filter_map(|n| n.attribute("PartName"))
        .map(|v| v.to_string())
        .collect();

    assert!(
        !overrides.iter().any(|p| p.ends_with("calcChain.xml")),
        "expected [Content_Types].xml to drop only the calcChain override"
    );
    for rich in [
        "/xl/richData/richValue.xml",
        "/xl/richData/richValueRel.xml",
        "/xl/richData/richValueTypes.xml",
        "/xl/richData/richValueStructure.xml",
    ] {
        assert!(
            overrides.iter().any(|p| p == rich),
            "expected [Content_Types].xml to preserve richData override {rich}"
        );
    }
    assert_eq!(
        overrides.len(),
        7,
        "expected [Content_Types].xml to contain 7 <Override> entries after dropping calcChain"
    );

    Ok(())
}

