use std::io::{Cursor, Read, Write};

use formula_model::CellRef;
use formula_xlsx::load_from_bytes;
use roxmltree::Document;
use zip::ZipArchive;

fn build_richdata_recalc_fixture(
    metadata_xml: &[u8],
    metadata_rels_xml: &[u8],
    rich_value_types_xml: &[u8],
    rich_values_xml: &[u8],
) -> Vec<u8> {
    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>"#;

    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/vnd.ms-excel.richtypes+xml"/>
  <Override PartName="/xl/richData/richValues.xml" ContentType="application/vnd.ms-excel.richvalues+xml"/>
</Types>"#;

    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    // Required core workbook parts.
    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml).unwrap();

    // Recalc-policy target part (should be removed on formula edit).
    zip.start_file("xl/calcChain.xml", options).unwrap();
    zip.write_all(b"calc chain bytes").unwrap();

    // Linked-data-type / rich-data infrastructure that must survive recalc-policy patching.
    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml).unwrap();

    zip.start_file("xl/_rels/metadata.xml.rels", options)
        .unwrap();
    zip.write_all(metadata_rels_xml).unwrap();

    zip.start_file("xl/richData/richValueTypes.xml", options)
        .unwrap();
    zip.write_all(rich_value_types_xml).unwrap();

    zip.start_file("xl/richData/richValues.xml", options)
        .unwrap();
    zip.write_all(rich_values_xml).unwrap();

    zip.finish().unwrap().into_inner()
}

fn read_zip_part<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut file = archive.by_name(name)?;
    let mut buf = Vec::new();
    file.read_to_end(&mut buf)?;
    Ok(buf)
}

#[test]
fn formula_edit_recalc_policy_preserves_linked_data_type_parts(
) -> Result<(), Box<dyn std::error::Error>> {
    const METADATA_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="0"/>
</metadata>"#;

    const METADATA_RELS_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://example.com/richValueTypes" Target="richData/richValueTypes.xml"/>
  <Relationship Id="rId2" Type="http://example.com/richValues" Target="richData/richValues.xml"/>
</Relationships>"#;

    const RICH_VALUE_TYPES_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypes xmlns="http://example.com/richData">stable richValueTypes bytes</rvTypes>"#;

    const RICH_VALUES_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValues xmlns="http://example.com/richData">stable richValues bytes</richValues>"#;

    let bytes = build_richdata_recalc_fixture(
        METADATA_XML,
        METADATA_RELS_XML,
        RICH_VALUE_TYPES_XML,
        RICH_VALUES_XML,
    );

    let mut doc = load_from_bytes(&bytes)?;
    let sheet_id = doc
        .workbook
        .sheets
        .first()
        .expect("fixture should have one sheet")
        .id;

    assert!(
        doc.set_cell_formula(sheet_id, CellRef::from_a1("A1")?, Some("=2+2".to_string())),
        "expected formula edit to succeed"
    );

    let saved = doc.save_to_vec()?;
    let mut archive = ZipArchive::new(Cursor::new(&saved))?;

    assert!(
        archive.by_name("xl/calcChain.xml").is_err(),
        "expected calcChain.xml to be removed after formula edit"
    );

    let workbook_rels_xml =
        String::from_utf8(read_zip_part(&mut archive, "xl/_rels/workbook.xml.rels")?)?;
    let rels_doc = Document::parse(&workbook_rels_xml)?;
    let rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships";
    let relationships: Vec<(String, String, String)> = rels_doc
        .descendants()
        .filter(|n| n.is_element() && n.has_tag_name((rels_ns, "Relationship")))
        .map(|n| {
            (
                n.attribute("Id").unwrap_or_default().to_string(),
                n.attribute("Type").unwrap_or_default().to_string(),
                n.attribute("Target").unwrap_or_default().to_string(),
            )
        })
        .collect();

    assert!(
        !relationships.iter().any(
            |(_id, ty, target)| ty.ends_with("/calcChain") || target.ends_with("calcChain.xml")
        ),
        "expected workbook.xml.rels to stop referencing calcChain.xml, got: {workbook_rels_xml}"
    );
    assert!(
        relationships
            .iter()
            .any(|(id, _ty, target)| id == "rId9" && target == "metadata.xml"),
        "expected workbook.xml.rels to preserve metadata relationship rId9 -> metadata.xml, got: {workbook_rels_xml}"
    );

    let content_types_xml = String::from_utf8(read_zip_part(&mut archive, "[Content_Types].xml")?)?;
    let ct_doc = Document::parse(&content_types_xml)?;
    let ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types";
    let overrides: Vec<String> = ct_doc
        .descendants()
        .filter(|n| n.is_element() && n.has_tag_name((ct_ns, "Override")))
        .filter_map(|n| n.attribute("PartName"))
        .map(|v| v.to_string())
        .collect();

    assert!(
        !overrides.iter().any(|p| p.ends_with("calcChain.xml")),
        "expected [Content_Types].xml override for calcChain.xml to be removed, got: {content_types_xml}"
    );
    for required in [
        "/xl/metadata.xml",
        "/xl/richData/richValueTypes.xml",
        "/xl/richData/richValues.xml",
    ] {
        assert!(
            overrides.iter().any(|p| p == required),
            "expected [Content_Types].xml to preserve {required}, got: {content_types_xml}"
        );
    }

    assert_eq!(
        read_zip_part(&mut archive, "xl/metadata.xml")?,
        METADATA_XML,
        "xl/metadata.xml should be preserved byte-for-byte"
    );
    assert_eq!(
        read_zip_part(&mut archive, "xl/_rels/metadata.xml.rels")?,
        METADATA_RELS_XML,
        "xl/_rels/metadata.xml.rels should be preserved byte-for-byte"
    );
    assert_eq!(
        read_zip_part(&mut archive, "xl/richData/richValueTypes.xml")?,
        RICH_VALUE_TYPES_XML,
        "xl/richData/richValueTypes.xml should be preserved byte-for-byte"
    );
    assert_eq!(
        read_zip_part(&mut archive, "xl/richData/richValues.xml")?,
        RICH_VALUES_XML,
        "xl/richData/richValues.xml should be preserved byte-for-byte"
    );

    Ok(())
}
