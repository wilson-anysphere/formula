use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{patch_xlsx_streaming_workbook_cell_patches, CellPatch, WorkbookCellPatches};
use zip::result::ZipError;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

const METADATA_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <customData marker="METADATA_MARKER_1"/>
</metadata>"#;

const METADATA_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2020/12/relationships/richValueTypes" Target="richData/richValueTypes.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2020/12/relationships/richValues" Target="richData/richValues.xml"/>
</Relationships>"#;

const RICH_VALUE_TYPES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypes xmlns="http://schemas.microsoft.com/office/spreadsheetml/2020/richdata">
  <rvType id="1" marker="RVT_MARKER_1"/>
</rvTypes>"#;

const RICH_VALUES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2020/richdata">
  <rv marker="RV_MARKER_1"/>
</rvData>"#;

fn build_richdata_calcchain_fixture_xlsx() -> Vec<u8> {
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
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
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/vnd.ms-excel.richValueTypes+xml"/>
  <Override PartName="/xl/richData/richValues.xml" ContentType="application/vnd.ms-excel.richValues+xml"/>
</Types>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let calc_chain_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/calcChain.xml", options).unwrap();
    zip.write_all(calc_chain_xml.as_bytes()).unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(METADATA_XML.as_bytes()).unwrap();

    zip.start_file("xl/_rels/metadata.xml.rels", options).unwrap();
    zip.write_all(METADATA_RELS_XML.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValueTypes.xml", options)
        .unwrap();
    zip.write_all(RICH_VALUE_TYPES_XML.as_bytes()).unwrap();

    zip.start_file("xl/richData/richValues.xml", options)
        .unwrap();
    zip.write_all(RICH_VALUES_XML.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn streaming_recalc_policy_preserves_richdata_parts_and_relationships(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_richdata_calcchain_fixture_xlsx();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellPatch::set_value_with_formula(CellValue::Number(2.0), "=1+1"),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(bytes), &mut out, &patches)?;

    let mut archive = ZipArchive::new(Cursor::new(out.get_ref()))?;

    assert!(
        matches!(
            archive.by_name("xl/calcChain.xml").err(),
            Some(ZipError::FileNotFound)
        ),
        "expected streaming patcher to drop xl/calcChain.xml after formula edits"
    );

    let mut metadata = Vec::new();
    archive.by_name("xl/metadata.xml")?.read_to_end(&mut metadata)?;
    assert_eq!(
        metadata,
        METADATA_XML.as_bytes(),
        "xl/metadata.xml must be preserved byte-for-byte"
    );

    let mut metadata_rels = Vec::new();
    archive
        .by_name("xl/_rels/metadata.xml.rels")?
        .read_to_end(&mut metadata_rels)?;
    assert_eq!(
        metadata_rels,
        METADATA_RELS_XML.as_bytes(),
        "xl/_rels/metadata.xml.rels must be preserved byte-for-byte"
    );

    let mut rich_value_types = Vec::new();
    archive
        .by_name("xl/richData/richValueTypes.xml")?
        .read_to_end(&mut rich_value_types)?;
    assert_eq!(
        rich_value_types,
        RICH_VALUE_TYPES_XML.as_bytes(),
        "xl/richData/richValueTypes.xml must be preserved byte-for-byte"
    );

    let mut rich_values = Vec::new();
    archive
        .by_name("xl/richData/richValues.xml")?
        .read_to_end(&mut rich_values)?;
    assert_eq!(
        rich_values,
        RICH_VALUES_XML.as_bytes(),
        "xl/richData/richValues.xml must be preserved byte-for-byte"
    );

    let mut workbook_rels = String::new();
    archive
        .by_name("xl/_rels/workbook.xml.rels")?
        .read_to_string(&mut workbook_rels)?;
    assert!(
        !workbook_rels.contains("relationships/calcChain") && !workbook_rels.contains("calcChain"),
        "xl/_rels/workbook.xml.rels must not reference calcChain after formula edits"
    );
    assert!(
        workbook_rels.contains(r#"Target="metadata.xml""#),
        "xl/_rels/workbook.xml.rels must retain the metadata relationship"
    );
    assert!(
        workbook_rels.contains(r#"Id="rId9""#),
        "xl/_rels/workbook.xml.rels should preserve the original relationship Id for metadata"
    );

    let mut content_types = String::new();
    archive
        .by_name("[Content_Types].xml")?
        .read_to_string(&mut content_types)?;
    assert!(
        !content_types.contains("/xl/calcChain.xml"),
        "[Content_Types].xml must remove the calcChain override after formula edits"
    );
    assert!(
        content_types.contains("/xl/metadata.xml"),
        "[Content_Types].xml must retain the metadata override"
    );
    assert!(
        content_types.contains("/xl/richData/richValueTypes.xml"),
        "[Content_Types].xml must retain the richValueTypes override"
    );
    assert!(
        content_types.contains("/xl/richData/richValues.xml"),
        "[Content_Types].xml must retain the richValues override"
    );

    Ok(())
}

