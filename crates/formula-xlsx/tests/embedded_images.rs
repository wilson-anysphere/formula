use base64::engine::general_purpose::STANDARD;
use base64::Engine as _;
use formula_model::CellRef;
use formula_xlsx::{extract_embedded_images, XlsxPackage};
use rust_xlsxwriter::Workbook;

const PNG_1X1_TRANSPARENT_B64: &str =
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMB/6XgdAAAAABJRU5ErkJggg==";

const METADATA_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk>
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
"#;

const METADATA_XML_RICH_VALUE_INDEX_ONE: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="1"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk>
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
"#;

const METADATA_XML_T_ZERO_BASED: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{00000000-0000-0000-0000-000000000000}">
          <xlrd:rvb i="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk>
      <rc t="0" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
"#;

const METADATA_XML_DIRECT_RICH_VALUE_INDEX: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <valueMetadata count="1">
    <bk>
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
"#;

const RDRICHVALUE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv s="0">
    <v>0</v>
    <v>6</v>
    <v>Example<![CDATA[Alt]]>Text</v>
  </rv>
</rvData>
"#;

const RDRICHVALUE_STRUCTURE_XML_REORDERED: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1">
  <s t="_localImage">
    <k n="CalcOrigin" t="i"/>
    <k n="_rvRel:LocalImageIdentifier" t="i"/>
    <k n="Text" t="s"/>
  </s>
</rvStructures>
"#;

const RDRICHVALUE_XML_REORDERED: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv s="0">
    <v>6</v>
    <v>0</v>
    <v>Reordered alt text</v>
  </rv>
</rvData>
"#;

const RICH_VALUE_XML_MINIMAL: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">0</v>
    </rv>
  </values>
</rvData>
"#;

const RICH_VALUE_XML_REL_INDEX_1: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0">
      <v kind="rel">1</v>
    </rv>
  </values>
</rvData>
"#;

const RICH_VALUE_REL_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.microsoft.com/office/2022/10/spreadsheetml/richvaluerelationships"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>
"#;

const RICH_VALUE_REL_XML_WITH_PLACEHOLDER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.microsoft.com/office/2022/10/spreadsheetml/richvaluerelationships"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel/>
  <rel r:id="rId1"/>
</richValueRel>
"#;

const RICH_VALUE_REL_XML_WITH_ID_WHITESPACE: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.microsoft.com/office/2022/10/spreadsheetml/richvaluerelationships"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id=" rId1 "/>
</richValueRel>
"#;

const RICH_VALUE_REL_XML_WITH_BOTH_ID_FORMS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.microsoft.com/office/2022/10/spreadsheetml/richvaluerelationships"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel id="wrong" r:id="rId1"/>
</richValueRel>
"#;

const RICH_VALUE_REL_XML_WITH_CUSTOM_REL_ID: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.microsoft.com/office/2022/10/spreadsheetml/richvaluerelationships"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rIdImg"/>
</richValueRel>
"#;

const METADATA_RELS_XML_CUSTOM_RICHVALUE_PARTS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.microsoft.com/office/2017/relationships/richValue"
                Target="richData/customRichValue.xml"/>
  <Relationship Id="rId2"
                Type="http://schemas.microsoft.com/office/2017/relationships/richValueRel"
                Target="richData/customRichValueRel.xml"/>
</Relationships>
"#;

const RICH_VALUE_REL_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                 Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                 Target="../media/image1.png"/>
</Relationships>
"#;

const RICH_VALUE_REL_RELS_XML_CUSTOM_REL_ID: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImg"
                 Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                 Target="../media/image1.png"/>
</Relationships>
"#;

const RICH_VALUE_REL_RELS_XML_MEDIA_RELATIVE_TO_XL: &str =
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                 Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                 Target="media/image1.png"/>
</Relationships>
"#;

const RICH_VALUE_REL_RELS_XML_MEDIA_RELATIVE_TO_XL_WITH_BACKSLASHES: &str =
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                 Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                 Target="media\image1.png"/>
</Relationships>
"#;

const RICH_VALUE_REL_RELS_XML_XL_MEDIA_NO_LEADING_SLASH: &str =
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                 Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                 Target="xl/media/image1.png"/>
</Relationships>
"#;

const RICH_VALUE_REL_RELS_XML_NON_IMAGE: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                Target="../media/image1.png"/>
</Relationships>
"#;

fn assert_extracts_embedded_image_from_cell_vm_metadata_richdata_schema_with_rels(
    metadata_xml: &str,
    vm: u32,
    rich_value_rel_xml: &str,
    rich_value_rel_rels_xml: &str,
) {
    let png_bytes = STANDARD
        .decode(PNG_1X1_TRANSPARENT_B64)
        .expect("decode png base64");

    // Create a minimal workbook with a concrete `B2` cell that we can attach `vm="1"` to.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_number(1, 1, 1.0).unwrap(); // B2
    let xlsx_bytes = workbook.save_to_buffer().unwrap();

    let mut pkg = XlsxPackage::from_bytes(&xlsx_bytes).unwrap();

    // Patch the worksheet cell to include the `vm` attribute expected by Excel's rich value schema.
    let sheet_part = "xl/worksheets/sheet1.xml";
    let mut sheet_xml = String::from_utf8(pkg.part(sheet_part).unwrap().to_vec()).unwrap();
    assert!(sheet_xml.contains(r#"r="B2""#), "expected B2 cell");
    let replacement = format!(r#"r="B2" vm="{vm}""#);
    sheet_xml = sheet_xml.replacen(r#"r="B2""#, &replacement, 1);
    pkg.set_part(sheet_part, sheet_xml.into_bytes());

    // Add the rich value parts + image payload.
    pkg.set_part("xl/metadata.xml", metadata_xml.as_bytes().to_vec());
    pkg.set_part(
        "xl/richData/rdrichvalue.xml",
        RDRICHVALUE_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/richValueRel.xml",
        rich_value_rel_xml.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/_rels/richValueRel.xml.rels",
        rich_value_rel_rels_xml.as_bytes().to_vec(),
    );
    pkg.set_part("xl/media/image1.png", png_bytes.clone());

    // Wire up the workbook relationships to the new parts (they are discovered via `workbook.xml.rels`).
    let rels_part = "xl/_rels/workbook.xml.rels";
    let mut rels_xml = String::from_utf8(pkg.part(rels_part).unwrap().to_vec()).unwrap();
    let insert_idx = rels_xml
        .rfind("</Relationships>")
        .expect("closing Relationships tag");
    rels_xml.insert_str(
        insert_idx,
        r#"
  <Relationship Id="rId1000" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>
  <Relationship Id="rId1001" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>
  <Relationship Id="rId1002" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>
"#,
    );
    pkg.set_part(rels_part, rels_xml.into_bytes());

    // Round-trip through ZIP writer to ensure the extractor works on a real package.
    let bytes = pkg.write_to_bytes().unwrap();
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    let images = extract_embedded_images(&pkg).unwrap();
    assert_eq!(images.len(), 1);

    let image = &images[0];
    assert_eq!(image.sheet_part, "xl/worksheets/sheet1.xml");
    assert_eq!(image.cell, CellRef::from_a1("B2").unwrap());
    assert_eq!(image.image_target, "xl/media/image1.png");
    assert_eq!(image.bytes, png_bytes);
    assert_eq!(image.alt_text.as_deref(), Some("ExampleAltText"));
    assert!(!image.decorative);
}

fn assert_extracts_embedded_image_from_cell_vm_metadata_richdata_schema(metadata_xml: &str, vm: u32) {
    assert_extracts_embedded_image_from_cell_vm_metadata_richdata_schema_with_rels(
        metadata_xml,
        vm,
        RICH_VALUE_REL_XML,
        RICH_VALUE_REL_RELS_XML,
    );
}

#[test]
fn extracts_embedded_image_from_cell_vm_metadata_richdata_schema() {
    assert_extracts_embedded_image_from_cell_vm_metadata_richdata_schema(METADATA_XML, 1);
}

#[test]
fn extracts_embedded_image_with_0_based_metadata_type_index() {
    // Some non-Excel producers have been observed to encode the `metadataTypes` index as 0-based
    // (`t="0"`) rather than Excel's typical 1-based (`t="1"`).
    assert_extracts_embedded_image_from_cell_vm_metadata_richdata_schema(METADATA_XML_T_ZERO_BASED, 1);
}

#[test]
fn extracts_embedded_image_when_metadata_uses_direct_rich_value_index() {
    // Some producers omit `<futureMetadata name="XLRICHVALUE">` and store the rich value record
    // index directly in `rc/@v`.
    assert_extracts_embedded_image_from_cell_vm_metadata_richdata_schema(
        METADATA_XML_DIRECT_RICH_VALUE_INDEX,
        1,
    );
}

#[test]
fn extracts_embedded_image_when_cell_vm_is_zero() {
    // Some workbooks have been observed to use 0-based `vm` indices.
    assert_extracts_embedded_image_from_cell_vm_metadata_richdata_schema(METADATA_XML, 0);
}

#[test]
fn extracts_when_rich_value_rel_rels_target_is_relative_to_xl() {
    // Some producers emit `Target="media/image1.png"` (relative to `xl/`) instead of
    // `Target="../media/image1.png"` (relative to `xl/richData/`).
    assert_extracts_embedded_image_from_cell_vm_metadata_richdata_schema_with_rels(
        METADATA_XML,
        1,
        RICH_VALUE_REL_XML,
        RICH_VALUE_REL_RELS_XML_MEDIA_RELATIVE_TO_XL,
    );
}

#[test]
fn extracts_when_rich_value_rel_rels_target_is_relative_to_xl_with_backslashes() {
    // Some producers emit `Target="media\\image1.png"` (relative to `xl/`) instead of
    // `Target="../media/image1.png"` (relative to `xl/richData/`).
    assert_extracts_embedded_image_from_cell_vm_metadata_richdata_schema_with_rels(
        METADATA_XML,
        1,
        RICH_VALUE_REL_XML,
        RICH_VALUE_REL_RELS_XML_MEDIA_RELATIVE_TO_XL_WITH_BACKSLASHES,
    );
}

#[test]
fn extracts_when_rich_value_rel_rels_target_starts_with_xl_prefix() {
    // Some producers emit `Target="xl/media/image1.png"` (missing the leading `/`).
    assert_extracts_embedded_image_from_cell_vm_metadata_richdata_schema_with_rels(
        METADATA_XML,
        1,
        RICH_VALUE_REL_XML,
        RICH_VALUE_REL_RELS_XML_XL_MEDIA_NO_LEADING_SLASH,
    );
}

#[test]
fn extracts_when_rich_value_rel_rid_has_whitespace() {
    // Some producers emit whitespace around the `r:id` value in richValueRel.xml.
    assert_extracts_embedded_image_from_cell_vm_metadata_richdata_schema_with_rels(
        METADATA_XML,
        1,
        RICH_VALUE_REL_XML_WITH_ID_WHITESPACE,
        RICH_VALUE_REL_RELS_XML,
    );
}

#[test]
fn extracts_when_rich_value_rel_has_both_prefixed_and_unprefixed_id_attributes() {
    // Some producers may include both `id` and `r:id`. Prefer the namespaced `r:id` value.
    assert_extracts_embedded_image_from_cell_vm_metadata_richdata_schema_with_rels(
        METADATA_XML,
        1,
        RICH_VALUE_REL_XML_WITH_BOTH_ID_FORMS,
        RICH_VALUE_REL_RELS_XML,
    );
}

#[test]
fn extracts_when_rich_value_rel_rid_is_non_numeric() {
    // While most workbooks use numeric relationship IDs (`rId1`, `rId2`, ...), OPC relationship IDs
    // are not required to be numeric. Ensure we still resolve targets when the ID has a non-numeric
    // suffix.
    assert_extracts_embedded_image_from_cell_vm_metadata_richdata_schema_with_rels(
        METADATA_XML,
        1,
        RICH_VALUE_REL_XML_WITH_CUSTOM_REL_ID,
        RICH_VALUE_REL_RELS_XML_CUSTOM_REL_ID,
    );
}

#[test]
fn extracts_when_richdata_parts_are_related_from_metadata_rels() {
    let png_bytes = STANDARD
        .decode(PNG_1X1_TRANSPARENT_B64)
        .expect("decode png base64");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_number(1, 1, 1.0).unwrap(); // B2
    let xlsx_bytes = workbook.save_to_buffer().unwrap();

    let mut pkg = XlsxPackage::from_bytes(&xlsx_bytes).unwrap();

    // Patch the worksheet cell to include the `vm` attribute expected by Excel's rich value schema.
    let sheet_part = "xl/worksheets/sheet1.xml";
    let mut sheet_xml = String::from_utf8(pkg.part(sheet_part).unwrap().to_vec()).unwrap();
    sheet_xml = sheet_xml.replacen(r#"r="B2""#, r#"r="B2" vm="1""#, 1);
    pkg.set_part(sheet_part, sheet_xml.into_bytes());

    // Use canonical `xl/metadata.xml` but store richData parts under non-canonical names.
    //
    // The extractor should discover these via `xl/_rels/metadata.xml.rels` (not via
    // workbook.xml.rels or canonical-path fallback).
    pkg.set_part("xl/metadata.xml", METADATA_XML.as_bytes().to_vec());
    pkg.set_part(
        "xl/_rels/metadata.xml.rels",
        METADATA_RELS_XML_CUSTOM_RICHVALUE_PARTS.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/customRichValue.xml",
        RICH_VALUE_XML_MINIMAL.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/customRichValueRel.xml",
        RICH_VALUE_REL_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/_rels/customRichValueRel.xml.rels",
        RICH_VALUE_REL_RELS_XML.as_bytes().to_vec(),
    );
    pkg.set_part("xl/media/image1.png", png_bytes.clone());

    // Round-trip through ZIP writer to ensure the extractor works on a real package.
    let bytes = pkg.write_to_bytes().unwrap();
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    let images = extract_embedded_images(&pkg).unwrap();
    assert_eq!(images.len(), 1);

    let image = &images[0];
    assert_eq!(image.sheet_part, "xl/worksheets/sheet1.xml");
    assert_eq!(image.cell, CellRef::from_a1("B2").unwrap());
    assert_eq!(image.image_target, "xl/media/image1.png");
    assert_eq!(image.bytes, png_bytes);
    assert_eq!(image.alt_text, None);
    assert!(!image.decorative);
}

#[test]
fn extracts_when_workbook_relationships_part_is_missing() {
    // Some synthetic/minimal packages may omit `xl/_rels/workbook.xml.rels`. The extractor should
    // still work via canonical-path fallbacks.
    let png_bytes = STANDARD
        .decode(PNG_1X1_TRANSPARENT_B64)
        .expect("decode png base64");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_number(1, 1, 1.0).unwrap(); // B2
    let xlsx_bytes = workbook.save_to_buffer().unwrap();

    let mut pkg = XlsxPackage::from_bytes(&xlsx_bytes).unwrap();

    let sheet_part = "xl/worksheets/sheet1.xml";
    let mut sheet_xml = String::from_utf8(pkg.part(sheet_part).unwrap().to_vec()).unwrap();
    sheet_xml = sheet_xml.replacen(r#"r="B2""#, r#"r="B2" vm="1""#, 1);
    pkg.set_part(sheet_part, sheet_xml.into_bytes());

    pkg.set_part("xl/metadata.xml", METADATA_XML.as_bytes().to_vec());
    pkg.set_part(
        "xl/richData/rdrichvalue.xml",
        RDRICHVALUE_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/richValueRel.xml",
        RICH_VALUE_REL_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/_rels/richValueRel.xml.rels",
        RICH_VALUE_REL_RELS_XML.as_bytes().to_vec(),
    );
    pkg.set_part("xl/media/image1.png", png_bytes.clone());

    // Remove the workbook relationships part entirely to force canonical fallbacks.
    pkg.parts_map_mut().remove("xl/_rels/workbook.xml.rels");

    let bytes = pkg.write_to_bytes().unwrap();
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    let images = extract_embedded_images(&pkg).unwrap();
    assert_eq!(images.len(), 1);
    assert_eq!(images[0].image_target, "xl/media/image1.png");
    assert_eq!(images[0].bytes, png_bytes);
    assert_eq!(images[0].alt_text.as_deref(), Some("ExampleAltText"));
}

#[test]
fn preserves_rich_value_rel_placeholders_to_avoid_index_shifts() {
    let png_bytes = STANDARD
        .decode(PNG_1X1_TRANSPARENT_B64)
        .expect("decode png base64");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_number(1, 1, 1.0).unwrap(); // B2
    let xlsx_bytes = workbook.save_to_buffer().unwrap();

    let mut pkg = XlsxPackage::from_bytes(&xlsx_bytes).unwrap();

    // Patch the worksheet cell to include the `vm` attribute expected by Excel's rich value schema.
    let sheet_part = "xl/worksheets/sheet1.xml";
    let mut sheet_xml = String::from_utf8(pkg.part(sheet_part).unwrap().to_vec()).unwrap();
    sheet_xml = sheet_xml.replacen(r#"r="B2""#, r#"r="B2" vm="1""#, 1);
    pkg.set_part(sheet_part, sheet_xml.into_bytes());

    // Force the extractor down the `richValue.xml` path (no rdrichvalue.xml). The rich value record
    // points at relationship index 1, which should resolve to the second `<rel>` entry.
    pkg.set_part("xl/metadata.xml", METADATA_XML.as_bytes().to_vec());
    pkg.set_part(
        "xl/richData/richValue.xml",
        RICH_VALUE_XML_REL_INDEX_1.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/richValueRel.xml",
        RICH_VALUE_REL_XML_WITH_PLACEHOLDER.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/_rels/richValueRel.xml.rels",
        RICH_VALUE_REL_RELS_XML.as_bytes().to_vec(),
    );
    pkg.set_part("xl/media/image1.png", png_bytes.clone());

    // Wire up workbook relationships for the parts we need.
    let rels_part = "xl/_rels/workbook.xml.rels";
    let mut rels_xml = String::from_utf8(pkg.part(rels_part).unwrap().to_vec()).unwrap();
    let insert_idx = rels_xml
        .rfind("</Relationships>")
        .expect("closing Relationships tag");
    rels_xml.insert_str(
        insert_idx,
        r#"
  <Relationship Id="rId1000" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>
  <Relationship Id="rId1001" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue" Target="richData/richValue.xml"/>
  <Relationship Id="rId1002" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>
"#,
    );
    pkg.set_part(rels_part, rels_xml.into_bytes());

    // Round-trip through ZIP writer to ensure the extractor works on a real package.
    let bytes = pkg.write_to_bytes().unwrap();
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    let images = extract_embedded_images(&pkg).unwrap();
    assert_eq!(images.len(), 1);
    assert_eq!(images[0].image_target, "xl/media/image1.png");
    assert_eq!(images[0].bytes, png_bytes);
}

#[test]
fn extracts_when_rich_value_relationship_index_is_one_based() {
    // Some producers have been observed to encode the relationship index in `richValue.xml` as
    // 1-based (so `1` refers to the first `<rel>` entry in `richValueRel.xml`).
    let png_bytes = STANDARD
        .decode(PNG_1X1_TRANSPARENT_B64)
        .expect("decode png base64");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_number(1, 1, 1.0).unwrap(); // B2
    let xlsx_bytes = workbook.save_to_buffer().unwrap();

    let mut pkg = XlsxPackage::from_bytes(&xlsx_bytes).unwrap();

    // Patch the worksheet cell to include the `vm` attribute expected by Excel's rich value schema.
    let sheet_part = "xl/worksheets/sheet1.xml";
    let mut sheet_xml = String::from_utf8(pkg.part(sheet_part).unwrap().to_vec()).unwrap();
    sheet_xml = sheet_xml.replacen(r#"r="B2""#, r#"r="B2" vm="1""#, 1);
    pkg.set_part(sheet_part, sheet_xml.into_bytes());

    // Force the extractor down the `richValue.xml` path (no rdrichvalue.xml).
    pkg.set_part("xl/metadata.xml", METADATA_XML.as_bytes().to_vec());
    pkg.set_part(
        "xl/richData/richValue.xml",
        RICH_VALUE_XML_REL_INDEX_1.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/richValueRel.xml",
        RICH_VALUE_REL_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/_rels/richValueRel.xml.rels",
        RICH_VALUE_REL_RELS_XML.as_bytes().to_vec(),
    );
    pkg.set_part("xl/media/image1.png", png_bytes.clone());

    // Wire up workbook relationships for the parts we need.
    let rels_part = "xl/_rels/workbook.xml.rels";
    let mut rels_xml = String::from_utf8(pkg.part(rels_part).unwrap().to_vec()).unwrap();
    let insert_idx = rels_xml
        .rfind("</Relationships>")
        .expect("closing Relationships tag");
    rels_xml.insert_str(
        insert_idx,
        r#"
  <Relationship Id="rId1000" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>
  <Relationship Id="rId1001" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue" Target="richData/richValue.xml"/>
  <Relationship Id="rId1002" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>
"#,
    );
    pkg.set_part(rels_part, rels_xml.into_bytes());

    // Round-trip through ZIP writer to ensure the extractor works on a real package.
    let bytes = pkg.write_to_bytes().unwrap();
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    let images = extract_embedded_images(&pkg).unwrap();
    assert_eq!(images.len(), 1);
    assert_eq!(images[0].image_target, "xl/media/image1.png");
    assert_eq!(images[0].bytes, png_bytes);
}

#[test]
fn extracts_from_multi_part_rich_value_tables() {
    // Some workbooks split the rich value table across multiple `richValue*.xml` parts. Ensure we
    // concatenate and index them correctly.
    let png_bytes = STANDARD
        .decode(PNG_1X1_TRANSPARENT_B64)
        .expect("decode png base64");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_number(1, 1, 1.0).unwrap(); // B2
    let xlsx_bytes = workbook.save_to_buffer().unwrap();

    let mut pkg = XlsxPackage::from_bytes(&xlsx_bytes).unwrap();

    let sheet_part = "xl/worksheets/sheet1.xml";
    let mut sheet_xml = String::from_utf8(pkg.part(sheet_part).unwrap().to_vec()).unwrap();
    sheet_xml = sheet_xml.replacen(r#"r="B2""#, r#"r="B2" vm="1""#, 1);
    pkg.set_part(sheet_part, sheet_xml.into_bytes());

    // vm=1 maps to richValue index 1 (second record globally).
    pkg.set_part(
        "xl/metadata.xml",
        METADATA_XML_RICH_VALUE_INDEX_ONE.as_bytes().to_vec(),
    );
    // Two rich value parts: richValue1.xml -> global idx 0, richValue2.xml -> global idx 1.
    pkg.set_part(
        "xl/richData/richValue1.xml",
        RICH_VALUE_XML_MINIMAL.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/richValue2.xml",
        RICH_VALUE_XML_MINIMAL.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/richValueRel.xml",
        RICH_VALUE_REL_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/_rels/richValueRel.xml.rels",
        RICH_VALUE_REL_RELS_XML.as_bytes().to_vec(),
    );
    pkg.set_part("xl/media/image1.png", png_bytes.clone());

    // Wire up workbook relationships. Point the `richValue` relationship at `richValue1.xml`, but
    // the extractor should still load and concatenate all `richValue*.xml` parts.
    let rels_part = "xl/_rels/workbook.xml.rels";
    let mut rels_xml = String::from_utf8(pkg.part(rels_part).unwrap().to_vec()).unwrap();
    let insert_idx = rels_xml
        .rfind("</Relationships>")
        .expect("closing Relationships tag");
    rels_xml.insert_str(
        insert_idx,
        r#"
  <Relationship Id="rId1000" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>
  <Relationship Id="rId1001" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue" Target="richData/richValue1.xml"/>
  <Relationship Id="rId1002" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>
"#,
    );
    pkg.set_part(rels_part, rels_xml.into_bytes());

    let bytes = pkg.write_to_bytes().unwrap();
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    let images = extract_embedded_images(&pkg).unwrap();
    assert_eq!(images.len(), 1);
    assert_eq!(images[0].image_target, "xl/media/image1.png");
    assert_eq!(images[0].bytes, png_bytes);
}

#[test]
fn skips_rich_value_relationships_that_are_not_images() {
    let png_bytes = STANDARD
        .decode(PNG_1X1_TRANSPARENT_B64)
        .expect("decode png base64");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_number(1, 1, 1.0).unwrap(); // B2
    let xlsx_bytes = workbook.save_to_buffer().unwrap();

    let mut pkg = XlsxPackage::from_bytes(&xlsx_bytes).unwrap();

    // Patch the worksheet cell to include the `vm` attribute expected by Excel's rich value schema.
    let sheet_part = "xl/worksheets/sheet1.xml";
    let mut sheet_xml = String::from_utf8(pkg.part(sheet_part).unwrap().to_vec()).unwrap();
    sheet_xml = sheet_xml.replacen(r#"r="B2""#, r#"r="B2" vm="1""#, 1);
    pkg.set_part(sheet_part, sheet_xml.into_bytes());

    // Add the rich value parts + image payload, but intentionally mark the relationship as a
    // *non-image* relationship type.
    pkg.set_part("xl/metadata.xml", METADATA_XML.as_bytes().to_vec());
    pkg.set_part(
        "xl/richData/rdrichvalue.xml",
        RDRICHVALUE_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/richValueRel.xml",
        RICH_VALUE_REL_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/_rels/richValueRel.xml.rels",
        RICH_VALUE_REL_RELS_XML_NON_IMAGE.as_bytes().to_vec(),
    );
    pkg.set_part("xl/media/image1.png", png_bytes);

    // Wire up the workbook relationships to the new parts (they are discovered via `workbook.xml.rels`).
    let rels_part = "xl/_rels/workbook.xml.rels";
    let mut rels_xml = String::from_utf8(pkg.part(rels_part).unwrap().to_vec()).unwrap();
    let insert_idx = rels_xml
        .rfind("</Relationships>")
        .expect("closing Relationships tag");
    rels_xml.insert_str(
        insert_idx,
        r#"
  <Relationship Id="rId1000" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>
  <Relationship Id="rId1001" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>
  <Relationship Id="rId1002" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>
"#,
    );
    pkg.set_part(rels_part, rels_xml.into_bytes());

    // Round-trip through ZIP writer to ensure the extractor works on a real package.
    let bytes = pkg.write_to_bytes().unwrap();
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    let images = extract_embedded_images(&pkg).unwrap();
    assert!(
        images.is_empty(),
        "expected extractor to ignore non-image relationship types"
    );
}

#[test]
fn uses_rdrichvaluestructure_to_locate_local_image_identifier_and_alt_text() {
    let png_bytes = STANDARD
        .decode(PNG_1X1_TRANSPARENT_B64)
        .expect("decode png base64");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_number(1, 1, 1.0).unwrap(); // B2
    let xlsx_bytes = workbook.save_to_buffer().unwrap();

    let mut pkg = XlsxPackage::from_bytes(&xlsx_bytes).unwrap();

    // Patch the worksheet cell to include the `vm` attribute expected by Excel's rich value schema.
    let sheet_part = "xl/worksheets/sheet1.xml";
    let mut sheet_xml = String::from_utf8(pkg.part(sheet_part).unwrap().to_vec()).unwrap();
    sheet_xml = sheet_xml.replacen(r#"r="B2""#, r#"r="B2" vm="1""#, 1);
    pkg.set_part(sheet_part, sheet_xml.into_bytes());

    // Add the rich value parts + image payload.
    pkg.set_part("xl/metadata.xml", METADATA_XML.as_bytes().to_vec());
    pkg.set_part(
        "xl/richData/rdrichvalue.xml",
        RDRICHVALUE_XML_REORDERED.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/rdrichvaluestructure.xml",
        RDRICHVALUE_STRUCTURE_XML_REORDERED.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/richValueRel.xml",
        RICH_VALUE_REL_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/_rels/richValueRel.xml.rels",
        RICH_VALUE_REL_RELS_XML.as_bytes().to_vec(),
    );
    pkg.set_part("xl/media/image1.png", png_bytes.clone());

    // Wire up the workbook relationships to the new parts (they are discovered via `workbook.xml.rels`).
    let rels_part = "xl/_rels/workbook.xml.rels";
    let mut rels_xml = String::from_utf8(pkg.part(rels_part).unwrap().to_vec()).unwrap();
    let insert_idx = rels_xml
        .rfind("</Relationships>")
        .expect("closing Relationships tag");
    rels_xml.insert_str(
        insert_idx,
        r#"
  <Relationship Id="rId1000" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>
  <Relationship Id="rId1001" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>
  <Relationship Id="rId1002" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="richData/rdrichvaluestructure.xml"/>
  <Relationship Id="rId1003" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>
"#,
    );
    pkg.set_part(rels_part, rels_xml.into_bytes());

    // Round-trip through ZIP writer to ensure the extractor works on a real package.
    let bytes = pkg.write_to_bytes().unwrap();
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    let images = extract_embedded_images(&pkg).unwrap();
    assert_eq!(images.len(), 1);
    assert_eq!(images[0].bytes, png_bytes);
    assert_eq!(images[0].alt_text.as_deref(), Some("Reordered alt text"));
}

#[test]
fn extracts_when_workbook_relationships_use_alternate_richdata_type_uris() {
    let png_bytes = STANDARD
        .decode(PNG_1X1_TRANSPARENT_B64)
        .expect("decode png base64");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_number(1, 1, 1.0).unwrap(); // B2
    let xlsx_bytes = workbook.save_to_buffer().unwrap();

    let mut pkg = XlsxPackage::from_bytes(&xlsx_bytes).unwrap();

    // Patch the worksheet cell to include the `vm` attribute expected by Excel's rich value schema.
    let sheet_part = "xl/worksheets/sheet1.xml";
    let mut sheet_xml = String::from_utf8(pkg.part(sheet_part).unwrap().to_vec()).unwrap();
    sheet_xml = sheet_xml.replacen(r#"r="B2""#, r#"r="B2" vm="1""#, 1);
    pkg.set_part(sheet_part, sheet_xml.into_bytes());

    // Store metadata/richdata parts under *non-canonical* names so the extractor must discover them
    // via workbook relationships (not via canonical-path fallback).
    pkg.set_part("xl/custom-metadata.xml", METADATA_XML.as_bytes().to_vec());
    pkg.set_part(
        "xl/richData/rdrichvalue.xml",
        RDRICHVALUE_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/customRichValueRel.xml",
        RICH_VALUE_REL_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/_rels/customRichValueRel.xml.rels",
        RICH_VALUE_REL_RELS_XML.as_bytes().to_vec(),
    );
    pkg.set_part("xl/media/image1.png", png_bytes.clone());

    // Wire up the workbook relationships to the new parts, but use the alternate relationship type
    // URIs (`metadata` + 2017 richValueRel).
    let rels_part = "xl/_rels/workbook.xml.rels";
    let mut rels_xml = String::from_utf8(pkg.part(rels_part).unwrap().to_vec()).unwrap();
    let insert_idx = rels_xml
        .rfind("</Relationships>")
        .expect("closing Relationships tag");
    rels_xml.insert_str(
        insert_idx,
        r#"
  <Relationship Id="rId1000" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="custom-metadata.xml"/>
  <Relationship Id="rId1001" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>
  <Relationship Id="rId1002" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel" Target="richData/customRichValueRel.xml"/>
"#,
    );
    pkg.set_part(rels_part, rels_xml.into_bytes());

    // Round-trip through ZIP writer to ensure the extractor works on a real package.
    let bytes = pkg.write_to_bytes().unwrap();
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    let images = extract_embedded_images(&pkg).unwrap();
    assert_eq!(images.len(), 1);
    assert_eq!(images[0].image_target, "xl/media/image1.png");
    assert_eq!(images[0].bytes, png_bytes);
}

#[test]
fn extracts_when_workbook_relationships_use_alternate_richdata_type_uris_and_target_is_relative_to_xl() {
    // Like `extracts_when_workbook_relationships_use_alternate_richdata_type_uris`, but uses
    // `Target="media/..."` in the richValueRel `.rels` part.
    let png_bytes = STANDARD
        .decode(PNG_1X1_TRANSPARENT_B64)
        .expect("decode png base64");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_number(1, 1, 1.0).unwrap(); // B2
    let xlsx_bytes = workbook.save_to_buffer().unwrap();

    let mut pkg = XlsxPackage::from_bytes(&xlsx_bytes).unwrap();

    // Patch the worksheet cell to include the `vm` attribute expected by Excel's rich value schema.
    let sheet_part = "xl/worksheets/sheet1.xml";
    let mut sheet_xml = String::from_utf8(pkg.part(sheet_part).unwrap().to_vec()).unwrap();
    sheet_xml = sheet_xml.replacen(r#"r="B2""#, r#"r="B2" vm="1""#, 1);
    pkg.set_part(sheet_part, sheet_xml.into_bytes());

    // Store metadata/richdata parts under *non-canonical* names so the extractor must discover them
    // via workbook relationships (not via canonical-path fallback).
    pkg.set_part("xl/custom-metadata.xml", METADATA_XML.as_bytes().to_vec());
    pkg.set_part(
        "xl/richData/rdrichvalue.xml",
        RDRICHVALUE_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/customRichValueRel.xml",
        RICH_VALUE_REL_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/_rels/customRichValueRel.xml.rels",
        RICH_VALUE_REL_RELS_XML_MEDIA_RELATIVE_TO_XL.as_bytes().to_vec(),
    );
    pkg.set_part("xl/media/image1.png", png_bytes.clone());

    // Wire up the workbook relationships to the new parts, but use the alternate relationship type
    // URIs (`metadata` + 2017 richValueRel).
    let rels_part = "xl/_rels/workbook.xml.rels";
    let mut rels_xml = String::from_utf8(pkg.part(rels_part).unwrap().to_vec()).unwrap();
    let insert_idx = rels_xml
        .rfind("</Relationships>")
        .expect("closing Relationships tag");
    rels_xml.insert_str(
        insert_idx,
        r#"
  <Relationship Id="rId1000" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="custom-metadata.xml"/>
  <Relationship Id="rId1001" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/>
  <Relationship Id="rId1002" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel" Target="richData/customRichValueRel.xml"/>
"#,
    );
    pkg.set_part(rels_part, rels_xml.into_bytes());

    // Round-trip through ZIP writer to ensure the extractor works on a real package.
    let bytes = pkg.write_to_bytes().unwrap();
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    let images = extract_embedded_images(&pkg).unwrap();
    assert_eq!(images.len(), 1);
    assert_eq!(images[0].image_target, "xl/media/image1.png");
    assert_eq!(images[0].bytes, png_bytes);
}

#[test]
fn extracts_when_workbook_relationships_use_2017_rdrichvalue_type_uris() {
    // Some workbooks use the 2017 relationship type URIs without the `/06` date segment.
    // Ensure we can still discover non-canonical rdRichValue + rdRichValueStructure part names.
    let png_bytes = STANDARD
        .decode(PNG_1X1_TRANSPARENT_B64)
        .expect("decode png base64");

    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.write_number(1, 1, 1.0).unwrap(); // B2
    let xlsx_bytes = workbook.save_to_buffer().unwrap();

    let mut pkg = XlsxPackage::from_bytes(&xlsx_bytes).unwrap();

    // Patch the worksheet cell to include the `vm` attribute expected by Excel's rich value schema.
    let sheet_part = "xl/worksheets/sheet1.xml";
    let mut sheet_xml = String::from_utf8(pkg.part(sheet_part).unwrap().to_vec()).unwrap();
    sheet_xml = sheet_xml.replacen(r#"r="B2""#, r#"r="B2" vm="1""#, 1);
    pkg.set_part(sheet_part, sheet_xml.into_bytes());

    pkg.set_part("xl/metadata.xml", METADATA_XML.as_bytes().to_vec());

    // Store rdRichValue + rdRichValueStructure under non-canonical names so the extractor must
    // discover them via workbook relationships (not via canonical-path fallback).
    pkg.set_part(
        "xl/richData/customRdrichvalue.xml",
        RDRICHVALUE_XML_REORDERED.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/customRdrichvaluestructure.xml",
        RDRICHVALUE_STRUCTURE_XML_REORDERED.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/richValueRel.xml",
        RICH_VALUE_REL_XML.as_bytes().to_vec(),
    );
    pkg.set_part(
        "xl/richData/_rels/richValueRel.xml.rels",
        RICH_VALUE_REL_RELS_XML.as_bytes().to_vec(),
    );
    pkg.set_part("xl/media/image1.png", png_bytes.clone());

    // Wire up the workbook relationships using the non-/06 relationship type URIs.
    let rels_part = "xl/_rels/workbook.xml.rels";
    let mut rels_xml = String::from_utf8(pkg.part(rels_part).unwrap().to_vec()).unwrap();
    let insert_idx = rels_xml
        .rfind("</Relationships>")
        .expect("closing Relationships tag");
    rels_xml.insert_str(
        insert_idx,
        r#"
  <Relationship Id="rId1000" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>
  <Relationship Id="rId1001" Type="http://schemas.microsoft.com/office/2017/relationships/rdRichValue" Target="richData/customRdrichvalue.xml"/>
  <Relationship Id="rId1002" Type="http://schemas.microsoft.com/office/2017/relationships/rdRichValueStructure" Target="richData/customRdrichvaluestructure.xml"/>
  <Relationship Id="rId1003" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>
"#,
    );
    pkg.set_part(rels_part, rels_xml.into_bytes());

    // Round-trip through ZIP writer to ensure the extractor works on a real package.
    let bytes = pkg.write_to_bytes().unwrap();
    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();

    let images = extract_embedded_images(&pkg).unwrap();
    assert_eq!(images.len(), 1);
    assert_eq!(images[0].image_target, "xl/media/image1.png");
    assert_eq!(images[0].bytes, png_bytes);
    assert_eq!(images[0].alt_text.as_deref(), Some("Reordered alt text"));
}
