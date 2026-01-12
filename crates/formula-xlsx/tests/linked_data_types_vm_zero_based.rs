use std::io::{Cursor, Write};

use formula_model::CellRef;
use formula_xlsx::rich_data::extract_linked_data_types;
use formula_xlsx::XlsxPackage;
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_package(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn linked_data_types_support_zero_based_vm_indices() -> Result<(), Box<dyn std::error::Error>> {
    // Simulate a workbook where worksheet `c/@vm` is 0-based (cells use vm="0" and vm="1"), while
    // `xl/metadata.xml` uses the typical 1-based indexing for `<valueMetadata>` `<bk>` blocks.
    //
    // This is analogous to the embedded-images-in-cells `vm` edge case: if callers assume `vm` is
    // always 1-based, they will fail to resolve the first record and mis-map subsequent records.

    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let sheet1_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"/>
    </row>
    <row r="2">
      <c r="A2" vm="1"/>
    </row>
  </sheetData>
</worksheet>"#;

    // Canonical 1-based valueMetadata bk indices:
    // vm=1 -> richValue index 0, vm=2 -> richValue index 1.
    let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <valueMetadata count="2">
    <bk><rc t="1" v="0"/></bk>
    <bk><rc t="1" v="1"/></bk>
  </valueMetadata>
</metadata>"#;

    let rich_value_types_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypes xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <types>
    <type id="0" name="com.microsoft.excel.stocks" structure="s_stock"/>
    <type id="1" name="com.microsoft.excel.geography" structure="s_geo"/>
  </types>
</rvTypes>"#;

    let rich_value_structure_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStruct xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <structures>
    <structure id="s_stock">
      <member name="display" kind="s"/>
    </structure>
    <structure id="s_geo">
      <member name="display" kind="s"/>
    </structure>
  </structures>
</rvStruct>"#;

    let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <rv type="0"><v>MSFT</v></rv>
    <rv type="1"><v>Seattle</v></rv>
  </values>
</rvData>"#;

    let bytes = build_package(&[
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet1_xml),
        ("xl/metadata.xml", metadata_xml),
        ("xl/richData/richValueTypes.xml", rich_value_types_xml),
        ("xl/richData/richValueStructure.xml", rich_value_structure_xml),
        ("xl/richData/richValue.xml", rich_value_xml),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes)?;
    let extracted = extract_linked_data_types(&pkg)?;

    let a1 = extracted
        .get(&("Sheet1".to_string(), CellRef::from_a1("A1")?))
        .ok_or("missing Sheet1!A1 rich value")?;
    assert_eq!(a1.type_name.as_deref(), Some("com.microsoft.excel.stocks"));
    assert_eq!(a1.display.as_deref(), Some("MSFT"));

    let a2 = extracted
        .get(&("Sheet1".to_string(), CellRef::from_a1("A2")?))
        .ok_or("missing Sheet1!A2 rich value")?;
    assert_eq!(a2.type_name.as_deref(), Some("com.microsoft.excel.geography"));
    assert_eq!(a2.display.as_deref(), Some("Seattle"));

    Ok(())
}

