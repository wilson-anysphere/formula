use std::io::{Cursor, Write};

use formula_model::CellRef;
use formula_xlsx::load_from_bytes;
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
fn xlsx_document_rich_value_index_supports_zero_based_vm_with_multiple_value_metadata_records() {
    // Like `embedded_cell_images_vm_zero_based_multi_record`, but exercises `load_from_bytes()` +
    // `XlsxDocument::rich_value_index()`.
    //
    // Some producers encode worksheet `c/@vm` as 0-based while `xl/metadata.xml` uses the typical
    // 1-based indexing for `<valueMetadata>` `<bk>` blocks.

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
</Relationships>"#;

    // Two cells with 0-based vm values.
    let sheet1_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"><v>0</v></c>
    </row>
    <row r="2">
      <c r="A2" vm="1"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    // metadata.xml maps valueMetadata bk[0] -> richValue index 0, and bk[1] -> richValue index 1.
    let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="2">
    <bk><extLst><ext uri="{00000000-0000-0000-0000-000000000000}"><xlrd:rvb i="0"/></ext></extLst></bk>
    <bk><extLst><ext uri="{00000000-0000-0000-0000-000000000000}"><xlrd:rvb i="1"/></ext></extLst></bk>
  </futureMetadata>
  <valueMetadata count="2">
    <bk><rc t="1" v="0"/></bk>
    <bk><rc t="1" v="1"/></bk>
  </valueMetadata>
</metadata>"#;

    let bytes = build_package(&[
        ("xl/workbook.xml", workbook_xml),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/worksheets/sheet1.xml", sheet1_xml),
        ("xl/metadata.xml", metadata_xml),
    ]);

    let doc = load_from_bytes(&bytes).expect("load xlsx document");
    let sheet_id = doc.workbook.sheets[0].id;
    assert_eq!(
        doc.rich_value_index(sheet_id, CellRef::from_a1("A1").unwrap()),
        Some(0)
    );
    assert_eq!(
        doc.rich_value_index(sheet_id, CellRef::from_a1("A2").unwrap()),
        Some(1)
    );
}

