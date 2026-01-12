use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

fn zip_selected_parts(zip_bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut out = BTreeMap::new();

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip entry");
        if file.is_dir() {
            continue;
        }
        let name = file.name().to_string();
        let keep = name == "xl/metadata.xml"
            || name == "xl/media/image1.png"
            || (name.starts_with("xl/richData/")
                && !name.starts_with("xl/richData/_rels/")
                && name.ends_with(".xml"))
            || (name.starts_with("xl/richData/_rels/") && name.ends_with(".rels"));
        if !keep {
            continue;
        }

        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read part");
        out.insert(name, buf);
    }

    out
}

fn build_fixture_xlsx_with_metadata_and_richdata() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/xml"/>
  <Override PartName="/xl/richData/richValue.xml" ContentType="application/xml"/>
  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/xml"/>
</Types>
"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"#;

    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1">
      <c r="A1" cm="7" vm="9" customAttr="x">
        <v>1</v>
        <extLst>
          <ext uri="{123}">
            <test xmlns="http://example.com">images-in-cells</test>
          </ext>
        </extLst>
      </c>
    </row>
  </sheetData>
</worksheet>
"#;

    let metadata_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="2">
    <metadataType name="SOMEOTHERTYPE"/>
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
    <bk><rc t="2" v="0"/></bk>
  </valueMetadata>
  <extLst>
    <ext uri="{DEADBEEF-DEAD-BEEF-DEAD-BEEFDEADBEEF}">
      <custom xmlns="urn:example:metadata">METADATA-UNIQUE</custom>
    </ext>
  </extLst>
</metadata>
"#;

    let rich_value_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xlrd:rvData xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <xlrd:rv id="0">
    <xlrd:v>RICHVALUE-UNIQUE</xlrd:v>
  </xlrd:rv>
</xlrd:rvData>
"#;

    let rich_value_rel_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xlrd:rvRel xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xlrd:rel r:id="rId1"/>
  <xlrd:note>RICHVALUEREL-UNIQUE</xlrd:note>
</xlrd:rvRel>
"#;

    let rich_value_rel_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
"#;

    // 1x1 transparent PNG.
    // Embedded as raw bytes so round-tripping can be validated byte-for-byte.
    const IMAGE1_PNG: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
        0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
        0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78,
        0x9C, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options)
        .expect("zip file");
    zip.write_all(content_types.as_bytes()).expect("zip write");

    zip.start_file("_rels/.rels", options).expect("zip file");
    zip.write_all(root_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/workbook.xml", options).expect("zip file");
    zip.write_all(workbook_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("zip file");
    zip.write_all(workbook_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("zip file");
    zip.write_all(sheet1_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/metadata.xml", options).expect("zip file");
    zip.write_all(metadata_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/richData/richValue.xml", options)
        .expect("zip file");
    zip.write_all(rich_value_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/richData/richValueRel.xml", options)
        .expect("zip file");
    zip.write_all(rich_value_rel_xml.as_bytes())
        .expect("zip write");

    zip.start_file("xl/richData/_rels/richValueRel.xml.rels", options)
        .expect("zip file");
    zip.write_all(rich_value_rel_rels.as_bytes())
        .expect("zip write");

    zip.start_file("xl/media/image1.png", options)
        .expect("zip file");
    zip.write_all(IMAGE1_PNG).expect("zip write");

    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn preserves_metadata_and_richdata_parts_through_document_roundtrip() {
    let fixture = build_fixture_xlsx_with_metadata_and_richdata();

    let original_parts = zip_selected_parts(&fixture);
    assert!(
        original_parts.contains_key("xl/metadata.xml"),
        "fixture must include xl/metadata.xml"
    );
    assert!(
        original_parts.contains_key("xl/richData/richValue.xml"),
        "fixture must include xl/richData/richValue.xml"
    );
    assert!(
        original_parts.contains_key("xl/richData/_rels/richValueRel.xml.rels"),
        "fixture must include xl/richData/_rels/richValueRel.xml.rels"
    );
    assert!(
        original_parts.contains_key("xl/media/image1.png"),
        "fixture must include xl/media/image1.png"
    );

    let mut doc = load_from_bytes(&fixture).expect("load fixture");
    let sheet_id = doc.workbook.sheet_by_name("Sheet1").expect("Sheet1").id;
    assert!(
        doc.set_cell_value(
            sheet_id,
            CellRef::from_a1("A1").unwrap(),
            CellValue::Number(2.0)
        ),
        "expected set_cell_value to succeed"
    );

    let saved = doc.save_to_vec().expect("save");
    let saved_parts = zip_selected_parts(&saved);

    assert_eq!(
        saved_parts, original_parts,
        "expected metadata/richData/media parts to be preserved byte-for-byte"
    );
}

