use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn build_synthetic_xlsx_with_richdata_parts() -> (Vec<u8>, RichDataFixtures) {
    let root_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#;

    let workbook_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#;

    let workbook_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"#;

    let worksheet_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>
"#;

    let rich_value = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValue xmlns="http://example.com/richData"><v>alpha</v></richValue>
"#
    .to_vec();
    let rich_value_rel = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://example.com/richData"><v>beta</v></richValueRel>
"#
    .to_vec();
    let rich_value_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueTypes xmlns="http://example.com/richData"><v>gamma</v></richValueTypes>
"#
    .to_vec();
    let rich_value_structure = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueStructure xmlns="http://example.com/richData"><v>delta</v></richValueStructure>
"#
    .to_vec();

    let rich_value_rel_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://example.com/rel" Target="../dummy.xml"/>
</Relationships>
"#
    .to_vec();

    // Include RichData overrides. We intentionally omit the `styles.xml` override so the
    // `XlsxDocument` writer has to patch `[Content_Types].xml`, ensuring unknown overrides (our
    // RichData entries) survive those edits.
    let content_types_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/richData/richValue.xml" ContentType="application/xml"/>
  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/xml"/>
  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/xml"/>
  <Override PartName="/xl/richData/richValueStructure.xml" ContentType="application/xml"/>
</Types>
"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types_xml).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml).unwrap();

    zip.start_file("xl/richData/richValue.xml", options).unwrap();
    zip.write_all(&rich_value).unwrap();

    zip.start_file("xl/richData/richValueRel.xml", options).unwrap();
    zip.write_all(&rich_value_rel).unwrap();

    zip.start_file("xl/richData/richValueTypes.xml", options)
        .unwrap();
    zip.write_all(&rich_value_types).unwrap();

    zip.start_file("xl/richData/richValueStructure.xml", options)
        .unwrap();
    zip.write_all(&rich_value_structure).unwrap();

    // Optional relationship part: not currently interpreted by Formula but should be preserved by
    // virtue of the generic "preserve unknown parts" pipeline.
    zip.start_file(
        "xl/richData/_rels/richValueRel.xml.rels",
        options,
    )
    .unwrap();
    zip.write_all(&rich_value_rel_rels).unwrap();

    let bytes = zip.finish().unwrap().into_inner();
    (
        bytes,
        RichDataFixtures {
            rich_value,
            rich_value_rel,
            rich_value_types,
            rich_value_structure,
            rich_value_rel_rels,
        },
    )
}

#[derive(Debug)]
struct RichDataFixtures {
    rich_value: Vec<u8>,
    rich_value_rel: Vec<u8>,
    rich_value_types: Vec<u8>,
    rich_value_structure: Vec<u8>,
    rich_value_rel_rels: Vec<u8>,
}

fn read_zip_part(archive: &mut ZipArchive<Cursor<Vec<u8>>>, name: &str) -> Vec<u8> {
    let mut file = archive.by_name(name).unwrap_or_else(|err| {
        panic!("expected output zip to contain {name}, got err: {err}");
    });
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).unwrap();
    buf
}

#[test]
fn xlsx_document_roundtrip_preserves_richdata_parts_byte_for_byte() {
    let (input_xlsx, fixtures) = build_synthetic_xlsx_with_richdata_parts();

    let mut doc = formula_xlsx::load_from_bytes(&input_xlsx).expect("load_from_bytes should succeed");

    // Mutate a normal worksheet cell via the model to ensure we take the full XlsxDocument write
    // path, but do not touch any RichData parts (Formula doesn't interpret them yet).
    let sheet_id = doc.workbook.sheets[0].id;
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet must exist")
        .set_value(CellRef::from_a1("A1").unwrap(), CellValue::Number(2.0));
    // Also introduce a string cell to force sharedStrings generation, ensuring the save path
    // performs additional workbook.xml.rels + [Content_Types].xml edits while still preserving
    // RichData parts.
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet must exist")
        .set_value(
            CellRef::from_a1("B1").unwrap(),
            CellValue::String("World".to_string()),
        );

    let output_xlsx = doc.save_to_vec().expect("save_to_vec should succeed");

    let mut archive = ZipArchive::new(Cursor::new(output_xlsx)).expect("open output zip");

    // Assert parts still exist and are byte-for-byte preserved.
    assert_eq!(
        read_zip_part(&mut archive, "xl/richData/richValue.xml"),
        fixtures.rich_value
    );
    assert_eq!(
        read_zip_part(&mut archive, "xl/richData/richValueRel.xml"),
        fixtures.rich_value_rel
    );
    assert_eq!(
        read_zip_part(&mut archive, "xl/richData/richValueTypes.xml"),
        fixtures.rich_value_types
    );
    assert_eq!(
        read_zip_part(&mut archive, "xl/richData/richValueStructure.xml"),
        fixtures.rich_value_structure
    );
    assert_eq!(
        read_zip_part(&mut archive, "xl/richData/_rels/richValueRel.xml.rels"),
        fixtures.rich_value_rel_rels
    );

    // Ensure the RichData content type overrides remain present (even if `[Content_Types].xml` is
    // otherwise patched by the writer).
    let content_types = String::from_utf8(read_zip_part(&mut archive, "[Content_Types].xml"))
        .expect("[Content_Types].xml should be utf-8");
    let doc = roxmltree::Document::parse(&content_types).expect("parse [Content_Types].xml");
    let overrides: std::collections::HashSet<&str> = doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Override")
        .filter_map(|n| n.attribute("PartName"))
        .collect();
    for part in [
        "/xl/richData/richValue.xml",
        "/xl/richData/richValueRel.xml",
        "/xl/richData/richValueTypes.xml",
        "/xl/richData/richValueStructure.xml",
    ] {
        assert!(
            overrides.contains(part),
            "expected [Content_Types].xml to contain an Override for {part}, got:\n{content_types}"
        );
    }

    assert!(
        overrides.contains("/xl/sharedStrings.xml"),
        "expected writer to add sharedStrings override when writing string cell, got:\n{content_types}"
    );
}
