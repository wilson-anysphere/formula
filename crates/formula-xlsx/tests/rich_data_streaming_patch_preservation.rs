use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{CellPatch, WorkbookCellPatches, XlsxPackage};

fn build_rich_data_fixture_xlsx() -> Vec<u8> {
    // This is intentionally "minimal but plausible" rather than a complete Excel workbook.
    // The regression we're guarding against is the patch pipeline dropping or corrupting
    // richData parts (used by images-in-cell) when applying unrelated cell edits.

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

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
</Relationships>"#;

    // A1 is meant to stand in for an "image in cell" (vm=... points at cell metadata in
    // xl/metadata.xml which in turn points at richData parts). We patch B1 in the test.
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1">
      <c r="A1" vm="1"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <extLst>
    <ext uri="{F2E3A63F-0A32-4A6B-8C25-2C079F5B8B1B}">
      <test xmlns="http://example.com/richData">metadata</test>
    </ext>
  </extLst>
</metadata>
"#;

    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<test xmlns="http://example.com/richData">richValueRel</test>
"#;

    let rich_value_rel_rels_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
"#;

    let rich_value_types_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<test xmlns="http://example.com/richData">richValueTypes</test>
"#;

    let image_bytes: &[u8] = b"\x89PNG\r\n\x1a\nfake-png-payload";

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    // Rich-data (images in cell) support parts.
    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml).unwrap();

    zip.start_file("xl/richData/richValueRel.xml", options)
        .unwrap();
    zip.write_all(rich_value_rel_xml).unwrap();

    zip.start_file("xl/richData/_rels/richValueRel.xml.rels", options)
        .unwrap();
    zip.write_all(rich_value_rel_rels_xml).unwrap();

    zip.start_file("xl/richData/richValueTypes.xml", options)
        .unwrap();
    zip.write_all(rich_value_types_xml).unwrap();

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(image_bytes).unwrap();

    zip.finish().unwrap().into_inner()
}

fn read_zip_parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut parts = BTreeMap::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).expect("zip entry");
        if !file.is_file() {
            continue;
        }
        // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
        // advertise enormous uncompressed sizes (zip-bomb style OOM).
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).expect("read zip entry");
        parts.insert(file.name().to_string(), buf);
    }
    parts
}

#[test]
fn apply_cell_patches_preserves_rich_data_parts_byte_for_byte() {
    let original_bytes = build_rich_data_fixture_xlsx();
    let original_parts = read_zip_parts(&original_bytes);

    let mut pkg = XlsxPackage::from_bytes(&original_bytes).expect("read pkg");

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("B1").unwrap(),
        CellPatch::set_value(CellValue::Number(1.0)),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let updated_bytes = pkg.write_to_bytes().expect("write xlsx");
    let updated_parts = read_zip_parts(&updated_bytes);

    // Sanity check: the edit should have changed the target worksheet.
    let original_sheet = original_parts
        .get("xl/worksheets/sheet1.xml")
        .expect("original sheet1.xml");
    let updated_sheet = updated_parts
        .get("xl/worksheets/sheet1.xml")
        .expect("updated sheet1.xml");
    assert_ne!(
        original_sheet, updated_sheet,
        "expected the worksheet part to change after patching"
    );
    let updated_sheet_str = std::str::from_utf8(updated_sheet).expect("updated sheet utf-8");
    assert!(
        updated_sheet_str.contains(r#"r="B1""#) && updated_sheet_str.contains("<v>1</v>"),
        "expected patched B1 cell value in worksheet xml, got: {updated_sheet_str}"
    );

    // All input parts should still be present in the output (no stripping of unknown parts).
    assert_eq!(
        original_parts.keys().collect::<Vec<_>>(),
        updated_parts.keys().collect::<Vec<_>>(),
        "expected output zip part set to match the original (only worksheet bytes should differ)"
    );

    // RichData (images-in-cell) parts must survive byte-for-byte.
    for rich_part in [
        "xl/metadata.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
        "xl/richData/richValueTypes.xml",
        "xl/media/image1.png",
    ] {
        let before = original_parts.get(rich_part).unwrap_or_else(|| {
            panic!("expected {rich_part} to exist in original zip fixture")
        });
        let after = updated_parts
            .get(rich_part)
            .unwrap_or_else(|| panic!("expected {rich_part} to exist in patched zip"));
        assert_eq!(
            before, after,
            "expected {rich_part} to be preserved byte-for-byte"
        );
    }

    // Additionally, everything except the edited worksheet should be preserved verbatim.
    for (name, original) in original_parts {
        if name == "xl/worksheets/sheet1.xml" {
            continue;
        }
        let updated = updated_parts
            .get(&name)
            .unwrap_or_else(|| panic!("expected {name} to exist in patched zip"));
        assert_eq!(
            original, *updated,
            "expected unchanged part {name} to be preserved byte-for-byte"
        );
    }
}
